Attribute VB_Name = "EventosDS"
Option Explicit

Public Const MAX_EVENT_SIMULTANEO As Byte = 5
Public Const MAX_USERS_EVENT As Byte = 64
Public Const MAX_MAP_FIGHT As Byte = 4
Public Const MAP_TILE_VS As Byte = 17

Public Enum eModalityEvent
    CastleMode = 1
    DagaRusa = 2
    DeathMatch = 3
    Aracnus = 4
    HombreLobo = 5
    Minotauro = 6
    Busqueda = 7
    Unstoppable = 8
    Invasion = 9
    Enfrentamientos = 10
End Enum

Private Type tUserEvent
    Id As Integer
    Team As Byte
    Value As Integer
    Selected As Byte
    MapFight As Byte
End Type

Public Enum eFaction
    fCrim = 0
    fCiu = 1
    fLegion = 2
    fArmada = 3
End Enum

Private Type tEvents
    Enabled As Boolean
    Run As Boolean
    Modality As eModalityEvent
    TeamCant As Byte
    
    Quotas As Byte
    Inscribed As Byte
    
    LvlMax As Byte
    LvlMin As Byte
    
    GldInscription As Long
    DspInscription As Long
    
    AllowedClasses() As Byte
    AllowedFaction() As eFaction
    
    PrizeAccumulated As Boolean
    PrizeDsp As Integer
    PrizeGld As Long
    PrizeObj As Obj
    
    LimitRed As Integer
    
    ValidItem As Boolean
    WinFollow As Boolean
                      
    TimeInit As Long
    TimeCancel As Long
    TimeCount As Long
    TimeFinish As Long
    
    Users() As tUserEvent
    
    ' Por si alguno es con NPC
    NpcIndex As Integer
    
    ' Por si cambia el body del personaje y saca todo lo otro.
    CharBody As Integer
    CharHp As Integer
    
    npcUserIndex As Integer
End Type

Public Events(1 To MAX_EVENT_SIMULTANEO) As tEvents

Private Type tMap
    Run As Boolean
    map As Integer
    X As Byte
    Y As Byte
End Type

Private Type tMapEvent
    Fight(1 To MAX_MAP_FIGHT) As tMap
End Type

Private MapEvent As tMapEvent



Public Sub LoadMapEvent()
10        With MapEvent
20            .Fight(1).Run = False
30            .Fight(1).map = 217
40            .Fight(1).X = 16 '+17
50            .Fight(1).Y = 12 '+17
              
60            .Fight(2).Run = False
70            .Fight(2).map = 217
80            .Fight(2).X = 16 '+17
90            .Fight(2).Y = 41 '+17
100           .Fight(3).Run = False
110           .Fight(3).map = 217
120           .Fight(3).X = 16 '+17
130           .Fight(3).Y = 68 '+17
              
140           .Fight(4).Run = False
150           .Fight(4).map = 217
160           .Fight(4).X = 46 '+17
170           .Fight(4).Y = 12 '+17
          
          
          
180       End With
End Sub
'/MANEJO DE LOS TIEMPOS '/
Public Sub LoopEvent()
10    On Error GoTo error
          Dim LoopC As Long
          Dim LoopY As Integer
          
20        For LoopC = 1 To MAX_EVENT_SIMULTANEO
30            With Events(LoopC)
40                If .Enabled Then
50                    If .TimeInit > 0 Then
60                        .TimeInit = .TimeInit - 1
                              
70                        Select Case .TimeInit
                              Case 0
                                  
80                            Case 60
90                                SendData SendTarget.ToAll, 0, PrepareMessageShortMsj(29, FontTypeNames.FONTTYPE_GUILD, Int(.TimeInit / 60), , , , strModality(LoopC, .Modality))
                                  'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(LoopC, .Modality) & "» Las inscripciones abren en " & Int(.TimeInit / 60) & " minuto.", FontTypeNames.FONTTYPE_GUILD)
100                           Case 120
110                               SendData SendTarget.ToAll, 0, PrepareMessageShortMsj(28, FontTypeNames.FONTTYPE_GUILD, Int(.TimeInit / 60), , , , strModality(LoopC, .Modality))
                                  'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(LoopC, .Modality) & "» Las inscripciones abren en " & Int(.TimeInit / 60) & " minutos.", FontTypeNames.FONTTYPE_GUILD)
120                           Case 180
130                               SendData SendTarget.ToAll, 0, PrepareMessageShortMsj(28, FontTypeNames.FONTTYPE_GUILD, Int(.TimeInit / 60), , , , strModality(LoopC, .Modality))
                                  'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(LoopC, .Modality) & "» Las inscripciones abren en " & Int(.TimeInit / 60) & " minutos.", FontTypeNames.FONTTYPE_GUILD)
140                           Case 240
150                               SendData SendTarget.ToAll, 0, PrepareMessageShortMsj(28, FontTypeNames.FONTTYPE_GUILD, Int(.TimeInit / 60), , , , strModality(LoopC, .Modality))
                                  'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(LoopC, .Modality) & "» Las inscripciones abren en " & Int(.TimeInit / 60) & " minutos.", FontTypeNames.FONTTYPE_GUILD)
160                       End Select
                          
170                       If .TimeInit <= 0 Then
180                           SendData SendTarget.ToAll, 0, PrepareMessageShortMsj(30, FontTypeNames.FONTTYPE_GUILD, , , , , strModality(LoopC, .Modality))
                              'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(LoopC, .Modality) & "» Inscripciones abiertas. /INGRESAR " & strModality(LoopC, .Modality) & " para ingresar al evento. /INFOEVENTO para que entiendas en que consiste el evento.", FontTypeNames.FONTTYPE_GUILD)
190                           .TimeCancel = 0
200                       End If
                          
                      
210                   End If
                      
220                   If .TimeCancel > 0 And .TimeInit > 0 Then
230                       .TimeCancel = .TimeCancel - 1
                          
240                       If .TimeCancel <= 0 Then
                              'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(.Modality) & "» Ha sido cancelado ya que no se completaron los cupos.", FontTypeNames.FONTTYPE_WARNING)
250                           EventosDS.CloseEvent LoopC, "Evento " & strModality(LoopC, .Modality) & " cancelado.", True
260                       End If
270                   End If
                      
280                   If .TimeCount > 0 Then
290                       .TimeCount = .TimeCount - 1
                          
300                       For LoopY = LBound(.Users()) To UBound(.Users())
310                           If .Users(LoopY).Id > 0 Then
320                               If .TimeCount = 0 Then
                                      'WriteConsoleMsg .Users(LoopY).Id, "Cuenta» ¡Comienza!", FontTypeNames.FONTTYPE_FIGHT
330                                   WriteShortMsj .Users(LoopY).Id, 31, FontTypeNames.FONTTYPE_FIGHT
340                               Else
                                      'WriteConsoleMsg .Users(LoopY).Id, "Cuenta» " & .TimeCount, FontTypeNames.FONTTYPE_GUILD
350                                   WriteShortMsj .Users(LoopY).Id, 32, FontTypeNames.FONTTYPE_GUILD, .TimeCount
360                               End If
370                           End If
380                       Next LoopY
390                   End If
                      
400                   If .NpcIndex > 0 Then
410                      If Events(Npclist(.NpcIndex).flags.SlotEvent).TimeCount > 0 Then Exit Sub
420                      Call DagaRusa_MoveNpc(.NpcIndex)
430                   End If
                      
440                   If .TimeFinish > 0 Then
450                       .TimeFinish = .TimeFinish - 1
                          
460                       If .TimeFinish = 0 Then
470                           Call FinishEvent(LoopC)
480                       End If
490                   End If
500               End If
          
          
510           End With
520       Next LoopC
          
530   Exit Sub

error:
540       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : LoopEvent()"
End Sub

'/ FIN MANEJO DE LOS TIEMPOS
Public Function SetInfoEvento() As String
          Dim strTemp As String
          Dim LoopC As Integer
          
10        For LoopC = 1 To EventosDS.MAX_EVENT_SIMULTANEO
20            With Events(LoopC)
30                If .Enabled Then
40                    strTemp = strModality(LoopC, .Modality)
50                    SetInfoEvento = SetInfoEvento & strTemp & "» " & strDescEvent(LoopC, .Modality) & ". Se ingresa mediante: /INGRESAR " & strTemp
                      
60                    If .Run Then
70                        SetInfoEvento = SetInfoEvento & " Inscripciones cerradas."
80                    Else
90                        If .TimeInit > 0 Then
100                           SetInfoEvento = SetInfoEvento & " Inscripciones abren en " & Int(.TimeInit / 60) & " minuto/s"
110                       Else
120                           SetInfoEvento = SetInfoEvento & " Inscripciones abiertas."
130                       End If
140                   End If
                      
                      
150                   SetInfoEvento = SetInfoEvento & vbCrLf
160               End If
170           End With
180       Next LoopC

End Function

'// Funciones generales '//
Private Function FreeSlotEvent() As Byte
          Dim LoopC As Integer
          
10        For LoopC = 1 To MAX_EVENT_SIMULTANEO
20            If Not Events(LoopC).Enabled Then
30                FreeSlotEvent = LoopC
40                Exit For
50            End If
60        Next LoopC
End Function

Private Function FreeSlotUser(ByVal SlotEvent As Byte) As Byte
          Dim LoopC As Integer
          
10        With Events(SlotEvent)
20            For LoopC = 1 To MAX_USERS_EVENT
30                If .Users(LoopC).Id = 0 Then
40                    FreeSlotUser = LoopC
50                    Exit For
60                End If
70            Next LoopC
80        End With
          
End Function

Private Function FreeSlotArena() As Byte
          Dim LoopC As Integer
          
10        FreeSlotArena = 0
          
20        For LoopC = 1 To MAX_MAP_FIGHT
30            If MapEvent.Fight(LoopC).Run = False Then
40                FreeSlotArena = LoopC
50                Exit For
60            End If
70        Next LoopC
End Function
Public Function strUsersEvent(ByVal SlotEvent As Byte) As String

          ' Texto que marca los personajes que están en el evento.
          Dim LoopC As Integer
          
10        With Events(SlotEvent)
20            For LoopC = LBound(.Users()) To UBound(.Users())
30                If .Users(LoopC).Id > 0 Then
40                    strUsersEvent = strUsersEvent & UserList(.Users(LoopC).Id).Name & "-"
50                End If
60            Next LoopC
70        End With
End Function
Private Function CheckAllowedClasses(ByRef AllowedClasses() As Byte) As String
          Dim LoopC As Integer
          
10        For LoopC = 1 To NUMCLASES
20            If AllowedClasses(LoopC) = 1 Then
30                If CheckAllowedClasses = vbNullString Then
40                    CheckAllowedClasses = ListaClases(LoopC)
50                Else
60                    CheckAllowedClasses = CheckAllowedClasses & ", " & ListaClases(LoopC)
70                End If
80            End If
90        Next LoopC
          
End Function

Private Function SearchLastUserEvent(ByVal SlotEvent As Byte) As Integer

          ' Busca el último usuario que está en el torneo. En todos los eventos será el ganador.
          
          Dim LoopC As Integer
          
10        With Events(SlotEvent)
20            For LoopC = LBound(.Users()) To UBound(.Users())
30                If .Users(LoopC).Id > 0 Then
40                    SearchLastUserEvent = .Users(LoopC).Id
50                    Exit For
60                End If
70            Next LoopC
80        End With
End Function

Private Function SearchSlotEvent(ByVal Modality As String) As Byte
          Dim LoopC As Integer
          
10        SearchSlotEvent = 0
          
20        For LoopC = 1 To MAX_EVENT_SIMULTANEO
30            With Events(LoopC)
40                If StrComp(UCase$(strModality(LoopC, .Modality)), UCase$(Modality)) = 0 Then
50                    SearchSlotEvent = LoopC
60                    Exit For
70                End If
80            End With
90        Next LoopC

End Function

Private Sub EventWarpUser(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)
10    On Error GoTo error

          ' Teletransportamos a cualquier usuario que cumpla con la regla de estar en un evento.
          
          Dim Pos As WorldPos
          
20        With UserList(UserIndex)
30            Pos.map = map
40            Pos.X = X
50            Pos.Y = Y
              
60            ClosestStablePos Pos, Pos
70            WarpUserChar UserIndex, Pos.map, Pos.X, Pos.Y, False
          
80        End With
          
90    Exit Sub

error:
100       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : EventWarpUser()"
End Sub
Private Sub ResetEvent(ByVal Slot As Byte)
10    On Error GoTo error

          Dim LoopC As Integer
          
20        With Events(Slot)
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Id > 0 Then
50                    AbandonateEvent .Users(LoopC).Id, False
60                End If
70            Next LoopC
              
80            If .NpcIndex > 0 Then Call QuitarNPC(.NpcIndex)
              
90            .Enabled = False
100           .Run = False
110           .npcUserIndex = 0
120           .TimeFinish = 0
130           .TeamCant = 0
140           .Quotas = 0
150           .Inscribed = 0
160           .DspInscription = 0
170           .GldInscription = 0
180           .LvlMax = 0
190           .LvlMin = 0
200           .TimeCancel = 0
210           .NpcIndex = 0
220           .TimeInit = 0
230           .TimeCount = 0
240           .CharBody = 0
250           .CharHp = 0
260           .Modality = 0
              
270           For LoopC = LBound(.AllowedClasses()) To UBound(.AllowedClasses())
280               .AllowedClasses(LoopC) = 0
290           Next LoopC
              
300       End With
310   Exit Sub

error:
320       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : ResetEvent()"
End Sub

Private Function CheckUserEvent(ByVal UserIndex As Integer, ByVal SlotEvent As Byte, ByRef ErrorMsg As Integer) As Boolean
10    On Error GoTo error

20        CheckUserEvent = False
              
30        With UserList(UserIndex)
40            If .flags.Muerto Then
50                ErrorMsg = 33
60                Exit Function
70            End If
80            If .flags.Mimetizado Then
90                ErrorMsg = 34
100               Exit Function
110           End If
              
120           If .flags.Montando Then
130               ErrorMsg = 35
140               Exit Function
150           End If
              
160           If .flags.invisible Then
170               ErrorMsg = 36
180               Exit Function
190           End If
              
200           If .flags.SlotEvent > 0 Then
210               ErrorMsg = 37
220               Exit Function
230           End If
              
240           If .flags.SlotReto > 0 Then
250               ErrorMsg = 37
260               Exit Function
270           End If
              
280           If .flags.InCVC Then
290               ErrorMsg = 37
300               Exit Function
310           End If
              
320           If .Counters.Pena > 0 Then
330               ErrorMsg = 38
340               Exit Function
350           End If
              
360           If MapInfo(.Pos.map).Pk Then
370               ErrorMsg = 39
380               Exit Function
390           End If
              
400           If .flags.Comerciando Then
410               ErrorMsg = 40
420               Exit Function
430           End If
              
440           If Not Events(SlotEvent).Enabled Or Events(SlotEvent).TimeInit > 0 Then
450               ErrorMsg = 41
460               Exit Function
470           End If
              
480           If Events(SlotEvent).Run Then
490               ErrorMsg = 42
500               Exit Function
510           End If
              
              
520           If Events(SlotEvent).LvlMin <> 0 Then
530               If Events(SlotEvent).LvlMin > .Stats.ELV Then
540                   ErrorMsg = 43
550                   Exit Function
560               End If
570           End If
              
580           If Events(SlotEvent).LvlMin <> 0 Then
590               If Events(SlotEvent).LvlMax < .Stats.ELV Then
600                   ErrorMsg = 43
610                   Exit Function
620               End If
630           End If
              
640           If Events(SlotEvent).AllowedClasses(.clase) = 0 Then
650               ErrorMsg = 44
660               Exit Function
670           End If
              
              
680           If Events(SlotEvent).GldInscription > .Stats.Gld Then
690               ErrorMsg = 45
700               Exit Function
710           End If
              
720           If Events(SlotEvent).DspInscription > 0 Then
730               If Not TieneObjetos(880, Events(SlotEvent).DspInscription, UserIndex) Then
740                   ErrorMsg = 46
750                   Exit Function
760               End If
770           End If
              
780           If Events(SlotEvent).Inscribed = Events(SlotEvent).Quotas Then
790               ErrorMsg = 47
800               Exit Function
810           End If
              
820       End With
830       CheckUserEvent = True
          
840   Exit Function

error:
850       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : CheckUserEvent()"
End Function

' EDICIÓN GENERAL
Public Function strModality(ByVal SlotEvent As Byte, ByVal Modality As eModalityEvent) As String

          ' Modalidad de cada evento
          
10        Select Case Modality
              Case eModalityEvent.CastleMode
20                strModality = "CastleMode"
                  
30            Case eModalityEvent.DagaRusa
40                strModality = "DagaRusa"
                  
50            Case eModalityEvent.DeathMatch
60                strModality = "DeathMatch"
                  
70            Case eModalityEvent.Aracnus
80                strModality = "Aracnus"
                  
90            Case eModalityEvent.HombreLobo
100               strModality = "HombreLobo"
                  
110           Case eModalityEvent.Minotauro
120               strModality = "Minotauro"
              
130           Case eModalityEvent.Busqueda
140               strModality = "Busqueda"
              
150           Case eModalityEvent.Unstoppable
160               strModality = "Unstoppable"
              
170           Case eModalityEvent.Invasion
180               strModality = "Invasion"
                  
190           Case eModalityEvent.Enfrentamientos
200               strModality = Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant
210       End Select
End Function
Private Function strDescEvent(ByVal SlotEvent As Byte, ByVal Modality As eModalityEvent) As String

          ' Descripción del evento en curso.
10        Select Case Modality
              Case eModalityEvent.CastleMode
20                strDescEvent = "» Los usuarios entrarán de forma aleatorea para formar dos equipos. Ambos equipos deberán defender a su rey y a su vez atacar al del equipo contrario."
30            Case eModalityEvent.DagaRusa
40                strDescEvent = "» Los usuarios se teletransportarán a una posición donde estará un asesino dispuesto a apuñalarlos y acabar con su vida. El último que quede en pie es el ganador del evento."
50            Case eModalityEvent.DeathMatch
60                strDescEvent = "» Los usuarios ingresan y luchan en una arena donde se toparan con todos los demás concursantes. El que logre quedar en pie, será el ganador."
70            Case eModalityEvent.Aracnus
80                strDescEvent = "» Un personaje es escogido al azar, para convertirse en una araña gigante la cual podrá envenenar a los demas concursantes acabando con su vida en el evento."
90            Case eModalityEvent.Busqueda
100               strDescEvent = "» Los personajes son teletransportados en un mapa donde su función principal será la recolección de objetos en el piso, para que así luego de tres minutos, el que recolecte más, ganará el evento."
110           Case eModalityEvent.Unstoppable
120               strDescEvent = "» Los personajes lucharan en un TODOS vs TODOS, donde los muertos no irán a su mapa de origen, si no que volverán a revivir para tener chances de ganar el evento. El que logre matar más personajes, ganará el evento."
130           Case eModalityEvent.Invasion
140               strDescEvent = "» Los personajes son llevados a un mapa donde aparecerán criaturas únicas de DesteriumAO, cada criatura dará una recompensa única y los usuarios tendrán chances de entrenar sus personajes."
150           Case eModalityEvent.Enfrentamientos
160               If Events(SlotEvent).TeamCant = 1 Then
170                   strDescEvent = "» Los usuarios combatirán en duelos 1vs1"
180               Else
190                   strDescEvent = "» Los usuarios combatirán en duelos " & Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant & " donde se escogerán las parejas al azar."
200               End If
210       End Select
End Function
Private Sub InitEvent(ByVal SlotEvent As Byte)
          
10        Select Case Events(SlotEvent).Modality
              Case eModalityEvent.CastleMode
20                Call InitCastleMode(SlotEvent)
                  
30            Case eModalityEvent.DagaRusa
40                Call InitDagaRusa(SlotEvent)
                  
50            Case eModalityEvent.DeathMatch
60                Call InitDeathMatch(SlotEvent)
                  
70            Case eModalityEvent.Aracnus
80                Call InitEventTransformation(SlotEvent, 254, 6500, 211, 70, 36)
                  
90            Case eModalityEvent.HombreLobo
100               Call InitEventTransformation(SlotEvent, 255, 3500, 211, 70, 36)
                  
110           Case eModalityEvent.Minotauro
120               Call InitEventTransformation(SlotEvent, 253, 2500, 211, 70, 36)
              
130           Case eModalityEvent.Busqueda
140               Call InitBusqueda(SlotEvent)
                  
150           Case eModalityEvent.Unstoppable
160               InitUnstoppable SlotEvent
                  
170           Case eModalityEvent.Invasion
              
180           Case eModalityEvent.Enfrentamientos
190               Call InitFights(SlotEvent)
              
200           Case Else
210               Exit Sub
              
220       End Select
230   Exit Sub

error:
240       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : InitEvent() EN EL EVENTO " & Events(SlotEvent).Modality & "."
End Sub
Public Function CanAttackUserEvent(ByVal UserIndex As Integer, ByVal Victima As Integer) As Boolean
          
          ' Si el personaje es del mismo team, no se puede atacar al usuario.
          Dim VictimaSlotUserEvent As Byte
          
10      VictimaSlotUserEvent = UserList(Victima).flags.SlotUserEvent
          
        If UserList(UserIndex).flags.SlotEvent > 0 And UserList(Victima).flags.SlotEvent > 0 Then
            With UserList(UserIndex)
40                If Events(.flags.SlotEvent).Users(VictimaSlotUserEvent).Team = Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).Team Then
50                    CanAttackUserEvent = False
60                    Exit Function
70                End If
            End With
        End If
   
   
           CanAttackUserEvent = True
          
110   Exit Function

error:
120       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : CanAttackUserEvent()"
End Function


Private Sub PrizeUser(ByVal UserIndex As Integer, Optional ByVal MsjConsole As Boolean = True)
10        On Error GoTo error
          
          ' Premios de los eventos
          
          Dim SlotEvent As Byte
          Dim SlotUserEvent As Byte
          Dim Obj As Obj
          Dim strReWard As String
          
20        SlotEvent = UserList(UserIndex).flags.SlotEvent
30        SlotUserEvent = UserList(UserIndex).flags.SlotUserEvent
          
40        With Events(SlotEvent)
50            If .GldInscription > 0 Then
60                With UserList(UserIndex)
70                    .Stats.Gld = .Stats.Gld + (Events(SlotEvent).GldInscription * Events(SlotEvent).Quotas)
80                    WriteUpdateGold UserIndex
90                    strReWard = (Events(SlotEvent).GldInscription * Events(SlotEvent).Quotas) & " Monedas de oro. "
100               End With
110           End If
              
120           If .DspInscription > 0 Then
130               Obj.ObjIndex = 880
140               Obj.Amount = .DspInscription * .Quotas
                  
150               With UserList(UserIndex)

160                   If Not MeterItemEnInventario(UserIndex, Obj) Then
                          'Call TirarItemAlPiso(.Pos, Obj)
                          
170                       SendData SendTarget.ToAdmins, 0, PrepareMessageShortMsj(49, FontTypeNames.FONTTYPE_ADMIN, , , , , .Name)
180                       WriteShortMsj UserIndex, 50, FontTypeNames.FONTTYPE_WARNING
190                   End If
                      
200                   strReWard = strReWard & (Events(SlotEvent).DspInscription * Events(SlotEvent).Quotas) & " Monedas DSP."
                      
210               End With
220           End If
              
230           With UserList(UserIndex)
                  .Stats.Points = .Stats.Points + 15
240               .Stats.TorneosGanados = .Stats.TorneosGanados + 1

                  WriteUpdatePoints UserIndex
250           End With
              
              
260           If MsjConsole Then Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(SlotEvent, .Modality) & "» Premio recibido: " & strReWard, FontTypeNames.FONTTYPE_GUILD))
270       End With
          
280   Exit Sub

error:
290       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : PrizeUser()"
End Sub
Private Sub ChangeBodyEvent(ByVal SlotEvent As Byte, ByVal UserIndex As Integer, ByVal ChangeHead As Boolean)
10    On Error GoTo error

          ' En caso de que el evento cambie el body, de lo cambiamos.
20        With UserList(UserIndex)
30            .CharMimetizado.body = .Char.body
40            .CharMimetizado.Head = .Char.Head
50            .CharMimetizado.CascoAnim = .Char.CascoAnim
60            .CharMimetizado.ShieldAnim = .Char.ShieldAnim
70            .CharMimetizado.WeaponAnim = .Char.WeaponAnim

80            .Char.body = Events(SlotEvent).CharBody
90            .Char.Head = IIf(ChangeHead = False, .Char.Head, 0)
100           .Char.CascoAnim = 0
110           .Char.ShieldAnim = 0
120           .Char.WeaponAnim = 0
                      
130           ChangeUserChar UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, True
140           RefreshCharStatus UserIndex
          
150       End With
          
160   Exit Sub

error:
170       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : ChangeBodyEvent()"
End Sub

Private Function ResetBodyEvent(ByVal SlotEvent As Byte, ByVal UserIndex As Integer)

10    On Error GoTo error

          ' En caso de que el evento cambie el body del personaje, se lo restauramos.
          
20        With UserList(UserIndex)
30            If .flags.Muerto Then Exit Function
              'If Events(SlotEvent).Users(.flags.SlotUserEvent).Selected = 0 Then Exit Function
              
40            If .CharMimetizado.body > 0 Then
50                .Char.body = .CharMimetizado.body
60                .Char.Head = .CharMimetizado.Head
70                .Char.CascoAnim = .CharMimetizado.CascoAnim
80                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
90                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                  
                  
100               .CharMimetizado.body = 0
110               .CharMimetizado.Head = 0
120               .CharMimetizado.CascoAnim = 0
130               .CharMimetizado.ShieldAnim = 0
140               .CharMimetizado.WeaponAnim = 0
                  
150               .showName = True
                  
160               ChangeUserChar UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, True
170               RefreshCharStatus UserIndex
180           End If
          
190       End With
          
200   Exit Function

error:
210       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : ResetBodyEvent()"
End Function

Private Sub ChangeHpEvent(ByVal UserIndex As Integer)

10    On Error GoTo error
          ' En caso de que el evento edite la vida del personaje, se la editamos.
          
          Dim SlotEvent As Byte
          
20        With UserList(UserIndex)
30            SlotEvent = .flags.SlotEvent
              
40            .Stats.OldHp = .Stats.MaxHp
              
50            .Stats.MaxHp = Events(SlotEvent).CharHp
60            .Stats.MinHp = .Stats.MaxHp
              
70            WriteUpdateUserStats UserIndex
          
80        End With
90    Exit Sub

error:
100       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : ChangeHpEvent()"
End Sub

Private Sub ResetHpEvent(ByVal UserIndex As Integer)

10    On Error GoTo error
          ' En caso de que el evento haya editado la vida de un personaje, se la volvemos a restaurar.
          
20        With UserList(UserIndex)
30            If .Stats.OldHp = 0 Then Exit Sub
40            .Stats.MaxHp = .Stats.OldHp
              '.Stats.MinHp = .Stats.MaxHp
50            .Stats.OldHp = 0
60            WriteUpdateHP UserIndex
              
70        End With
          
80    Exit Sub

error:
90        LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : ResetHpEvent()"
End Sub




'// Fin Funciones generales '//

Public Sub NewEvent(ByVal UserIndex As Integer, _
                    ByVal Modality As eModalityEvent, _
                    ByVal Quotas As Byte, _
                    ByVal LvlMin As Byte, _
                    ByVal LvlMax As Byte, _
                    ByVal GldInscription As Long, _
                    ByVal DspInscription As Long, _
                    ByVal TimeInit As Long, _
                    ByVal TimeCancel As Long, _
                    ByVal TeamCant As Byte, _
                    ByVal PrizeAccumulated As Boolean, _
                    ByVal LimitRed As Integer, _
                    ByVal PrizeDsp As Integer, _
                    ByVal PrizeGld As Integer, _
                    ByVal ObjIndex As Integer, _
                    ByVal ObjAmount As Integer, _
                    ByVal WinFollow As Boolean, _
                    ByVal ValidItem As Boolean, _
                    ByRef AllowedFaction() As eFaction, _
                    ByRef AllowedClasses() As Byte)
                          
10        On Error GoTo error
                          
          Dim Slot As Integer
          Dim strTemp As String

20        Slot = FreeSlotEvent()
          
30        If Slot = 0 Then
              'WriteConsoleMsg Userindex, "No hay más lugar disponible para crear un evento simultaneo. Espera a que termine alguno o bien cancela alguno.", FontTypeNames.FONTTYPE_INFO
40            WriteShortMsj UserIndex, 48, FontTypeNames.FONTTYPE_INFO
50            Exit Sub
60        Else
70            With Events(Slot)
80                .Enabled = True
90                .Modality = Modality
100               .TeamCant = TeamCant
110               .Quotas = Quotas
120               .LvlMin = LvlMin
130               .LvlMax = LvlMax
140               .GldInscription = GldInscription
150               .DspInscription = DspInscription
160               .AllowedClasses = AllowedClasses
                  .AllowedFaction = AllowedFaction
170               .TimeInit = TimeInit
180               .TimeCancel = TimeCancel
                
                  .ValidItem = ValidItem
                  .PrizeAccumulated = PrizeAccumulated
                  .LimitRed = LimitRed
                  .PrizeDsp = PrizeDsp
                  .PrizeGld = PrizeGld
                  .PrizeObj.ObjIndex = ObjIndex
                  .PrizeObj.Amount = ObjAmount
                  .WinFollow = WinFollow
                  
190               ReDim .Users(1 To .Quotas) As tUserEvent
                  
                  ' strModality devuelve: "Evento '1vs1' : Descripción"
200               strTemp = strModality(Slot, .Modality) & strDescEvent(Slot, .Modality) & vbCrLf & "Cupos: " & .Quotas & ". Nivel permitido: " & .LvlMin & "-" & .LvlMax & "." & vbCrLf
                  
                  
210               If .GldInscription > 0 And .DspInscription > 0 Then
220                   strTemp = strTemp & "Inscripción requerida: " & .GldInscription & " monedas de oro y " & .DspInscription & " monedas DSP."
230               ElseIf .GldInscription > 0 Then
240                   strTemp = strTemp & "Inscripción requerida: " & .GldInscription & " monedas de oro."
250               ElseIf .DspInscription > 0 Then
260                   strTemp = strTemp & "Inscripción requerida: " & .DspInscription & " monedas DSP."
270               Else
280                   strTemp = strTemp & "Inscripción GRATIS. "
290               End If
                  
300               strTemp = strTemp & "Clases permitidas: " & CheckAllowedClasses(AllowedClasses) & ". Comando para ingresar /INGRESAR " & strModality(Slot, .Modality) & vbCrLf
                  
310               If .TimeInit = 60 Then
320                   strTemp = strTemp & "Las inscripciones abren en " & Int(.TimeInit / 60) & " minuto."
330               Else
340                   strTemp = strTemp & "Las inscripciones abren en " & Int(.TimeInit / 60) & " minutos."
350               End If
360               LoadMapEvent
370           End With
              
380           SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strTemp, FontTypeNames.FONTTYPE_INFOBOLD)
390       End If
          
400   Exit Sub

error:
410       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : NewEvent()"
End Sub
Private Sub GiveBack_Inscription(ByVal SlotEvent As Byte)
10    On Error GoTo error

          Dim LoopC As Integer
          Dim Obj As Obj
          
20        With Events(SlotEvent)
          
30            Obj.ObjIndex = 880
40            Obj.Amount = .DspInscription
              
50            For LoopC = LBound(.Users()) To UBound(.Users())
60                If .Users(LoopC).Id > 0 Then
70                    If .DspInscription > 0 Then
80                        If Not MeterItemEnInventario(.Users(LoopC).Id, Obj) Then
                              'Call TirarItemAlPiso(UserList(.Users(LoopC).Id).Pos, Obj)
                              
                              'SendData SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("¡¡ATENCIÓN GM!! Al personaje " & UserList(.Users(LoopC).Id).Name & " no se le entrego el dsp porque no tenia espacio en el inventario.", FontTypeNames.FONTTYPE_ADMIN)
                              'WriteConsoleMsg .Users(LoopC).Id, "¡¡HEMOS NOTADO que no tienes espacio en el inventario para recibir los DSP ganadores. Un gm se contactará contigo a la brevedad.", FontTypeNames.FONTTYPE_WARNING
90                            SendData SendTarget.ToAdmins, 0, PrepareMessageShortMsj(49, FontTypeNames.FONTTYPE_ADMIN, , , , , UserList(.Users(LoopC).Id).Name)
100                           WriteShortMsj .Users(LoopC).Id, 50, FontTypeNames.FONTTYPE_WARNING
                          
110                       End If
120                   End If
                      
130                   If .GldInscription > 0 Then
140                       UserList(.Users(LoopC).Id).Stats.Gld = UserList(.Users(LoopC).Id).Stats.Gld + .GldInscription
150                       WriteUpdateGold (.Users(LoopC).Id)
160                   End If
170               End If
180           Next LoopC
190       End With
          
200   Exit Sub

error:
210       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : GiveBack_Inscription()"
End Sub
Public Sub CloseEvent(ByVal Slot As Byte, Optional ByVal MsgConsole As String = vbNullString, Optional ByVal Cancel As Boolean = False)
10    On Error GoTo error
          
20        With Events(Slot)
              ' Devolvemos la inscripción
30            If Cancel Then
40                Call GiveBack_Inscription(Slot)
50            End If
              
60            If MsgConsole <> vbNullString Then SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(MsgConsole, FontTypeNames.FONTTYPE_ORO)
              

              
70            Call ResetEvent(Slot)
80        End With
          
90    Exit Sub

error:
100       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : CloseEvent()"
End Sub




Public Sub ParticipeEvent(ByVal UserIndex As Integer, ByVal Modality As String)
10    On Error GoTo error

          Dim ErrorMsg As Integer
          Dim SlotUser As Byte
          Dim Pos As WorldPos
          Dim SlotEvent As Integer
          
20        SlotEvent = SearchSlotEvent(Modality)
          
30        If SlotEvent = 0 Then
              'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Error Fatal TESTEO", FontTypeNames.FONTTYPE_ADMIN)
40            Exit Sub
50        End If
          
60        With UserList(UserIndex)
70            If CheckUserEvent(UserIndex, SlotEvent, ErrorMsg) Then
80                SlotUser = FreeSlotUser(SlotEvent)
                  
90                .flags.SlotEvent = SlotEvent
100               .flags.SlotUserEvent = SlotUser
                  
110               .PosAnt.map = .Pos.map
120               .PosAnt.X = .Pos.X
130               .PosAnt.Y = .Pos.Y
                  
140               .Stats.Gld = .Stats.Gld - Events(SlotEvent).GldInscription
150               Call WriteUpdateGold(UserIndex)
                  
160               Call QuitarObjetos(880, Events(SlotEvent).DspInscription, UserIndex)
                  
170               With Events(SlotEvent)
180                   Pos.map = 211
190                   Pos.X = 30
200                   Pos.Y = 21
                      
210                   Call FindLegalPos(UserIndex, Pos.map, Pos.X, Pos.Y)
220                   Call WarpUserChar(UserIndex, Pos.map, Pos.X, Pos.Y, False)
                  
230                   .Users(SlotUser).Id = UserIndex
240                   .Inscribed = .Inscribed + 1
                      
                      
                      'WriteConsoleMsg Userindex, "Has ingresado al evento " & strModality(SlotEvent, .Modality) & ". Espera a que se completen los cupos para que comience.", FontTypeNames.FONTTYPE_INFO
250                   WriteShortMsj UserIndex, 51, FontTypeNames.FONTTYPE_INFO, , , , , strModality(SlotEvent, .Modality)
                      LogEventos "El personaje " & UserList(UserIndex).Name & " ingresó el evento de modalidad " & strModality(SlotEvent, .Modality)
                      
260                   If .Inscribed = .Quotas Then
                          'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(SlotEvent, .Modality) & "» Los cupos han sido alcanzados. Les deseamos mucha suerte a cada uno de los participantes y que gane el mejor!", FontTypeNames.FONTTYPE_GUILD)
270                       SendData SendTarget.ToAll, 0, PrepareMessageShortMsj(52, FontTypeNames.FONTTYPE_GUILD, , , , , strModality(SlotEvent, .Modality))
                          
280                       .Run = True
290                       InitEvent SlotEvent
300                       Exit Sub
310                   End If
320               End With
              
330           Else
340               WriteShortMsj UserIndex, ErrorMsg, FontTypeNames.FONTTYPE_WARNING
              
350           End If
360       End With
370   Exit Sub

error:
380       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : ParticipeEvent()"
End Sub



Public Sub AbandonateEvent(ByVal UserIndex As Integer, _
                            Optional ByVal MsgAbandonate As Boolean = False, _
                            Optional ByVal Forzado As Boolean = False)
          
10    On Error GoTo error

          Dim Pos As WorldPos
          Dim SlotEvent As Byte
          Dim SlotUserEvent As Byte
          Dim UserTeam As Byte
          Dim UserMapFight As Byte
          
20        With UserList(UserIndex)
30            SlotEvent = .flags.SlotEvent
40            SlotUserEvent = .flags.SlotUserEvent
              
50            If SlotEvent > 0 And SlotUserEvent > 0 Then
60                With Events(SlotEvent)
                      LogEventos "El personaje " & UserList(UserIndex).Name & " abandonó el evento de modalidad " & strModality(SlotEvent, .Modality)
70
                        If .Inscribed > 0 Then .Inscribed = .Inscribed - 1

80                        UserTeam = .Users(SlotUserEvent).Team
90                        UserMapFight = .Users(SlotUserEvent).MapFight
                          
100                       .Users(SlotUserEvent).Id = 0
110                       .Users(SlotUserEvent).Team = 0
                          .Users(SlotUserEvent).Value = 0
130                       .Users(SlotUserEvent).Selected = 0
140                       .Users(SlotUserEvent).MapFight = 0
                          
150                       UserList(UserIndex).flags.SlotEvent = 0
160                       UserList(UserIndex).flags.SlotUserEvent = 0
170                       UserList(UserIndex).flags.FightTeam = 0
                          
180                       Select Case .Modality
                              Case eModalityEvent.Aracnus, eModalityEvent.HombreLobo, eModalityEvent.Minotauro
190                               If Forzado And .Inscribed > 1 Then
200                                   If .Users(SlotUserEvent).Selected = 1 Then
210                                       Transformation_SelectionUser SlotEvent
220                                   End If
230                               End If
                                  
240                           Case eModalityEvent.DagaRusa
250                               If Forzado And .Run Then
260                                   Call WriteUserInEvent(UserIndex)
                                      
270                                   If .Users(SlotUserEvent).Value = 0 Then
280                                       Npclist(.NpcIndex).flags.InscribedPrevio = Npclist(.NpcIndex).flags.InscribedPrevio - 1
290                                   End If
300                               End If
                                  
310                           Case eModalityEvent.Enfrentamientos
320                               If Forzado Then
330                                   If UserMapFight > 0 Then
340                                       If Not Fight_CheckContinue(UserIndex, SlotEvent, UserTeam) Then
350                                           Fight_WinForzado UserIndex, SlotEvent, UserMapFight
360                                       End If
370                                   End If
380                               End If
                                  
390                               If UserList(UserIndex).Counters.TimeFight > 0 Then
400                                   UserList(UserIndex).Counters.TimeFight = 0
410                                   Call WriteUserInEvent(UserIndex)
420                               End If
                                  
430                       End Select
                                  
440                       Pos.map = UserList(UserIndex).PosAnt.map
450                       Pos.X = UserList(UserIndex).PosAnt.X
460                       Pos.Y = UserList(UserIndex).PosAnt.Y
                          
470                       Call FindLegalPos(UserIndex, Pos.map, Pos.X, Pos.Y)
480                       Call WarpUserChar(UserIndex, Pos.map, Pos.X, Pos.Y, False)
                          
490                       If Events(SlotEvent).CharBody <> 0 Then
500                           Call ResetBodyEvent(SlotEvent, UserIndex)
510                       End If
                  
520                       If UserList(UserIndex).Stats.OldHp <> 0 Then
530                           ResetHpEvent UserIndex
540                       End If
                  
550                       UserList(UserIndex).showName = True
560                       RefreshCharStatus UserIndex
                          
                          'If MsgAbandonate Then WriteConsoleMsg Userindex, "Has abandonado el evento. Podrás recibir una pena por hacer esto.", FontTypeNames.FONTTYPE_WARNING
570                       If MsgAbandonate Then WriteShortMsj UserIndex, 53, FontTypeNames.FONTTYPE_WARNING
                          
                          
                          ' Abandono general del evento
580                       If .Inscribed = 1 And Forzado Then
590                           Call FinishEvent(SlotEvent)
                          
600                           CloseEvent SlotEvent
610                           Exit Sub
620                       End If
                          
                          
630               End With
640           End If
              
              
650       End With
          
660   Exit Sub

error:
670       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : AbandonateEvent()"
End Sub

Private Sub FinishEvent(ByVal SlotEvent As Byte)

10    On Error GoTo error
          Dim UserIndex As Integer
          Dim IsSelected As Boolean
          
20        With Events(SlotEvent)
30            Select Case .Modality
                  Case eModalityEvent.CastleMode
40                    UserIndex = SearchLastUserEvent(SlotEvent)
50                    CastleMode_Premio UserIndex, False
                      
60                Case eModalityEvent.DagaRusa
70                    DagaRusa_CheckWin SlotEvent
                      
80                Case eModalityEvent.DeathMatch
90                    UserIndex = SearchLastUserEvent(SlotEvent)
100                   DeathMatch_Premio UserIndex
                      
110               Case eModalityEvent.Aracnus, eModalityEvent.HombreLobo, eModalityEvent.Minotauro
120                   UserIndex = SearchLastUserEvent(SlotEvent)
                      
130                   If .Users(UserList(UserIndex).flags.SlotUserEvent).Selected = 1 Then IsSelected = True
                      
140                   Transformation_Premio UserIndex, IsSelected, 250000
                      
150               Case eModalityEvent.Busqueda
160                   Busqueda_SearchWin SlotEvent
                      
170               Case eModalityEvent.Unstoppable
180                   Unstoppable_UserWin SlotEvent
                      
190           End Select
200       End With
          
          
210   Exit Sub

error:
220       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : FinishEvent()"
End Sub


'#################EVENTO CASTLE MODE##########################
Public Function CanAttackReyCastle(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
10        With UserList(UserIndex)
20            If .flags.SlotEvent > 0 Then
30                If Npclist(NpcIndex).flags.TeamEvent = Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).Team Then
40                    CanAttackReyCastle = False
50                    Exit Function
60                End If
70            End If
          
          
80            CanAttackReyCastle = True
90        End With
End Function
Private Sub CastleMode_InitRey()
10        On Error GoTo error
          
          Dim NpcIndex As Integer
          Const NumRey As Integer = 697
          Dim Pos As WorldPos
          Dim LoopX As Integer, LoopY As Integer
          Const Rango As Byte = 5
          
20        For LoopX = YMinMapSize To YMaxMapSize
30            For LoopY = XMinMapSize To XMaxMapSize
40                If InMapBounds(212, LoopX, LoopY) Then
50                    If MapData(212, LoopX, LoopY).NpcIndex > 0 Then
60                        Call QuitarNPC(MapData(212, LoopX, LoopY).NpcIndex)
70                    End If
80                End If
90            Next LoopY
100       Next LoopX
          
110       Pos.map = 212
              
120       Pos.X = 74
130       Pos.Y = 24
140       NpcIndex = SpawnNpc(NumRey, Pos, False, False)
150       Npclist(NpcIndex).flags.TeamEvent = 1
          
160       Pos.X = 19
170       Pos.Y = 34
180       NpcIndex = SpawnNpc(NumRey, Pos, False, False)
190       Npclist(NpcIndex).flags.TeamEvent = 2
          
200   Exit Sub

error:
210       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : CastleMode_InitRey()"
End Sub

Public Sub InitCastleMode(ByVal SlotEvent As Byte)
10    On Error GoTo error

          Dim LoopC As Integer
          
          Const NumRey As Integer = 697
          Dim NpcIndex As Integer
          Dim Pos As WorldPos
          
          ' Spawn the npc castle mode
20        CastleMode_InitRey
          
30        With Events(SlotEvent)
40            For LoopC = LBound(.Users()) To UBound(.Users())
50                If .Users(LoopC).Id > 0 Then
60                    If LoopC > (UBound(.Users()) / 2) Then
70                        .Users(LoopC).Team = 2
80                        Pos.map = 212
90                        Pos.X = 19
100                       Pos.Y = 34
                          
110                       Call FindLegalPos(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y)
120                       Call WarpUserChar(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y, False)
130                   Else
140                       .Users(LoopC).Team = 1
150                       Pos.map = 212
160                       Pos.X = 74
170                       Pos.Y = 24
                          
180                       Call FindLegalPos(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y)
190                       Call WarpUserChar(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y, False)
                          
200                   End If
210               End If
220           Next LoopC
230       End With
          
240   Exit Sub

error:
250       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : InitCastleMode()"
End Sub
Public Sub CastleMode_UserRevive(ByVal UserIndex As Integer)

10    On Error GoTo error
          Dim LoopC As Integer
          Dim Pos As WorldPos
          
20        With UserList(UserIndex)
30            If .flags.SlotEvent > 0 Then
40                Call RevivirUsuario(UserIndex)
                  
                  
50                Pos.map = 212
60                Pos.X = RandomNumber(20, 80)
70                Pos.Y = RandomNumber(20, 80)
                  
80                Call ClosestLegalPos(Pos, Pos)
                  'Call FindLegalPos(Userindex, Pos.Map, Pos.X, Pos.Y)
90                Call WarpUserChar(UserIndex, Pos.map, Pos.X, Pos.Y, True)
              
100           End If
110       End With
          
120   Exit Sub

error:
130       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : CastleMode_UserRevive()"
End Sub

Public Sub FinishCastleMode(ByVal SlotEvent As Byte, ByVal UserEventSlot As Integer)
10    On Error GoTo error

          Dim LoopC As Integer
          Dim strTemp As String
          Dim NpcIndex As Integer
          Dim MiObj As Obj
          
20        With Events(SlotEvent)
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Id > 0 Then
50                    If .Users(LoopC).Team = .Users(UserEventSlot).Team Then
60                        If LoopC = UserEventSlot Then
70                            CastleMode_Premio .Users(LoopC).Id, True
80                        Else
90                            CastleMode_Premio .Users(LoopC).Id, False
100                       End If
                          
110                       If strTemp = vbNullString Then
120                           strTemp = UserList(.Users(LoopC).Id).Name
130                       Else
140                           strTemp = strTemp & ", " & UserList(.Users(LoopC).Id).Name
150                       End If
160                   End If
170               End If
180           Next LoopC
              
              
190           CloseEvent SlotEvent, "CastleMode» Ha finalizado. Ha ganado el equipo de " & UCase$(strTemp)
200       End With
          
210   Exit Sub

error:
220       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : FinishCastleMode()"
End Sub

Private Sub CastleMode_Premio(ByVal UserIndex As Integer, ByVal KillRey As Boolean)
10    On Error GoTo error

          ' Entregamos el premio del CastleMode
          Dim MiObj As Obj
          
20        With UserList(UserIndex)
30            .Stats.Gld = .Stats.Gld + 250000
              'WriteConsoleMsg Userindex, "Felicitaciones, has recibido 250.000 monedas de oro por haber ganado el evento!", FontTypeNames.FONTTYPE_INFO
40            WriteShortMsj UserIndex, 54, FontTypeNames.FONTTYPE_INFO, , , , 250000
              
50            If KillRey Then
                  'WriteConsoleMsg Userindex, "Hemos notado que has aniquilado con la vida del rey oponente. ¡FELICITACIONES! Aquí tienes tu recompensa! 250.000 monedas de oro extra y su equipamiento", FontTypeNames.FONTTYPE_INFO
60                WriteShortMsj UserIndex, 55, FontTypeNames.FONTTYPE_INFO, , , , 250000
70                .Stats.Gld = .Stats.Gld + 250000
                  
80            End If
              
90            MiObj.ObjIndex = 899
100           MiObj.Amount = 1
                              
110           If Not MeterItemEnInventario(UserIndex, MiObj) Then
120               Call TirarItemAlPiso(.Pos, MiObj)
130           End If
                              
140           MiObj.ObjIndex = 900
150           MiObj.Amount = 1
                              
160           If Not MeterItemEnInventario(UserIndex, MiObj) Then
170               Call TirarItemAlPiso(.Pos, MiObj)
180           End If
              
190           WriteUpdateGold UserIndex
              
200           .Stats.TorneosGanados = .Stats.TorneosGanados + 1
210       End With
          
220   Exit Sub

error:
230       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : CastleMode_Premio()"
End Sub

' FIN EVENTO CASTLE MODE #####################################

' ###################### EVENTO DAGA RUSA ###########################
Public Sub InitDagaRusa(ByVal SlotEvent As Byte)
10    On Error GoTo error

          Dim LoopC As Integer
          Dim NpcIndex As Integer
          Dim Pos As WorldPos
          
          Dim Num As Integer
          
20        With Events(SlotEvent)
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Id > 0 Then
50                    Call WarpUserChar(.Users(LoopC).Id, 211, 21 + Num, 60, False)
60                    Num = Num + 1
70                    Call WriteUserInEvent(.Users(LoopC).Id)
80                End If
90            Next LoopC
              
100           Pos.map = 211
110           Pos.X = 21
120           Pos.Y = 59
130           NpcIndex = SpawnNpc(704, Pos, False, False)
          
140           If NpcIndex <> 0 Then
150               Npclist(NpcIndex).Movement = NpcDagaRusa
160               Npclist(NpcIndex).flags.SlotEvent = SlotEvent
170               Npclist(NpcIndex).flags.InscribedPrevio = .Inscribed
180               .NpcIndex = NpcIndex
                  
190               DagaRusa_MoveNpc NpcIndex, True
200           End If
              
              
210           .TimeCount = 10
220       End With

230   Exit Sub

error:
240       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : InitDagaRusa()"
End Sub
Public Function DagaRusa_NextUser(ByVal SlotEvent As Byte) As Byte
10    On Error GoTo error

          Dim LoopC As Integer
          
20        DagaRusa_NextUser = 0
          
30        With Events(SlotEvent)
40            For LoopC = LBound(.Users()) To UBound(.Users())
50                If (.Users(LoopC).Id > 0) And (.Users(LoopC).Value = 0) Then
60                    DagaRusa_NextUser = .Users(LoopC).Id
                      '.Users(LoopC).Value = 1
70                    Exit For
80                End If
90            Next LoopC
100       End With
              
110   Exit Function

error:
120       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : DagaRusa_NextUser()"
End Function
Public Sub DagaRusa_ResetRonda(ByVal SlotEvent As Byte)

          Dim LoopC As Integer
          
10        With Events(SlotEvent)
20            For LoopC = LBound(.Users()) To UBound(.Users())
30                .Users(LoopC).Value = 0
40            Next LoopC
          
50        End With
End Sub
Private Sub DagaRusa_CheckWin(ByVal SlotEvent As Byte)

10    On Error GoTo error

          Dim UserIndex As Integer
          Dim MiObj As Obj
          
20        With Events(SlotEvent)
30            If .Inscribed = 1 Then
40                UserIndex = SearchLastUserEvent(SlotEvent)
50                DagaRusa_Premio UserIndex
                  
                    WriteUserInEvent (UserIndex)
                    
60                Call QuitarNPC(.NpcIndex)
70                CloseEvent SlotEvent
                  
80            End If
90        End With
          
100   Exit Sub

error:
110       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : DagaRusa_CheckWin()"
End Sub

Private Sub DagaRusa_Premio(ByVal UserIndex As Integer)

10    On Error GoTo error

          Dim MiObj As Obj
          
20        With UserList(UserIndex)
30             MiObj.Amount = 1
40             MiObj.ObjIndex = 1037
              
              'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("DagaRusa» El ganador es " & UserList(Userindex).Name & ". Felicitaciones para el personaje, quien se ha ganado una MD! (Espada mata dragones)", FontTypeNames.FONTTYPE_GUILD)
50            SendData SendTarget.ToAll, 0, PrepareMessageShortMsj(56, FontTypeNames.FONTTYPE_GUILD, , , , , .Name)
              
60            If Not MeterItemEnInventario(UserIndex, MiObj) Then
70                Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
80            End If
              
90            .Stats.TorneosGanados = .Stats.TorneosGanados + 1
              
100       End With
          
110   Exit Sub

error:
120       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : DagaRusa_Premio()"
End Sub
Public Sub DagaRusa_AttackUser(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
10    On Error GoTo error

          Dim N As Integer
          Dim Slot As Byte
          
20        With UserList(UserIndex)
              
30            N = 10
              
40            If RandomNumber(1, 100) <= N Then
              
                  ' Sound
50                SendData SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y)
                  ' Fx
60                SendData SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0)
                  ' Cambio de Heading
70                ChangeNPCChar NpcIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.Head, SOUTH
                  'Apuñalada en el piso
80                SendData SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, 1000, DAMAGE_PUÑAL)
                  
90                WriteConsoleMsg UserIndex, "¡Has sido apuñalado por 1.000!", FontTypeNames.FONTTYPE_FIGHT
                  
                  WriteUserInEvent UserIndex
                  
100               Slot = .flags.SlotEvent
                  
                  
110               Call UserDie(UserIndex)
120               EventosDS.AbandonateEvent (UserIndex)
130               Call DagaRusa_CheckWin(Slot)
                 
                  
140           Else
                  ' Sound
150               SendData SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y)
                  ' Fx
160               SendData SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0)
                  ' Cambio de Heading
170               ChangeNPCChar NpcIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.Head, SOUTH

180               WriteConsoleMsg UserIndex, "¡Parece que no te he apuñalado, ya verás!", FontTypeNames.FONTTYPE_FIGHT
                 ' SendData SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, 1000, DAMAGE_PUÑAL)
190           End If
              
              
              
200       End With
210   Exit Sub

error:
220       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : DagaRusa_AttackUser()"
End Sub

' FIN EVENTO DAGA RUSA ###########################################
Private Function SelectModalityDeathMatch(ByVal SlotEvent As Byte) As Integer
          Dim Random As Integer
          
10        Randomize
20        Random = RandomNumber(1, 8)
          
30        With Events(SlotEvent)
40            Select Case Random
                  Case 1 ' Zombie
50                    .CharBody = 11
60                Case 2 ' Golem
70                    .CharBody = 11
80                Case 3 ' Araña
90                    .CharBody = 42
100               Case 4 ' Asesino
110                   .CharBody = 11 '48
120               Case 5 'Medusa suprema
130                   .CharBody = 151
140               Case 6 'Dragón azul
150                   .CharBody = 42 '247
160               Case 7 'Viuda negra 185
170                   .CharBody = 185
180               Case 8 'Tigre salvaje
190                   .CharBody = 147
200           End Select
210       End With
End Function

' DEATHMATCH ####################################################
Private Sub InitDeathMatch(ByVal SlotEvent As Byte)
10    On Error GoTo error

          Dim LoopC As Integer
          Dim Pos As WorldPos
          
20        Call SelectModalityDeathMatch(SlotEvent)
          
30        With Events(SlotEvent)
40            For LoopC = LBound(.Users()) To UBound(.Users())
50                If .Users(LoopC).Id > 0 Then
60                    .Users(LoopC).Team = LoopC
70                    .Users(LoopC).Selected = 1
                      
80                    ChangeBodyEvent SlotEvent, .Users(LoopC).Id, True
90                    UserList(.Users(LoopC).Id).showName = False
100                   RefreshCharStatus .Users(LoopC).Id
                      
                      
110                   Pos.map = 211
120                   Pos.X = RandomNumber(58, 84)
130                   Pos.Y = RandomNumber(28, 44)
                  
140                   Call ClosestLegalPos(Pos, Pos)
150                   Call WarpUserChar(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y, True)
160               End If
              
170           Next LoopC
          
180           .TimeCount = 20
190       End With
          
200   Exit Sub

error:
210       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : InitDeathMatch()"
End Sub

Public Sub DeathMatch_UserDie(ByVal SlotEvent As Byte, ByVal UserIndex As Integer)

10    On Error GoTo error

20        AbandonateEvent (UserIndex)
              
30        If Events(SlotEvent).Inscribed = 1 Then
40            UserIndex = SearchLastUserEvent(SlotEvent)
50            DeathMatch_Premio UserIndex
60            CloseEvent SlotEvent
70        End If
          
80    Exit Sub

error:
90        LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : DeathMatch_UserDie()"
End Sub
Private Sub DeathMatch_Premio(ByVal UserIndex As Integer)
10    On Error GoTo error

20        With UserList(UserIndex)
              'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("DeathMatch» El ganador es " & .Name & " quien se lleva 1 punto de torneo y 450.000 monedas de oro.", FontTypeNames.FONTTYPE_GUILD)
30            SendData SendTarget.ToAll, 0, PrepareMessageShortMsj(57, FontTypeNames.FONTTYPE_GUILD, , , , , .Name)
                  
40            .Stats.Gld = .Stats.Gld + 450000
50            WriteUpdateGold UserIndex
              
60            .Stats.TorneosGanados = .Stats.TorneosGanados + 1
70        End With
          
80    Exit Sub

error:
90        LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : DeathMatch_Premio()"
End Sub

' FIN DEATHMATCH ################################################
' EVENTOS DONDE LOS USUARIOS SE TRANSFORMAN EN CRIATURAS
Private Sub InitEventTransformation(ByVal SlotEvent As Byte, _
                                    ByVal NewBody As Integer, _
                                    ByVal NewHp As Integer, _
                                    ByVal map As Integer, _
                                    ByVal X As Byte, _
                                    ByVal Y As Byte)
          
10        On Error GoTo error
          
          Dim LoopC As Integer
          Dim UserSelected As Integer
          Dim Pos As WorldPos
          
          Const Rango As Byte = 4
          
20        With Events(SlotEvent)
30            .CharBody = NewBody
40            .CharHp = NewHp
              
50            For LoopC = LBound(.Users()) To UBound(.Users())
60                If .Users(LoopC).Id > 0 Then
70                    .Users(LoopC).Team = 2
                      
                      
80                    Pos.map = map
90                    Pos.X = RandomNumber(X - Rango, X + Rango)
100                   Pos.Y = RandomNumber(Y - Rango, Y + Rango)
                  
110                   Call ClosestLegalPos(Pos, Pos)
120                   Call WarpUserChar(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y, True)
                      
130               End If
140           Next LoopC
              
150           Transformation_SelectionUser SlotEvent
160       End With
          
170   Exit Sub

error:
180       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : InitEventTransformation()"
End Sub

Private Function Transformation_SelectionUser(ByVal SlotEvent As Byte)
10    On Error GoTo error

          Dim LoopC As Integer
20        With Events(SlotEvent)
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                Transformation_SelectionUser = RandomNumber(LBound(.Users()), UBound(.Users()))
                  
50                If .Users(Transformation_SelectionUser).Id > 0 And .Users(Transformation_SelectionUser).Selected = 0 Then
60                    Exit For
70                End If
80            Next LoopC
              
90            .Users(Transformation_SelectionUser).Selected = 1
100           .Users(Transformation_SelectionUser).Team = 1
                          
110           Call ChangeHpEvent(.Users(Transformation_SelectionUser).Id)
120           Call ChangeBodyEvent(SlotEvent, .Users(Transformation_SelectionUser).Id, IIf(.Modality = Minotauro, False, True))
130       End With
140   Exit Function

error:
150       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Transformation_SelectionUser()"
End Function

Public Sub Transformation_UserDie(ByVal UserIndex As Integer, ByVal AttackerIndex As Integer)
10    On Error GoTo error

          Dim SlotEvent As Byte
          Dim Exituser As Boolean
          
20        With UserList(UserIndex)
30            SlotEvent = .flags.SlotEvent
40            AbandonateEvent UserIndex
              
50            Transformation_CheckWin UserIndex, SlotEvent, AttackerIndex
60        End With
70    Exit Sub

error:
80        LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Transformation_UserDie()"
End Sub
Private Function Transformation_SearchUserSelected(ByVal SlotEvent As Byte) As Integer
10    On Error GoTo error

          Dim LoopC As Integer
          
20        With Events(SlotEvent)
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Id > 0 Then
50                    If .Users(LoopC).Selected = 1 Then
60                        Transformation_SearchUserSelected = LoopC
70                    End If
80                End If
90            Next LoopC
100       End With
          
110   Exit Function

error:
120       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Transformation_SearchUserSelected()"
End Function
Public Sub Transformation_CheckWin(ByVal UserIndex As Integer, ByVal SlotEvent As Byte, Optional ByVal AttackerIndex As Integer = 0)
10        On Error GoTo error
    
    ' VER LAUTARO
    Dim IsSelected As Boolean
    Dim tUser As Integer
20
30        With Events(SlotEvent)
40          If .Inscribed = 1 Then
50              tUser = SearchLastUserEvent(SlotEvent)
60
70              If .Users(UserList(tUser).flags.SlotUserEvent).Selected = 1 Then IsSelected = True
                
80              Transformation_Premio tUser, IsSelected, 250000
90
100             CloseEvent SlotEvent
110             Exit Sub
120         End If
130
        
        If AttackerIndex <> 0 Then
            'Significa que hay más de un usuario. Por lo tanto podría haber muerto el bicho transformado
140           If UserList(UserIndex).flags.SlotUserEvent = Transformation_SearchUserSelected(SlotEvent) Then
150               Transformation_Premio AttackerIndex, False, 250000
160
170                CloseEvent SlotEvent
180         End If
        End If
190       End With
    
200   Exit Sub

error:
210       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Transformation_CheckWin() at line " & Erl
End Sub

Private Sub Transformation_Premio(ByVal UserIndex As Integer, _
                                    ByVal IsSelected As Boolean, _
                                    ByVal Gld As Long)
                                    
10        On Error GoTo error
20
    Dim UserWin As Integer
    
30    With UserList(UserIndex)
        Dim SlotEvent As Byte
40        SlotEvent = .flags.SlotEvent
        
50        If IsSelected Then
60            .Stats.Gld = .Stats.Gld + (Gld * 2)
            'WriteConsoleMsg Userindex, "Has recibido " & (Gld * 2) & " por haber aniquilado a todos los usuarios.", FontTypeNames.FONTTYPE_INFO
70            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(SlotEvent, Events(SlotEvent).Modality) & "» Ha logrado derrotar a todos los participantes. Felicitaciones para " & .Name & " quien fue escogido como " & strModality(SlotEvent, Events(SlotEvent).Modality), FontTypeNames.FONTTYPE_GUILD)
80            WriteShortMsj UserIndex, 58, FontTypeNames.FONTTYPE_INFO, , , , (Gld * 2)

90      Else
100           .Stats.Gld = .Stats.Gld + Gld
              'WriteConsoleMsg Userindex, "Has recibido " & Gld & " por haber aniquilado a " & strModality(SlotEvent, Events(SlotEvent).Modality), FontTypeNames.FONTTYPE_INFO
110           WriteShortMsj UserIndex, 59, FontTypeNames.FONTTYPE_INFO, , , , Gld, strModality(SlotEvent, Events(SlotEvent).Modality)
120           SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(SlotEvent, Events(SlotEvent).Modality) & "» Felicitaciones para " & .Name & " quien derrotó a " & strModality(SlotEvent, Events(SlotEvent).Modality), FontTypeNames.FONTTYPE_GUILD)

130     End If
        
140        WriteUpdateGold UserIndex
        
150        .Stats.TorneosGanados = .Stats.TorneosGanados + 1
    
160       End With
    
170   Exit Sub

error:
180       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Transformation_Premio() AT LINE: " & Erl
End Sub


' FIN EVENTOS DONDE LOS USUARIOS SE TRANSFORMAN EN CRIATURAS

' ARACNUS #######################################################

Public Sub Aracnus_Veneno(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)

10    On Error GoTo error
          ' El personaje transformado en Aracnus, tiene 10% de probabilidad de envenenar a la víctima y dejarla fuera del torneo.

          Const N As Byte = 10
          
20        With UserList(AttackerIndex)
30            If RandomNumber(1, 100) <= 10 Then
                  'WriteConsoleMsg Victimindex, "Has sido envenenado por Aracnus, has muerto de inmediato por su veneno letal.", FontTypeNames.FONTTYPE_FIGHT
40                WriteShortMsj VictimIndex, 60, FontTypeNames.FONTTYPE_FIGHT
50                Call UserDie(VictimIndex)
                  
60                Transformation_CheckWin VictimIndex, .flags.SlotEvent, AttackerIndex
70            End If
          
80        End With
          
90    Exit Sub

error:
100       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Aracnus_Veneno()"
End Sub

Public Sub Minotauro_Veneno(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)
10    On Error GoTo error

          ' El personaje transformado en Minotauro, tiene 10% de posibilidad de dar un golpe mortal
          Const N As Byte = 10
          
20        With UserList(AttackerIndex)
30            If RandomNumber(1, 100) <= 10 Then
                  'WriteConsoleMsg Victimindex, "¡El minotauro ha logrado paralizar tu cuerpo con su dosis de veneno. Has quedado afuera del evento.", FontTypeNames.FONTTYPE_FIGHT
40                WriteShortMsj VictimIndex, 61, FontTypeNames.FONTTYPE_FIGHT
50                Call UserDie(VictimIndex)
                  
60                Transformation_CheckWin VictimIndex, .flags.SlotEvent, AttackerIndex
              
70            End If
          
80        End With
          
90    Exit Sub

error:
100       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Minotauro_Veneno()"
End Sub

' FIN ARACNUS ###################################################

' EVENTO BUSQUEDA '
Private Sub InitBusqueda(ByVal SlotEvent As Byte)
10    On Error GoTo error

          
          Dim LoopC As Integer
          Dim Pos As WorldPos
          
20        With Events(SlotEvent)
30            For LoopC = 1 To 20
40                Busqueda_CreateObj 216, RandomNumber(20, 80), RandomNumber(20, 80)
50            Next LoopC
              
60            For LoopC = LBound(.Users()) To UBound(.Users())
70                If .Users(LoopC).Id > 0 Then
80                    Pos.map = 216
90                    Pos.X = RandomNumber(50, 60)
100                   Pos.Y = RandomNumber(50, 60)
                      
110                   Call ClosestLegalPos(Pos, Pos)
120                   Call WarpUserChar(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y, True)
130               End If
140           Next LoopC
              
150           .TimeFinish = 60
          
160       End With
          
170   Exit Sub

error:
180       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : InitBusqueda()"
End Sub

Private Sub Busqueda_CreateObj(ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)
10    On Error GoTo error

          ' Creamos un objeto en el mapa de búsqueda.
          
          Dim Pos As WorldPos
          Dim Obj As Obj
          
20        Pos.map = map
30        Pos.X = X
40        Pos.Y = Y
50        ClosestStablePos Pos, Pos
          
60        Obj.ObjIndex = 1037
70        Obj.Amount = 1
80        Call MakeObj(Obj, Pos.map, Pos.X, Pos.Y)
90        MapData(Pos.map, Pos.X, Pos.Y).ObjEvent = 1
          
100   Exit Sub

error:
110       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Busqueda_CreateObj()"
End Sub
Private Sub Busqueda_SearchWin(ByVal SlotEvent As Byte)
10    On Error GoTo error

          Dim UserIndex As Integer
          Dim CopyUsers() As tUserEvent
          
20        With Events(SlotEvent)
30             Event_OrdenateUsersValue SlotEvent, CopyUsers
              
40            UserIndex = CopyUsers(1).Id
              
50            If UserIndex > 0 Then
60                UserList(UserIndex).Stats.TorneosGanados = UserList(UserIndex).Stats.TorneosGanados + 1
70                UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld + 350000
80                WriteUpdateGold UserIndex
                  
                  ' vercoso este userindex 0
90                SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Busqueda de objetos» El ganador de la búsqueda de objetos es " & UserList(UserIndex).Name & ". Felicitaciones! Se lleva como premio 350.000 monedas de oro." & vbCrLf & _
                      "Tabla final de posiciones: " & Event_GenerateTablaPos(SlotEvent, CopyUsers), FontTypeNames.FONTTYPE_GUILD)
                   'SendData SendTarget.ToAll, 0, PrepareMessageShortMsj(62, FontTypeNames.FONTTYPE_GUILD, , , , , UserList(Userindex).Name)
100           End If
              
110           CloseEvent SlotEvent
              
120       End With
          
130   Exit Sub

error:
140       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Busqueda_SearchWin()"
End Sub
Private Function Busqueda_UserRecolectedObj(ByVal SlotEvent As Byte) As Integer
10    On Error GoTo error

          Dim LoopC As Integer
          
20        With Events(SlotEvent)
30            For LoopC = LBound(.Users()) To UBound(.Users())
                  
40                If .Users(LoopC).Id > 0 Then
50                    If Busqueda_UserRecolectedObj = 0 Then Busqueda_UserRecolectedObj = LoopC
60                    If .Users(LoopC).Value > .Users(Busqueda_UserRecolectedObj).Value Then
70                        Busqueda_UserRecolectedObj = LoopC
80                    End If
90                End If
                      
100           Next LoopC
              
110           Busqueda_UserRecolectedObj = .Users(Busqueda_UserRecolectedObj).Id
120       End With
          
130   Exit Function

error:
140       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Busqueda_UserRecolectedObj()"
End Function

Public Sub Busqueda_GetObj(ByVal SlotEvent As Byte, ByVal SlotUserEvent As Byte)
10    On Error GoTo error

20        With Events(SlotEvent)
30            .Users(SlotUserEvent).Value = .Users(SlotUserEvent).Value + 1
              
              'WriteConsoleMsg .Users(SlotUserEvent).Id, "Has recolectado un objeto del piso. En total llevas " & .Users(SlotUserEvent).value & " objetos recolectados. Sigue así!", FontTypeNames.FONTTYPE_INFO
40            WriteShortMsj .Users(SlotUserEvent).Id, 63, FontTypeNames.FONTTYPE_INFO, .Users(SlotUserEvent).Value
50            Busqueda_CreateObj 216, RandomNumber(30, 80), RandomNumber(30, 80)
60        End With
70    Exit Sub

error:
80        LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Busqueda_GetObj()"
End Sub

' ENFRENTAMIENTOS ###############################################

Private Sub InitFights(ByVal SlotEvent As Byte)
10    On Error GoTo error
          
20        With Events(SlotEvent)
30            Fight_SelectedTeam SlotEvent
40            Fight_Combate SlotEvent
50        End With
60    Exit Sub

error:
70        LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : InitFights()"
End Sub
Private Sub Fight_SelectedTeam(ByVal SlotEvent As Byte)
          
10    On Error GoTo error

          ' En los enfrentamientos utilizamos este procedimiento para seleccionar los grupos o bien el usuario queda solo por 1vs1.
          Dim LoopX As Integer
          Dim LoopY As Integer
          Dim Team As Byte
          
20        Team = 1
          
30        With Events(SlotEvent)
40            For LoopX = LBound(.Users()) To UBound(.Users()) Step .TeamCant
50                For LoopY = 0 To (.TeamCant - 1)
60                    .Users(LoopX + LoopY).Team = Team
70                Next LoopY
                  
80                Team = Team + 1
90            Next LoopX
          
100       End With
          
110   Exit Sub

error:
120       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Fight_SelectedTeam()"
End Sub

Private Sub Fight_WarpTeam(ByVal SlotEvent As Byte, _
                                        ByVal ArenaSlot As Byte, _
                                        ByVal TeamEvent As Byte, _
                                        ByVal IsContrincante As Boolean, _
                                        ByRef StrTeam As String)

10    On Error GoTo error

          Dim LoopC As Integer
          Dim strTemp As String, strTemp1 As String, strTemp2 As String
          
20        With Events(SlotEvent)
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Id > 0 And .Users(LoopC).Team = TeamEvent Then
50                    If IsContrincante Then
60                        Call EventWarpUser(.Users(LoopC).Id, MapEvent.Fight(ArenaSlot).map, MapEvent.Fight(ArenaSlot).X + MAP_TILE_VS, MapEvent.Fight(ArenaSlot).Y + MAP_TILE_VS)
                          
                          ' / Update color char team
70                        UserList(.Users(LoopC).Id).flags.FightTeam = 2
                          
80                        RefreshCharStatus (.Users(LoopC).Id)
90                    Else
100                       Call EventWarpUser(.Users(LoopC).Id, MapEvent.Fight(ArenaSlot).map, MapEvent.Fight(ArenaSlot).X, MapEvent.Fight(ArenaSlot).Y)
                          
                          ' / Update color char team
110                       UserList(.Users(LoopC).Id).flags.FightTeam = 1
120                       RefreshCharStatus (.Users(LoopC).Id)
130                   End If
                      
140                   If StrTeam = vbNullString Then
150                       StrTeam = UserList(.Users(LoopC).Id).Name
160                   Else
170                       StrTeam = StrTeam & "-" & UserList(.Users(LoopC).Id).Name
180                   End If
                      
190                   .Users(LoopC).Value = 1
200                   .Users(LoopC).MapFight = ArenaSlot
                      
210                   UserList(.Users(LoopC).Id).Counters.TimeFight = 10
220                   Call WriteUserInEvent(.Users(LoopC).Id)
230               End If
240           Next LoopC
250       End With
          
260   Exit Sub

error:
270       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Fight_WarpTeam()"
End Sub

Private Function Fight_Search_Enfrentamiento(ByVal UserIndex As Integer, ByVal UserTeam As Byte, ByVal SlotEvent As Byte) As Byte
10    On Error GoTo error

          ' Chequeamos que tengamos contrincante para luchar.
          Dim LoopC As Integer
          
20        Fight_Search_Enfrentamiento = 0
          
30        With Events(SlotEvent)
40            For LoopC = LBound(.Users()) To UBound(.Users())
50                If .Users(LoopC).Id > 0 And .Users(LoopC).Value = 0 Then
60                    If .Users(LoopC).Id <> UserIndex And .Users(LoopC).Team <> UserTeam Then
70                        Fight_Search_Enfrentamiento = .Users(LoopC).Team
80                        Exit For
90                    End If
100               End If
110           Next LoopC
          
120       End With
          
130   Exit Function

error:
140       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Fight_Search_Enfrentamiento()"
End Function

Private Sub NewRound(ByVal SlotEvent As Byte)
          Dim LoopC As Long
          Dim Count As Long
          
10        With Events(SlotEvent)
20            Count = 0
              
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Id > 0 Then
                      ' Hay esperando
50                    If .Users(LoopC).Value = 0 Then
60                        Exit Sub
70                    End If
                      
                      ' Hay luchando
80                    If .Users(LoopC).MapFight > 0 Then
90                        Exit Sub
100                   End If
110               End If
120           Next LoopC
              
130           For LoopC = LBound(.Users()) To UBound(.Users())
140               .Users(LoopC).Value = 0
150           Next LoopC

            LogEventos "Se reinicio la informacion de los fights()"
              
160       End With
End Sub
Private Sub Fight_Combate(ByVal SlotEvent As Byte)
10    On Error GoTo error

          ' Buscamos una arena disponible y mandamos la mayor cantidad de usuarios disponibles.
          Dim LoopC As Integer
          Dim FreeArena As Byte
          Dim OponentTeam As Byte
          Dim strTemp As String
          Dim strTeam1 As String
          Dim strTeam2 As String
          
20        With Events(SlotEvent)
cheking:
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Id > 0 And .Users(LoopC).Value = 0 Then
50                    FreeArena = FreeSlotArena()
                      
60                    If FreeArena > 0 Then
70                        OponentTeam = Fight_Search_Enfrentamiento(.Users(LoopC).Id, .Users(LoopC).Team, SlotEvent)
                          
80                        If OponentTeam > 0 Then
90                            StatsEvent .Users(LoopC).Id
100                           Fight_WarpTeam SlotEvent, FreeArena, .Users(LoopC).Team, False, strTeam1
110                           Fight_WarpTeam SlotEvent, FreeArena, OponentTeam, True, strTeam2
120                           MapEvent.Fight(FreeArena).Run = True
                              
130                           strTemp = "Duelos " & Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant & "» "
140                           strTemp = strTemp & strTeam1 & " vs " & strTeam2
150                           SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strTemp, FontTypeNames.FONTTYPE_GUILD)
                              
160                           strTemp = vbNullString
170                           strTeam1 = vbNullString
180                           strTeam2 = vbNullString
                              
190                       Else
                              ' Pasa de ronda automaticamente
200                           .Users(LoopC).Value = 1
210                           WriteConsoleMsg .Users(LoopC).Id, "Hemos notado que no tienes un adversario. Pasaste a la siguiente ronda.", FontTypeNames.FONTTYPE_INFO
220                           NewRound SlotEvent
                              GoTo cheking:
230                       End If
240                   End If
250               End If
260           Next LoopC
              
270       End With
          
280   Exit Sub

error:
290       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Fight_Combate()"
End Sub
Private Sub ResetValue(ByVal SlotEvent As Byte)
          Dim LoopC As Integer
          
10        With Events(SlotEvent)
20            For LoopC = LBound(.Users()) To UBound(.Users())
30                .Users(LoopC).Value = 0
40            Next LoopC
50        End With
End Sub
Private Function CheckTeam_UserDie(ByVal SlotEvent As Integer, ByVal TeamUser As Byte) As Boolean

10    On Error GoTo error

          Dim LoopC As Integer
          ' Encontramos a uno del Team vivo, significa que no hay terminación del duelo.
          
          
20        With Events(SlotEvent)
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Id > 0 Then
50                    If .Users(LoopC).Team = TeamUser Then
60                        If UserList(.Users(LoopC).Id).flags.Muerto = 0 Then
70                            CheckTeam_UserDie = False
80                            Exit Function
90                        End If
100                   End If
110               End If
120           Next LoopC
              
130           CheckTeam_UserDie = True
          
140       End With
          
150   Exit Function

error:
160       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : CheckTeam_UserDie()"
End Function
Private Sub Team_UserDie(ByVal SlotEvent As Byte, ByVal TeamSlot As Byte)
10    On Error GoTo error

          Dim LoopC As Integer
20        With Events(SlotEvent)
              
              
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Id > 0 Then
50                    If .Users(LoopC).Team = TeamSlot Then
60                        AbandonateEvent .Users(LoopC).Id
70                    End If
80                End If
90            Next LoopC
          
100       End With
          
110   Exit Sub

error:
120       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Team_UserDie()"
End Sub
Public Function Fight_CheckContinue(ByVal UserIndex As Integer, ByVal SlotEvent As Byte, ByVal TeamSlot As Byte) As Boolean
          ' Esta función devuelve un TRUE cuando el enfrentamiento puede seguir.
          
          Dim LoopC As Integer, cant As Integer
          
10        With Events(SlotEvent)
              
20            Fight_CheckContinue = False
              
30            For LoopC = LBound(.Users()) To UBound(.Users())
                  ' User válido
40                If .Users(LoopC).Id > 0 And .Users(LoopC).Id <> UserIndex Then
50                    If .Users(LoopC).Team = TeamSlot Then
60                        If UserList(.Users(LoopC).Id).flags.Muerto = 0 Then
70                            Fight_CheckContinue = True
80                            Exit For
90                        End If
100                   End If
110               End If
120           Next LoopC

130       End With
          
140   Exit Function

error:
150       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Team_CheckContinue()"
End Function
Public Sub Fight_WinForzado(ByVal UserIndex As Integer, ByVal SlotEvent As Byte, ByVal MapFight As Byte)
10        On Error GoTo error
          
          Dim LoopC As Integer
          Dim strTempWin As String
          Dim TeamWin As Byte
          
20        With Events(SlotEvent)

              LogEventos "El personaje " & UserList(UserIndex).Name & " deslogeó en lucha."
              
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                With .Users(LoopC)
50                    If .Id > 0 And UserIndex <> .Id Then
60                        If .MapFight = MapFight Then
70                            If strTempWin = vbNullString Then
80                                strTempWin = UserList(.Id).Name
90                            Else
100                               strTempWin = strTempWin & "-" & UserList(.Id).Name
110                           End If
                              
                              '.value = 0
130                           .MapFight = 0
                              
140                           EventWarpUser .Id, 211, 30, 21
                              'WriteConsoleMsg .Id, "Felicitaciones. Has ganado el enfrentamiento", FontTypeNames.FONTTYPE_INFO
                              LogEventos "El personaje " & UserList(.Id).Name & " ha ganado el enfrentamiento"
                              
150                           WriteShortMsj .Id, 64, FontTypeNames.FONTTYPE_INFO

                              ' / Update color char team
160                           UserList(.Id).flags.FightTeam = 0
170                           RefreshCharStatus (.Id)
180                           TeamWin = .Team
190                       End If
200                   End If
210               End With
220           Next LoopC

              MapEvent.Fight(MapFight).Run = False
              
              
230           If strTempWin <> vbNullString Then SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Duelos " & Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant & "» Duelo ganado por " & strTempWin & ".", FontTypeNames.FONTTYPE_GUILD)
              
              ' Nos fijamos si resetea el Value
240           Call NewRound(SlotEvent)
              
              ' Nos fijamos si eran los últimos o si podemos mandar otro combate..
250           If TeamCant(SlotEvent, TeamWin) = .Inscribed Then
260               Fight_SearchTeamWin SlotEvent, TeamWin
270               CloseEvent SlotEvent
280           Else
290               Fight_Combate SlotEvent
300           End If
          
310       End With
          
320   Exit Sub

error:
330       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Fight_WinForzado()"
End Sub
Private Sub StatsEvent(ByVal UserIndex As Integer)
10    On Error GoTo error

20        With UserList(UserIndex)
30            If .flags.Muerto Then
40                Call RevivirUsuario(UserIndex)
50                Exit Sub
60            End If
              
70            .Stats.MinHp = .Stats.MaxHp
80            .Stats.MinMAN = .Stats.MaxMAN
90            .Stats.MinAGU = 100
100           .Stats.MinHam = 100
              
110           WriteUpdateUserStats UserIndex
          
120       End With
          
130   Exit Sub

error:
140       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : StatsEvent()"
End Sub

Private Function SearchTeamAttacker(ByVal TeamUser As Byte)

End Function
Public Sub Fight_UserDie(ByVal SlotEvent As Byte, ByVal SlotUserEvent As Byte, ByVal AttackerIndex As Integer)
10    On Error GoTo error
    Dim TeamSlot As Byte
    Dim LoopC As Integer
    Dim strTempWin As String
    Dim TeamWin As Byte
    Dim MapFight As Byte
    
    ' Aca se hace que el que gané no siga luchando sino que espere.
    
20    With Events(SlotEvent)
30        TeamSlot = .Users(SlotUserEvent).Team
40        TeamWin = .Users(UserList(AttackerIndex).flags.SlotUserEvent).Team
        
50        If CheckTeam_UserDie(SlotEvent, TeamSlot) = False Then Exit Sub
        
60        For LoopC = LBound(.Users()) To UBound(.Users())
70            If .Users(LoopC).Id > 0 Then
80                    With .Users(LoopC)
90                        If .Team = TeamWin Then
100                           StatsEvent .Id
110
120                            If strTempWin = vbNullString Then
130                                strTempWin = UserList(.Id).Name
140                            Else
150                               strTempWin = strTempWin & "-" & UserList(.Id).Name
160                         End If
                            
                            
                            MapFight = .MapFight
170
                            
                            '.value = 0
180                            .MapFight = 0
190                            EventWarpUser .Id, 211, 30, 21
                               'WriteConsoleMsg .Id, "Felicitaciones. Has ganado el enfrentamiento", FontTypeNames.FONTTYPE_INFO
200                            WriteShortMsj .Id, 64, FontTypeNames.FONTTYPE_INFO
                           
                            ' / Update color char team
210                            UserList(.Id).flags.FightTeam = 0
220                           RefreshCharStatus (.Id)
230                     End If
240                 End With
250             End If
260     Next LoopC
        
        MapEvent.Fight(MapFight).Run = False
        
        ' Abandono del user/team
270     Team_UserDie SlotEvent, TeamSlot
        
280     If strTempWin <> vbNullString Then SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Duelos " & Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant & "» Enfrentamiento ganado por " & strTempWin & ".", FontTypeNames.FONTTYPE_GUILD)
        
        ' // Se fija de poder pasar a la siguiente ronda o esperar a los combates que faltan.
290     Call NewRound(SlotEvent)
        
        ' Si la cantidad es igual al inscripto quedó final.
300     If TeamCant(SlotEvent, TeamWin) = .Inscribed Then
310            Fight_SearchTeamWin SlotEvent, TeamWin
320            CloseEvent SlotEvent
330     Else
340            Fight_Combate SlotEvent
350     End If
        
360       End With
    
370   Exit Sub

error:
380       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Fight_UserDie()" & " AT LINE: " & Erl
End Sub
Private Function TeamCant(ByVal SlotEvent As Byte, ByVal TeamSlot As Byte) As Byte

10    On Error GoTo error
          ' Devuelve la cantidad de miembros que tiene un clan
          Dim LoopC As Integer
          
20        TeamCant = 0
          
30        With Events(SlotEvent)
40            For LoopC = LBound(.Users()) To UBound(.Users())
50                If .Users(LoopC).Team = TeamSlot Then
60                    TeamCant = TeamCant + 1
70                End If
80            Next LoopC
90        End With
          
100   Exit Function

error:
110       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : TeamCant()"
End Function
Private Sub Fight_SearchTeamWin(ByVal SlotEvent As Byte, ByVal TeamWin As Byte)

10    On Error GoTo error

          Dim LoopC As Integer
          Dim strTemp As String
          Dim strReWard As String
          
          
20        With Events(SlotEvent)
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Id > 0 And .Users(LoopC).Team = TeamWin Then
                      'riteConsoleMsg .Users(LoopC).Id, "Has ganado el evento. ¡Felicitaciones!", FontTypeNames.FONTTYPE_INFO
50                    WriteShortMsj .Users(LoopC).Id, 65, FontTypeNames.FONTTYPE_INFO
                      
60                    PrizeUser .Users(LoopC).Id, False
                      
70                    If strTemp = vbNullString Then
80                        strTemp = UserList(.Users(LoopC).Id).Name
90                    Else
100                       strTemp = strTemp & ", " & UserList(.Users(LoopC).Id).Name
110                   End If
120               End If
130           Next LoopC
          
          
140       If .TeamCant > 1 Then
150           If .GldInscription > 0 Or .DspInscription > 0 Then strReWard = "Los participantes han recibido "
160           If .GldInscription > 0 Then strReWard = strReWard & .GldInscription * .Quotas & " Monedas de oro. "
170           If .DspInscription > 0 Then strReWard = strReWard & .DspInscription * .Quotas & " Monedas DSP. "
              
180           SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Duelos " & .TeamCant & "vs" & .TeamCant & _
                  "» Evento terminado. Felicitamos a " & strTemp & " por haber ganado el torneo." & vbCrLf & strReWard, FontTypeNames.FONTTYPE_PREMIUM)
190       Else
200           If .GldInscription > 0 Or .DspInscription > 0 Then strReWard = "El participante recibió "
210           If .GldInscription > 0 Then strReWard = strReWard & .GldInscription * .Quotas & " Monedas de oro"
220           If .DspInscription > 0 Then strReWard = strReWard & " y " & .DspInscription * .Quotas & " Monedas DSP."
              
230           SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Duelos " & .TeamCant & "vs" & .TeamCant & "» Evento terminado. Felicitamos a " & strTemp & _
                  " por haber ganado el evento." & vbCrLf & strReWard, FontTypeNames.FONTTYPE_PREMIUM)
240       End If
          
250       End With
          
260   Exit Sub

error:
270       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Fight_SearchTeamWin()"
End Sub


' ############################## USUARIO UNSTOPPABLE ###########################################
Public Sub InitUnstoppable(ByVal SlotEvent As Byte)
10    On Error GoTo error

          Dim LoopC As Integer
          
20        With Events(SlotEvent)
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Id > 0 Then
50                    EventWarpUser .Users(LoopC).Id, 218, RandomNumber(30, 54), RandomNumber(25, 39)
                      
60                End If
70            Next LoopC
              
80            .TimeCount = 10
90            .TimeFinish = 60 + .TimeCount
100       End With
          
110   Exit Sub

error:
120       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : InitUnstoppable()"
End Sub
Public Sub Unstoppable_Userdie(ByVal SlotEvent As Byte, ByVal VictimSlot As Byte, ByVal AttackerSlot As Byte)
10    On Error GoTo error

20        With Events(SlotEvent)
30            With .Users(VictimSlot)
40                Call EventWarpUser(.Id, 218, RandomNumber(30, 54), RandomNumber(25, 39))
50                Call RevivirUsuario(.Id)
                  'Call WriteConsoleMsg(.Id, "Has sido aniquilado. Pero no pierdas las esperanzas joven guerrero, reviviste y tu sangre está hambrienta, ve trás el que te asesino y haz justicia!", FontTypeNames.FONTTYPE_FIGHT)
60                Call WriteShortMsj(.Id, 66, FontTypeNames.FONTTYPE_FIGHT)
70            End With
              
80            With .Users(AttackerSlot)
90                .Value = .Value + 1
100               WriteShortMsj .Id, 67, FontTypeNames.FONTTYPE_FIGHT, .Value
                  'WriteConsoleMsg .Id, "Felicitaciones, has sumado una muerte más a tu lista. Actualmente llevas " & .value & " asesinatos. Sigue así y ganarás el evento.", FontTypeNames.FONTTYPE_INFO
110           End With
120       End With
130   Exit Sub

error:
140       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Unstoppable_Userdie()"
End Sub

Private Function Event_GenerateTablaPos(ByVal SlotEvent As Byte, ByRef CopyUsers() As tUserEvent) As String
          Dim LoopC As Integer
          
10        With Events(SlotEvent)
20            For LoopC = LBound(.Users()) To UBound(.Users())
30                If CopyUsers(LoopC).Id > 0 Then
40                    Event_GenerateTablaPos = Event_GenerateTablaPos & _
                          LoopC & "° »» " & UserList(CopyUsers(LoopC).Id).Name & " (" & CopyUsers(LoopC).Value & ")" & vbCrLf
50                End If
60            Next LoopC
70        End With
          
End Function
Private Sub Unstoppable_UserWin(ByVal SlotEvent As Byte)

10    On Error GoTo error

          Dim UserIndex As Integer
          Dim strTemp As String
          Dim CopyUsers() As tUserEvent
          
20        Event_OrdenateUsersValue SlotEvent, CopyUsers
          
30        UserIndex = CopyUsers(1).Id
          
40        With UserList(UserIndex)
50            WriteShortMsj UserIndex, 68, FontTypeNames.FONTTYPE_GUILD, Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).Value
              'WriteConsoleMsg Userindex, "Felicitaciones. Tus " & Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).value & " asesinatos han hecho que ganes el evento. Aquí tienes 500.000 monedas de oro como recompensa.", FontTypeNames.FONTTYPE_INFO
60            .Stats.Gld = .Stats.Gld + 350000
70            .Stats.TorneosGanados = .Stats.TorneosGanados + 1
80            WriteUpdateGold UserIndex

90            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Usuario Unstoppable» El ganador del evento es " & .Name & " con " & _
                  Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).Value & " asesinatos." & vbCrLf & _
                  "Tabla de posiciones: " & Event_GenerateTablaPos(SlotEvent, CopyUsers), FontTypeNames.FONTTYPE_GUILD)
                  
              'SendData SendTarget.ToAll, 0, PrepareMessageShortMsj(69, FontTypeNames.FONTTYPE_GUILD, Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).value, , , , Event_GenerateTablaPos)
100           CloseEvent SlotEvent
110       End With
          
120   Exit Sub

error:
130       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Unstoppable_UserWin()"
End Sub
Private Sub Event_OrdenateUsersValue(ByVal SlotEvent As Byte, ByRef CopyUsers() As tUserEvent)

10    On Error GoTo error

    ' Utilizados para buscar ganador según VALUE
    Dim LoopX As Integer
    Dim LoopY As Integer
    Dim aux As tUserEvent
   ' Dim CopyUsers() As tUserEvent
    
20        With Events(SlotEvent)
        ' Utilizamos la copia para no dañar archivos originales
30        ReDim CopyUsers(LBound(.Users()) To UBound(.Users())) As tUserEvent
        
40        For LoopY = LBound(.Users()) To UBound(.Users())
50            CopyUsers(LoopY) = .Users(LoopY)
60        Next LoopY
        
70        For LoopX = LBound(.Users()) To UBound(.Users())
80            For LoopY = LBound(.Users()) To UBound(.Users()) - 1
90                If .Users(LoopY).Id > 0 Then
100                   If Not LoopX = UBound(.Users()) Then
110                       If CopyUsers(LoopY).Value < CopyUsers(LoopY + 1).Value Then
                            
120                           aux = CopyUsers(LoopY)
                            
130                           CopyUsers(LoopY) = CopyUsers(LoopY + 1)
140                           CopyUsers(LoopY + 1) = aux
150                     End If
160                 End If
170             End If
180         Next LoopY
190     Next LoopX
        
200       End With
    
210   Exit Sub

error:
220       LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Event_OrdenateUsersValue()"
End Sub

