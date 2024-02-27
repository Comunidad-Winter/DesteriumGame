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
    value As Integer
    Selected As Byte
    MapFight As Byte
End Type


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
    With MapEvent
        .Fight(1).Run = False
        .Fight(1).map = 217
        .Fight(1).X = 16 '+17
        .Fight(1).Y = 12 '+17
        
        .Fight(2).Run = False
        .Fight(2).map = 217
        .Fight(2).X = 16 '+17
        .Fight(2).Y = 41 '+17
        
        .Fight(3).Run = False
        .Fight(3).map = 217
        .Fight(3).X = 16 '+17
        .Fight(3).Y = 68 '+17
        
        .Fight(4).Run = False
        .Fight(4).map = 217
        .Fight(4).X = 46 '+17
        .Fight(4).Y = 12 '+17
    
    
    
    End With
End Sub
'/MANEJO DE LOS TIEMPOS '/
Public Sub LoopEvent()
    Dim LoopC As Long
    Dim LoopY As Integer
    
    For LoopC = 1 To MAX_EVENT_SIMULTANEO
        With Events(LoopC)
            If .Enabled Then
                If .TimeInit > 0 Then
                    .TimeInit = .TimeInit - 1
                        
                    Select Case .TimeInit
                        Case 0
                            
                        Case 60
                            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(.Modality) & "» Las inscripciones abren en " & Int(.TimeInit / 60) & " minutos.", FontTypeNames.FONTTYPE_GUILD)
                        Case 120
                            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(.Modality) & "» Las inscripciones abren en " & Int(.TimeInit / 60) & " minutos.", FontTypeNames.FONTTYPE_GUILD)
                        Case 180
                            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(.Modality) & "» Las inscripciones abren en " & Int(.TimeInit / 60) & " minutos.", FontTypeNames.FONTTYPE_GUILD)
                        
                    End Select
                    
                    If .TimeInit <= 0 Then
                        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(.Modality) & "» Inscripciones abiertas. /INGRESAR " & strModality(.Modality) & " para ingresar al evento.", FontTypeNames.FONTTYPE_GUILD)
                        .TimeCancel = 0
                    End If
                    
                
                End If
                
                If .TimeCancel > 0 Then
                    .TimeCancel = .TimeCancel - 1
                    
                    If .TimeCancel <= 0 Then
                        'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(.Modality) & "» Ha sido cancelado ya que no se completaron los cupos.", FontTypeNames.FONTTYPE_WARNING)
                        EventosDS.CloseEvent LoopC, "Evento " & strModality(.Modality) & " cancelado."
                    End If
                End If
                
                If .TimeCount > 0 Then
                    .TimeCount = .TimeCount - 1
                    
                    For LoopY = LBound(.Users()) To UBound(.Users())
                        If .Users(LoopY).Id > 0 Then
                            If .TimeCount = 0 Then
                                WriteConsoleMsg .Users(LoopY).Id, "Mmmm que comience el juego!", FontTypeNames.FONTTYPE_GUILD
                            Else
                                WriteConsoleMsg .Users(LoopY).Id, "Cuenta» " & .TimeCount, FontTypeNames.FONTTYPE_GUILD
                            End If
                        End If
                    Next LoopY
                End If
                
                If .NpcIndex > 0 Then
                   If Events(Npclist(.NpcIndex).flags.SlotEvent).TimeCount > 0 Then Exit Sub
                   Call DagaRusa_MoveNpc(.NpcIndex)
                End If
                
                If .TimeFinish > 0 Then
                    .TimeFinish = .TimeFinish - 1
                    
                    If .TimeFinish = 0 Then
                        Call FinishEvent(LoopC)
                    End If
                End If
            End If
    
    
        End With
    Next LoopC
End Sub

'/ FIN MANEJO DE LOS TIEMPOS


'// Funciones generales '//
Private Function FreeSlotEvent() As Byte
    Dim LoopC As Integer
    
    For LoopC = 1 To MAX_EVENT_SIMULTANEO
        If Not Events(LoopC).Enabled Then
            FreeSlotEvent = LoopC
            Exit For
        End If
    Next LoopC
End Function

Private Function FreeSlotUser(ByVal SlotEvent As Byte) As Byte
    Dim LoopC As Integer
    
    With Events(SlotEvent)
        For LoopC = 1 To MAX_USERS_EVENT
            If .Users(LoopC).Id = 0 Then
                FreeSlotUser = LoopC
                Exit For
            End If
        Next LoopC
    End With
    
End Function

Private Function FreeSlotArena() As Byte
    Dim LoopC As Integer
    
    FreeSlotArena = 0
    
    For LoopC = 1 To MAX_MAP_FIGHT
        If MapEvent.Fight(LoopC).Run = False Then
            FreeSlotArena = LoopC
            Exit For
        End If
    Next LoopC
End Function
Public Function strUsersEvent(ByVal SlotEvent As Byte) As String

    ' Texto que marca los personajes que están en el evento.
    Dim LoopC As Integer
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                strUsersEvent = strUsersEvent & UserList(.Users(LoopC).Id).Name & "-"
            End If
        Next LoopC
    End With
End Function
Private Function CheckAllowedClasses(ByRef AllowedClasses() As Byte) As String
    Dim LoopC As Integer
    
    For LoopC = 1 To NUMCLASES
        If AllowedClasses(LoopC) = 1 Then
            If CheckAllowedClasses = vbNullString Then
                CheckAllowedClasses = ListaClases(LoopC)
            Else
                CheckAllowedClasses = CheckAllowedClasses & ", " & ListaClases(LoopC)
            End If
        End If
    Next LoopC
    
End Function

Private Function SearchLastUserEvent(ByVal SlotEvent As Byte) As Integer

    ' Busca el último usuario que está en el torneo. En todos los eventos será el ganador.
    
    Dim LoopC As Integer
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                SearchLastUserEvent = .Users(LoopC).Id
                Exit For
            End If
        Next LoopC
    End With
End Function

Private Function SearchSlotEvent(ByVal Modality As eModalityEvent) As Byte
    Dim LoopC As Integer
    
    SearchSlotEvent = 0
    
    For LoopC = 1 To MAX_EVENT_SIMULTANEO
        With Events(LoopC)
            If .Modality = Modality Then
                SearchSlotEvent = LoopC
                Exit For
            End If
        End With
    Next LoopC

End Function

Private Sub EventWarpUser(ByVal Userindex As Integer, ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)

    ' Teletransportamos a cualquier usuario que cumpla con la regla de estar en un evento.
    
    Dim Pos As WorldPos
    
    With UserList(Userindex)
        Pos.map = map
        Pos.X = X
        Pos.Y = Y
        
        ClosestStablePos Pos, Pos
        WarpUserChar Userindex, Pos.map, Pos.X, Pos.Y, False
    
    End With
End Sub
Private Sub ResetEvent(ByVal Slot As Byte)
    Dim LoopC As Integer
    
    With Events(Slot)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                AbandonateEvent .Users(LoopC).Id, False

            End If
        Next LoopC
        
        If .NpcIndex > 0 Then Call QuitarNPC(.NpcIndex)
        
        .Enabled = False
        .Run = False
        .npcUserIndex = 0
        .TimeFinish = 0
        .TeamCant = 0
        .Quotas = 0
        .Inscribed = 0
        .DspInscription = 0
        .GldInscription = 0
        .LvlMax = 0
        .LvlMin = 0
        .TimeCancel = 0
        .NpcIndex = 0
        .TimeInit = 0
        .TimeCount = 0
        .CharBody = 0
        .CharHp = 0
        .Modality = 0
        
        For LoopC = LBound(.AllowedClasses()) To UBound(.AllowedClasses())
            .AllowedClasses(LoopC) = 0
        Next LoopC
        
    End With
End Sub

Private Function CheckUserEvent(ByVal Userindex As Integer, ByVal SlotEvent As Byte, ByRef ErrorMsg As String) As Boolean
    CheckUserEvent = False
        
    With UserList(Userindex)
        If .flags.Muerto Then
            ErrorMsg = "No puedes participar en eventos estando muerto."
            Exit Function
        End If
        If .flags.Mimetizado Then
            ErrorMsg = "No puedes entrar mimetizado."
            Exit Function
        End If
        
        If .flags.Montando Then
            ErrorMsg = "No puedes entrar montando."
            Exit Function
        End If
        
        If .flags.invisible Then
            ErrorMsg = "No puedes entrar invisible."
            Exit Function
        End If
        
        If .flags.SlotEvent > 0 Then
            ErrorMsg = "Ya te encuentras en un evento. Tipea /SALIREVENTO para salir del mismo."
            Exit Function
        End If
        
        If .Counters.Pena > 0 Then
            ErrorMsg = "No puedes participar de los eventos en la cárcel. Maldito prisionero!"
            Exit Function
        End If
        
        If MapInfo(.Pos.map).Pk Then
            ErrorMsg = "No puedes participar de los eventos estando en zona insegura. Vé a la ciudad mas cercana"
            Exit Function
        End If
        
        If .flags.Comerciando Then
            ErrorMsg = "No puedes participar de los eventos si estás comerciando."
            Exit Function
        End If
        
        If Not Events(SlotEvent).Enabled Or Events(SlotEvent).TimeInit > 0 Then
            ErrorMsg = "No hay ningun torneo disponible con ese nombre o bien las inscripciones no están disponibles aún."
            Exit Function
        End If
        
        If Events(SlotEvent).Run Then
            ErrorMsg = "El torneo ya ha comenzado. Mejor suerte para la próxima."
            Exit Function
        End If
        
        
        If Events(SlotEvent).LvlMin <> 0 Then
            If Events(SlotEvent).LvlMin > .Stats.ELV Then
                ErrorMsg = "Tu nivel no te permite ingresar a este evento."
                Exit Function
            End If
        End If
        
        If Events(SlotEvent).LvlMin <> 0 Then
            If Events(SlotEvent).LvlMax < .Stats.ELV Then
                ErrorMsg = "Tu nivel no te permite ingresar al evento."
                Exit Function
            End If
        End If
        
        If Events(SlotEvent).AllowedClasses(.clase) = 0 Then
            ErrorMsg = "Tu clase no está permitida en el evento."
            Exit Function
        End If
        
        
        If Events(SlotEvent).GldInscription > .Stats.Gld Then
            ErrorMsg = "No tienes suficiente oro para pagar el torneo. Pide prestado a un compañero."
            Exit Function
        End If
        
        If Events(SlotEvent).DspInscription > 0 Then
            If Not TieneObjetos(880, Events(SlotEvent).DspInscription, Userindex) Then
                ErrorMsg = "No tienes suficientes monedas DSP para participar del evento."
                Exit Function
            End If
        End If
        
        If Events(SlotEvent).Inscribed = Events(SlotEvent).Quotas Then
            ErrorMsg = "Los cupos del evento al que deseas participar ya fueron alcanzados."
            Exit Function
        End If
        
    
    End With
    CheckUserEvent = True
End Function

' EDICIÓN GENERAL
Private Function strModality(ByVal Modality As eModalityEvent) As String

    ' Modalidad de cada evento
    
    Select Case Modality
        Case eModalityEvent.CastleMode
            strModality = "CastleMode"
            
        Case eModalityEvent.DagaRusa
            strModality = "DagaRusa"
            
        Case eModalityEvent.DeathMatch
            strModality = "DeathMatch"
            
        Case eModalityEvent.Aracnus
            strModality = "Aracnus"
            
        Case eModalityEvent.HombreLobo
            strModality = "HombreLobo"
            
        Case eModalityEvent.Minotauro
            strModality = "Minotauro"
        
        Case eModalityEvent.Busqueda
            strModality = "Busqueda"
        
        Case eModalityEvent.Unstoppable
            strModality = "Unstoppable"
        
        Case eModalityEvent.Invasion
            strModality = "Invasion"
            
        Case eModalityEvent.Enfrentamientos
            strModality = "Enfrentamientos"
        
    End Select
End Function
Private Function strDescEvent(ByVal Modality As eModalityEvent) As String

    ' Descripción del evento en curso.
    Select Case Modality
        Case eModalityEvent.CastleMode
            strDescEvent = "» Los usuarios entrarán de forma aleatorea para formar dos equipos. Ambos equipos deberán defender a su rey y a su vez atacar al del equipo contrario."
        Case eModalityEvent.DagaRusa
            strDescEvent = "» Los usuarios se teletransportarán a una posición donde estará un asesino dispuesto a apuñalarlos y acabar con su vida. El último que quede en pie es el ganador del evento."
        Case eModalityEvent.DeathMatch
            strDescEvent = "» Los usuarios ingresan y luchan en una arena donde se toparan con todos los demás concursantes. El que logre quedar en pie, será el ganador."
        Case eModalityEvent.Aracnus
            strDescEvent = "» Un personaje es escogido al azar, para convertirse en una araña gigante la cual podrá envenenar a los demas concursantes acabando con su vida en el evento."
        Case eModalityEvent.Busqueda
            strDescEvent = "» Los personajes son teletransportados en un mapa donde su función principal será la recolección de objetos en el piso, para que así luego de tres minutos, el que recolecte más, ganará el evento."
        Case eModalityEvent.Unstoppable
            strDescEvent = "» Los personajes lucharan en un TODOS vs TODOS, donde los muertos no irán a su mapa de origen, si no que volverán a revivir para tener chances de ganar el evento. El que logre matar más personajes, ganará el evento."
        Case eModalityEvent.Invasion
            strDescEvent = "» Los personajes son llevados a un mapa donde aparecerán criaturas únicas de DesteriumAO, cada criatura dará una recompensa única y los usuarios tendrán chances de entrenar sus personajes."
        Case eModalityEvent.Enfrentamientos
            strDescEvent = "» Esta modalidad es muy especial. Los usuarios combatirán en enfrentamientos 1vs1, 2vs2 , 3vs3 , entre otros. [A elección del GM y parejas AL AZAR]"
    End Select
End Function
Private Sub InitEvent(ByVal SlotEvent As Byte)
    
    Select Case Events(SlotEvent).Modality
        Case eModalityEvent.CastleMode
            Call InitCastleMode(SlotEvent)
            
        Case eModalityEvent.DagaRusa
            Call InitDagaRusa(SlotEvent)
            
        Case eModalityEvent.DeathMatch
            Call InitDeathMatch(SlotEvent)
            
        Case eModalityEvent.Aracnus
            Call InitEventTransformation(SlotEvent, 254, 6500, 211, 70, 36)
            
        Case eModalityEvent.HombreLobo
            Call InitEventTransformation(SlotEvent, 255, 3500, 211, 70, 36)
            
        Case eModalityEvent.Minotauro
            Call InitEventTransformation(SlotEvent, 253, 2500, 211, 70, 36)
        
        Case eModalityEvent.Busqueda
            Call InitBusqueda(SlotEvent)
            
        Case eModalityEvent.Unstoppable
            InitUnstoppable SlotEvent
            
        Case eModalityEvent.Invasion
        
        Case eModalityEvent.Enfrentamientos
            Call InitFights(SlotEvent)
        
        Case Else
            Exit Sub
        
    End Select
End Sub
Public Function CanAttackUserEvent(ByVal Userindex As Integer, ByVal Victima As Integer) As Boolean
    
    ' Si el personaje es del mismo team, no se puede atacar al usuario.
    Dim VictimaSlotUserEvent As Byte
    
    VictimaSlotUserEvent = UserList(Victima).flags.SlotUserEvent
    
    With UserList(Userindex)
        If .flags.SlotEvent > 0 Then
            If Events(.flags.SlotEvent).Users(VictimaSlotUserEvent).Team = Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).Team Then
                CanAttackUserEvent = False
                Exit Function
            End If
        End If
        CanAttackUserEvent = True
    End With
End Function

Private Function ChangeBodyEvent(ByVal SlotEvent As Byte, ByVal Userindex As Integer, ByVal ChangeHead As Boolean)
    
    ' En caso de que el evento cambie el body, de lo cambiamos.
    With UserList(Userindex)
        .CharMimetizado.body = .Char.body
        .CharMimetizado.Head = .Char.Head
        .CharMimetizado.CascoAnim = .Char.CascoAnim
        .CharMimetizado.ShieldAnim = .Char.ShieldAnim
        .CharMimetizado.WeaponAnim = .Char.WeaponAnim

        .Char.body = Events(SlotEvent).CharBody
        .Char.Head = IIf(ChangeHead = False, .Char.Head, 0)
        .Char.CascoAnim = 0
        .Char.ShieldAnim = 0
        .Char.WeaponAnim = 0
                
        ChangeUserChar Userindex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, True
        RefreshCharStatus Userindex
    
    End With
End Function

Private Function ResetBodyEvent(ByVal SlotEvent As Byte, ByVal Userindex As Integer)

    ' En caso de que el evento cambie el body del personaje, se lo restauramos.
    
    With UserList(Userindex)
        If .flags.Muerto Then Exit Function
        'If Events(SlotEvent).Users(.flags.SlotUserEvent).Selected = 0 Then Exit Function
        
        If .CharMimetizado.body > 0 Then
            .Char.body = .CharMimetizado.body
            .Char.Head = .CharMimetizado.Head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            
            
            .CharMimetizado.body = 0
            .CharMimetizado.Head = 0
            .CharMimetizado.CascoAnim = 0
            .CharMimetizado.ShieldAnim = 0
            .CharMimetizado.WeaponAnim = 0
            
            .showName = True
            
            ChangeUserChar Userindex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, True
            RefreshCharStatus Userindex
        End If
    
    End With
End Function

Private Sub ChangeHpEvent(ByVal Userindex As Integer)

    ' En caso de que el evento edite la vida del personaje, se la editamos.
    
    Dim SlotEvent As Byte
    
    With UserList(Userindex)
        SlotEvent = .flags.SlotEvent
        
        .Stats.OldHp = .Stats.MaxHp
        
        .Stats.MaxHp = Events(SlotEvent).CharHp
        .Stats.MinHp = .Stats.MaxHp
        
        WriteUpdateUserStats Userindex
    
    End With
End Sub

Private Sub ResetHpEvent(ByVal Userindex As Integer)

    ' En caso de que el evento haya editado la vida de un personaje, se la volvemos a restaurar.
    
    With UserList(Userindex)
        If .Stats.OldHp = 0 Then Exit Sub
        .Stats.MaxHp = .Stats.OldHp
        '.Stats.MinHp = .Stats.MaxHp
        .Stats.OldHp = 0
        WriteUpdateHP Userindex
        
    End With
End Sub




'// Fin Funciones generales '//

Public Sub NewEvent(ByVal Userindex As Integer, _
                    ByVal Modality As eModalityEvent, _
                    ByVal Quotas As Byte, _
                    ByVal LvlMin As Byte, _
                    ByVal LvlMax As Byte, _
                    ByVal GldInscription As Long, _
                    ByVal DspInscription As Long, _
                    ByVal TimeInit As Long, _
                    ByVal TimeCancel As Long, _
                    ByVal TeamCant As Byte, _
                    ByRef AllowedClasses() As Byte)
                    
    Dim Slot As Integer
    Dim strTemp As String

    Slot = FreeSlotEvent()
    
    If Slot = 0 Then
        WriteConsoleMsg Userindex, "No hay más lugar disponible para crear un evento simultaneo. Espera a que termine alguno o bien cancela alguno.", FontTypeNames.FONTTYPE_INFO
        Exit Sub
    Else
        With Events(Slot)
            .Enabled = True
            .Modality = Modality
            .TeamCant = TeamCant
            .Quotas = Quotas
            .LvlMin = LvlMin
            .LvlMax = LvlMax
            .GldInscription = GldInscription
            .DspInscription = DspInscription
            .AllowedClasses = AllowedClasses
            .TimeInit = TimeInit
            .TimeCancel = TimeCancel
        
            ReDim .Users(1 To .Quotas) As tUserEvent
            
            ' strModality devuelve: "Evento '1vs1' : Descripción"
            strTemp = strModality(.Modality) & strDescEvent(.Modality) & vbCrLf
            strTemp = strTemp & "Cupos máximos: " & .Quotas & vbCrLf

            strTemp = strTemp & IIf((.LvlMin > 0), "Nivel mínimo: " & .LvlMin & vbCrLf, vbNullString)
            strTemp = strTemp & IIf((.LvlMax > 0), "Nivel máximo: " & .LvlMax & vbCrLf, vbNullString)
            
            If .GldInscription > 0 And .DspInscription > 0 Then
                strTemp = strTemp & "Inscripción requerida: " & .GldInscription & " monedas de oro y " & .DspInscription & " monedas DSP."
            ElseIf .GldInscription > 0 Then
                strTemp = strTemp & "Inscripción requerida: " & .GldInscription & " monedas de oro."
            ElseIf .DspInscription > 0 Then
                strTemp = strTemp & "Inscripción requerida: " & .DspInscription & " monedas DSP."
            Else
                strTemp = strTemp & "Inscripción GRATIS"
            End If
            
            strTemp = strTemp & vbCrLf
            
            strTemp = strTemp & "Clases permitidas: " & CheckAllowedClasses(AllowedClasses) & ". Comando para ingresar /INGRESAR " & strModality(.Modality) & vbCrLf
            strTemp = strTemp & "Las inscripciones abren en " & Int(.TimeInit / 60) & " minutos."
            
            LoadMapEvent
        End With
        
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strTemp, FontTypeNames.FONTTYPE_INFOBOLD)
    End If
    
End Sub

Public Sub CloseEvent(ByVal Slot As Byte, Optional ByVal MsgConsole As String = vbNullString)
    With Events(Slot)
        If MsgConsole <> vbNullString Then SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(MsgConsole, FontTypeNames.FONTTYPE_ORO)
        

        
        Call ResetEvent(Slot)
    End With
End Sub




Public Sub ParticipeEvent(ByVal Userindex As Integer, ByVal Modality As eModalityEvent)
    
    Dim ErrorMsg As String
    Dim SlotUser As Byte
    Dim Pos As WorldPos
    Dim SlotEvent As Integer
    
    SlotEvent = SearchSlotEvent(Modality)
    
    If SlotEvent = 0 Then
        'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Error Fatal TESTEO", FontTypeNames.FONTTYPE_ADMIN)
        Exit Sub
    End If
    
    With UserList(Userindex)
        If CheckUserEvent(Userindex, SlotEvent, ErrorMsg) Then
            SlotUser = FreeSlotUser(SlotEvent)
            
            .flags.SlotEvent = SlotEvent
            .flags.SlotUserEvent = SlotUser
            
            .PosAnt.map = .Pos.map
            .PosAnt.X = .Pos.X
            .PosAnt.Y = .Pos.Y
            
            .Stats.Gld = .Stats.Gld - Events(SlotEvent).GldInscription
            Call WriteUpdateGold(Userindex)
            
            Call QuitarObjetos(880, Events(SlotEvent).DspInscription, Userindex)
            
            With Events(SlotEvent)
                Pos.map = 211
                Pos.X = 30
                Pos.Y = 21
                
                Call FindLegalPos(Userindex, Pos.map, Pos.X, Pos.Y)
                Call WarpUserChar(Userindex, Pos.map, Pos.X, Pos.Y, False)
            
                .Users(SlotUser).Id = Userindex
                .Inscribed = .Inscribed + 1
                
                
                WriteConsoleMsg Userindex, "Has ingresado al evento " & strModality(.Modality) & ". Espera a que se completen los cupos para que comience.", FontTypeNames.FONTTYPE_INFO
                
                If .Inscribed = .Quotas Then
                    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(.Modality) & "» Los cupos han sido alcanzados. Les deseamos mucha suerte a cada uno de los participantes y que gane el mejor!", FontTypeNames.FONTTYPE_GUILD)
                    .Run = True
                    InitEvent SlotEvent
                    Exit Sub
                End If
            End With
        
        Else
            WriteConsoleMsg Userindex, ErrorMsg, FontTypeNames.FONTTYPE_WARNING
        
        End If
    End With
End Sub



Public Sub AbandonateEvent(ByVal Userindex As Integer, _
                            Optional ByVal MsgAbandonate As Boolean = False, _
                            Optional ByVal Forzado As Boolean = False)
    
    Dim Pos As WorldPos
    Dim SlotEvent As Byte
    
    With UserList(Userindex)
        SlotEvent = .flags.SlotEvent
        
        If SlotEvent > 0 And .flags.SlotUserEvent > 0 Then
            With Events(SlotEvent)
                If .Inscribed > 0 Then .Inscribed = .Inscribed - 1
                    
                    If Forzado And .Inscribed > 1 Then
                        If .Users(UserList(Userindex).flags.SlotUserEvent).Selected = 1 Then
                             Transformation_SelectionUser SlotEvent
                        End If
                    End If
                    
                    
                    If .Run Then
                        If .Modality = DagaRusa Then
                            Call WriteUserInEvent(Userindex)
                            
                            If Forzado Then
                                ' Si estaba en DagaRusa y no pasó todavía, significa que es forzado. Por lo cual tenemos que descontarle al NPC.
                                If .Users(UserList(Userindex).flags.SlotUserEvent).value = 0 Then
                                    Npclist(.NpcIndex).flags.InscribedPrevio = Npclist(.NpcIndex).flags.InscribedPrevio - 1
                                End If
                            End If
                        End If
                    End If
                    
                    .Users(UserList(Userindex).flags.SlotUserEvent).Id = 0
                    .Users(UserList(Userindex).flags.SlotUserEvent).Team = 0
                    .Users(UserList(Userindex).flags.SlotUserEvent).value = 0
                    .Users(UserList(Userindex).flags.SlotUserEvent).Selected = 0
                    .Users(UserList(Userindex).flags.SlotUserEvent).MapFight = 0
                    
                    UserList(Userindex).flags.SlotEvent = 0
                    UserList(Userindex).flags.SlotUserEvent = 0
                
                            
                    Pos.map = UserList(Userindex).PosAnt.map
                    Pos.X = UserList(Userindex).PosAnt.X
                    Pos.Y = UserList(Userindex).PosAnt.Y
                    
                    Call FindLegalPos(Userindex, Pos.map, Pos.X, Pos.Y)
                    Call WarpUserChar(Userindex, Pos.map, Pos.X, Pos.Y, False)
                    
                    If Events(SlotEvent).CharBody <> 0 Then
                        Call ResetBodyEvent(SlotEvent, Userindex)
                
                    End If
            
                    If UserList(Userindex).Stats.OldHp <> 0 Then
                        ResetHpEvent Userindex
                    End If
            
                    UserList(Userindex).showName = True
                    RefreshCharStatus Userindex
                    
                    
                    If MsgAbandonate Then WriteConsoleMsg Userindex, "Has abandonado el evento. Podrás recibir una pena por hacer esto.", FontTypeNames.FONTTYPE_WARNING
                    
        
                    ' Abandono general del evento
                    If .Inscribed = 1 And Forzado Then
                        Call FinishEvent(SlotEvent)
                    
                        CloseEvent SlotEvent
                        Exit Sub
                    End If
                    
                    
            End With
        End If
        
        
    End With
End Sub

Private Sub FinishEvent(ByVal SlotEvent As Byte)
    
    Dim Userindex As Integer
    Dim IsSelected As Boolean
    
    With Events(SlotEvent)
        Select Case .Modality
            Case eModalityEvent.CastleMode
                Userindex = SearchLastUserEvent(SlotEvent)
                CastleMode_Premio Userindex, False
                
            Case eModalityEvent.DagaRusa
                DagaRusa_CheckWin SlotEvent
                
            Case eModalityEvent.DeathMatch
                Userindex = SearchLastUserEvent(SlotEvent)
                DeathMatch_Premio Userindex
                
            Case eModalityEvent.Aracnus, eModalityEvent.HombreLobo, eModalityEvent.Minotauro
                Userindex = SearchLastUserEvent(SlotEvent)
                
                If .Users(UserList(Userindex).flags.SlotUserEvent).Selected = 1 Then IsSelected = True
                
                Transformation_Premio Userindex, IsSelected, 250000
                
            Case eModalityEvent.Busqueda
                Busqueda_SearchWin SlotEvent
                
            Case eModalityEvent.Unstoppable
                Unstoppable_UserWin SlotEvent
                
        End Select
    End With
    
    
    
End Sub


'#################EVENTO CASTLE MODE##########################
Public Function CanAttackReyCastle(ByVal Userindex As Integer, ByVal NpcIndex As Integer) As Boolean
    With UserList(Userindex)
        If .flags.SlotEvent > 0 Then
            If Npclist(NpcIndex).flags.TeamEvent = Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).Team Then
                CanAttackReyCastle = False
                Exit Function
            End If
        End If
    
    
        CanAttackReyCastle = True
    End With
End Function
Private Sub CastleMode_InitRey()
    Dim NpcIndex As Integer
    Const NumRey As Integer = 697
    Dim Pos As WorldPos
    Dim LoopX As Integer, LoopY As Integer
    Const Rango As Byte = 5
    
    For LoopX = 74 - Rango To 74 + Rango
        For LoopY = 24 - Rango To 24 + Rango
            If MapData(212, LoopX, LoopY).NpcIndex > 0 Then
                Call QuitarNPC(MapData(212, LoopX, LoopY).NpcIndex)
            End If
        Next LoopY
    Next LoopX
    
    Pos.map = 212
        
    Pos.X = 74
    Pos.Y = 24
    NpcIndex = SpawnNpc(NumRey, Pos, False, False)
    Npclist(NpcIndex).flags.TeamEvent = 1
        
        
    For LoopX = 19 - Rango To 19 + Rango
        For LoopY = 34 - Rango To 34 + Rango
            If MapData(212, LoopX, LoopY).NpcIndex > 0 Then
                Call QuitarNPC(MapData(212, LoopX, LoopY).NpcIndex)
            End If
        Next LoopY
    Next LoopX
    
    Pos.X = 19
    Pos.Y = 34
    NpcIndex = SpawnNpc(NumRey, Pos, False, False)
    Npclist(NpcIndex).flags.TeamEvent = 2
    
End Sub

Public Sub InitCastleMode(ByVal SlotEvent As Byte)
    Dim LoopC As Integer
    
    Const NumRey As Integer = 697
    Dim NpcIndex As Integer
    Dim Pos As WorldPos
    
    ' Spawn the npc castle mode
    CastleMode_InitRey
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                If LoopC > (UBound(.Users()) / 2) Then
                    .Users(LoopC).Team = 2
                    Pos.map = 212
                    Pos.X = 19
                    Pos.Y = 34
                    
                    Call FindLegalPos(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y)
                    Call WarpUserChar(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y, False)
                Else
                    .Users(LoopC).Team = 1
                    Pos.map = 212
                    Pos.X = 74
                    Pos.Y = 24
                    
                    Call FindLegalPos(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y)
                    Call WarpUserChar(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y, False)
                    
                End If
            End If
        Next LoopC
    End With
    
End Sub
Public Sub CastleMode_UserRevive(ByVal Userindex As Integer)

    Dim LoopC As Integer
    Dim Pos As WorldPos
    
    With UserList(Userindex)
        If .flags.SlotEvent > 0 Then
            Call RevivirUsuario(Userindex)
            
            
            Pos.map = 212
            Pos.X = RandomNumber(20, 80)
            Pos.Y = RandomNumber(20, 80)
            
            Call ClosestLegalPos(Pos, Pos)
            'Call FindLegalPos(Userindex, Pos.Map, Pos.X, Pos.Y)
            Call WarpUserChar(Userindex, Pos.map, Pos.X, Pos.Y, True)
        
        End If
    End With
End Sub

Public Sub FinishCastleMode(ByVal SlotEvent As Byte, ByVal UserEventSlot As Integer)
    Dim LoopC As Integer
    Dim strTemp As String
    Dim NpcIndex As Integer
    Dim MiObj As Obj
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                If .Users(LoopC).Team = .Users(UserEventSlot).Team Then
                    If LoopC = UserEventSlot Then
                        CastleMode_Premio .Users(LoopC).Id, True
                    Else
                        CastleMode_Premio .Users(LoopC).Id, False
                    End If
                    
                    If strTemp = vbNullString Then
                        strTemp = UserList(.Users(LoopC).Id).Name
                    Else
                        strTemp = strTemp & ", " & UserList(.Users(LoopC).Id).Name
                    End If
                End If
            End If
        Next LoopC
        
        
        CloseEvent SlotEvent, "CastleMode» Ha finalizado. Ha ganado el equipo de " & UCase$(strTemp)
    End With
    
End Sub

Private Sub CastleMode_Premio(ByVal Userindex As Integer, ByVal KillRey As Boolean)

    ' Entregamos el premio del CastleMode
    Dim MiObj As Obj
    
    With UserList(Userindex)
        .Stats.Gld = .Stats.Gld + 250000
        WriteConsoleMsg Userindex, "Felicitaciones, has recibido 250.000 monedas de oro por haber ganado el evento!", FontTypeNames.FONTTYPE_INFO
        
        If KillRey Then
            WriteConsoleMsg Userindex, "Hemos notado que has aniquilado con la vida del rey oponente. ¡FELICITACIONES! Aquí tienes tu recompensa! 250.000 monedas de oro extra y su equipamiento", FontTypeNames.FONTTYPE_INFO
            .Stats.Gld = .Stats.Gld + 250000
            
        End If
        
        MiObj.objindex = 899
        MiObj.Amount = 1
                        
        If Not MeterItemEnInventario(Userindex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
                        
        MiObj.objindex = 900
        MiObj.Amount = 1
                        
        If Not MeterItemEnInventario(Userindex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
        
        WriteUpdateGold Userindex
        
        .Stats.TorneosGanados = .Stats.TorneosGanados + 1
    End With
End Sub

' FIN EVENTO CASTLE MODE #####################################

' ###################### EVENTO DAGA RUSA ###########################
Public Sub InitDagaRusa(ByVal SlotEvent As Byte)

    Dim LoopC As Integer
    Dim NpcIndex As Integer
    Dim Pos As WorldPos
    
    Dim Num As Integer
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                Call WarpUserChar(.Users(LoopC).Id, 211, 21 + Num, 60, False)
                Num = Num + 1
                Call WriteUserInEvent(.Users(LoopC).Id)
            End If
        Next LoopC
        
        Pos.map = 211
        Pos.X = 21
        Pos.Y = 59
        NpcIndex = SpawnNpc(704, Pos, False, False)
    
        If NpcIndex <> 0 Then
            Npclist(NpcIndex).Movement = NpcDagaRusa
            Npclist(NpcIndex).flags.SlotEvent = SlotEvent
            Npclist(NpcIndex).flags.InscribedPrevio = .Inscribed
            .NpcIndex = NpcIndex
            
            DagaRusa_MoveNpc NpcIndex, True
        End If
        
        
        .TimeCount = 10
    End With


End Sub
Public Function DagaRusa_NextUser(ByVal SlotEvent As Byte) As Byte
    Dim LoopC As Integer
    
    DagaRusa_NextUser = 0
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If (.Users(LoopC).Id > 0) And (.Users(LoopC).value = 0) Then
                DagaRusa_NextUser = .Users(LoopC).Id
                '.Users(LoopC).Value = 1
                Exit For
            End If
        Next LoopC
    End With
        
End Function
Public Sub DagaRusa_ResetRonda(ByVal SlotEvent As Byte)

    Dim LoopC As Integer
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            
            .Users(LoopC).value = 0
        Next LoopC
    
    End With
End Sub
Private Sub DagaRusa_CheckWin(ByVal SlotEvent As Byte)

    Dim Userindex As Integer
    Dim MiObj As Obj
    
    With Events(SlotEvent)
        If .Inscribed = 1 Then
            Userindex = SearchLastUserEvent(SlotEvent)
            DagaRusa_Premio Userindex
            

            Call QuitarNPC(.NpcIndex)
            CloseEvent SlotEvent
            
        End If
    End With
End Sub

Private Sub DagaRusa_Premio(ByVal Userindex As Integer)

    Dim MiObj As Obj
    
    With UserList(Userindex)
         MiObj.Amount = 1
         MiObj.objindex = 1037
        
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("DagaRusa» El ganador es " & UserList(Userindex).Name & ". Felicitaciones para el personaje, quien se ha ganado una MD! (Espada mata dragones)", FontTypeNames.FONTTYPE_GUILD)
        
        If Not MeterItemEnInventario(Userindex, MiObj) Then
            Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
        End If
        
        .Stats.TorneosGanados = .Stats.TorneosGanados + 1
        
    End With
End Sub
Public Sub DagaRusa_AttackUser(ByVal Userindex As Integer, ByVal NpcIndex As Integer)
    
    Dim N As Integer
    Dim Slot As Byte
    
    With UserList(Userindex)
        
        N = 10
        
        If RandomNumber(1, 100) <= N Then
        
            ' Sound
            SendData SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y)
            ' Fx
            SendData SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0)
            ' Cambio de Heading
            ChangeNPCChar NpcIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.Head, SOUTH
            'Apuñalada en el piso
            SendData SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, 1000, DAMAGE_PUÑAL)
            
            WriteConsoleMsg Userindex, "¡Has sido apuñalado por 1.000!", FontTypeNames.FONTTYPE_FIGHT
            
            Slot = .flags.SlotEvent
            
            
            Call UserDie(Userindex)
            EventosDS.AbandonateEvent (Userindex)
            Call DagaRusa_CheckWin(Slot)
           
            
        Else
            ' Sound
            SendData SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y)
            ' Fx
            SendData SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0)
            ' Cambio de Heading
            ChangeNPCChar NpcIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.Head, SOUTH

            WriteConsoleMsg Userindex, "¡Parece que no te he apuñalado, ya verás!", FontTypeNames.FONTTYPE_FIGHT
           ' SendData SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, 1000, DAMAGE_PUÑAL)
        End If
        
        
        
    End With
End Sub

' FIN EVENTO DAGA RUSA ###########################################
Private Function SelectModalityDeathMatch(ByVal SlotEvent As Byte) As Integer
    Dim Random As Integer
    
    Randomize
    Random = RandomNumber(1, 8)
    
    With Events(SlotEvent)
        Select Case Random
            Case 1 ' Zombie
                .CharBody = 11
            Case 2 ' Golem
                .CharBody = 11
            Case 3 ' Araña
                .CharBody = 42
            Case 4 ' Asesino
                .CharBody = 11 '48
            Case 5 'Medusa suprema
                .CharBody = 151
            Case 6 'Dragón azul
                .CharBody = 42 '247
            Case 7 'Viuda negra 185
                .CharBody = 185
            Case 8 'Tigre salvaje
                .CharBody = 147
        End Select
    End With
End Function

' DEATHMATCH ####################################################
Private Sub InitDeathMatch(ByVal SlotEvent As Byte)

    Dim LoopC As Integer
    Dim Pos As WorldPos
    
    Call SelectModalityDeathMatch(SlotEvent)
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                .Users(LoopC).Team = LoopC
                .Users(LoopC).Selected = 1
                
                ChangeBodyEvent SlotEvent, .Users(LoopC).Id, True
                UserList(.Users(LoopC).Id).showName = False
                RefreshCharStatus .Users(LoopC).Id
                
                
                Pos.map = 211
                Pos.X = RandomNumber(58, 84)
                Pos.Y = RandomNumber(28, 44)
            
                Call ClosestLegalPos(Pos, Pos)
                Call WarpUserChar(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y, True)
            End If
        
        Next LoopC
    
        .TimeCount = 20
    End With
    
End Sub

Public Sub DeathMatch_UserDie(ByVal SlotEvent As Byte, ByVal Userindex As Integer)
    AbandonateEvent (Userindex)
        
    If Events(SlotEvent).Inscribed = 1 Then
        Userindex = SearchLastUserEvent(SlotEvent)
        DeathMatch_Premio Userindex
        CloseEvent SlotEvent
    End If
End Sub
Private Sub DeathMatch_Premio(ByVal Userindex As Integer)
    With UserList(Userindex)
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("DeathMatch» El ganador es " & .Name & " quien se lleva 1 punto de torneo y 450.000 monedas de oro.", FontTypeNames.FONTTYPE_GUILD)
            
        .Stats.Gld = .Stats.Gld + 450000
        WriteUpdateGold Userindex
        
        .Stats.TorneosGanados = .Stats.TorneosGanados + 1
    End With
End Sub

' FIN DEATHMATCH ################################################
' EVENTOS DONDE LOS USUARIOS SE TRANSFORMAN EN CRIATURAS
Private Sub InitEventTransformation(ByVal SlotEvent As Byte, _
                                    ByVal NewBody As Integer, _
                                    ByVal NewHp As Integer, _
                                    ByVal map As Integer, _
                                    ByVal X As Byte, _
                                    ByVal Y As Byte)
    
    Dim LoopC As Integer
    Dim UserSelected As Integer
    Dim Pos As WorldPos
    
    Const Rango As Byte = 4
    
    With Events(SlotEvent)
        .CharBody = NewBody
        .CharHp = NewHp
        
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                .Users(LoopC).Team = 2
                
                
                Pos.map = map
                Pos.X = RandomNumber(X - Rango, X + Rango)
                Pos.Y = RandomNumber(Y - Rango, Y + Rango)
            
                Call ClosestLegalPos(Pos, Pos)
                Call WarpUserChar(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y, True)
                
            End If
        Next LoopC
        
        Transformation_SelectionUser SlotEvent
    End With
End Sub

Private Function Transformation_SelectionUser(ByVal SlotEvent As Byte)
    Dim LoopC As Integer
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            Transformation_SelectionUser = RandomNumber(LBound(.Users()), UBound(.Users()))
            
            If .Users(Transformation_SelectionUser).Id > 0 And .Users(Transformation_SelectionUser).Selected = 0 Then
                Exit For
            End If
        Next LoopC
        
        .Users(Transformation_SelectionUser).Selected = 1
        .Users(Transformation_SelectionUser).Team = 1
                    
        Call ChangeHpEvent(.Users(Transformation_SelectionUser).Id)
        Call ChangeBodyEvent(SlotEvent, .Users(Transformation_SelectionUser).Id, IIf(.Modality = Minotauro, False, True))
    End With
End Function

Public Sub Transformation_UserDie(ByVal Userindex As Integer, ByVal AttackerIndex As Integer)
    Dim SlotEvent As Byte
    Dim Exituser As Boolean
    
    With UserList(Userindex)
        SlotEvent = .flags.SlotEvent
        AbandonateEvent Userindex
        
        Transformation_CheckWin Userindex, SlotEvent, AttackerIndex
    End With
End Sub
Private Function Transformation_SearchUserSelected(ByVal SlotEvent As Byte) As Integer
    Dim LoopC As Integer
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Selected = 1 Then
                Transformation_SearchUserSelected = LoopC
            End If
        Next LoopC
    End With
End Function
Public Sub Transformation_CheckWin(ByVal Userindex As Integer, ByVal SlotEvent As Byte, Optional ByVal AttackerIndex As Integer = 0)
    Dim IsSelected As Boolean
    
    With Events(SlotEvent)
        If .Inscribed = 1 Then
            Userindex = SearchLastUserEvent(SlotEvent)
            If .Users(UserList(Userindex).flags.SlotUserEvent).Selected = 1 Then IsSelected = True
            
            Transformation_Premio Userindex, IsSelected, 250000
            
            CloseEvent SlotEvent
            Exit Sub
        End If
        
        
        'Significa que hay más de un usuario. Por lo tanto podría haber muerto el bicho transformado
        
        If UserList(Userindex).flags.SlotUserEvent = Transformation_SearchUserSelected(SlotEvent) Then
            'Userindex = SearchLastUserEvent(SlotEvent)
            Transformation_Premio AttackerIndex, False, 250000
            
            CloseEvent SlotEvent
        End If
    End With
End Sub

Private Sub Transformation_Premio(ByVal Userindex As Integer, _
                                    ByVal IsSelected As Boolean, _
                                    ByVal Gld As Long)
                                    
    Dim UserWin As Integer
    
    With UserList(Userindex)
        Dim SlotEvent As Byte
        SlotEvent = .flags.SlotEvent
        
        If IsSelected Then
            .Stats.Gld = .Stats.Gld + (Gld * 2)
            WriteConsoleMsg Userindex, "Has recibido " & (Gld * 2) & " por haber aniquilado a todos los usuarios.", FontTypeNames.FONTTYPE_INFO
            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(Events(SlotEvent).Modality) & "» Ha logrado derrotar a todos los participantes. Felicitaciones para " & .Name & " quien fue escogido como " & strModality(Events(SlotEvent).Modality), FontTypeNames.FONTTYPE_GUILD)

        Else
            .Stats.Gld = .Stats.Gld + Gld
            WriteConsoleMsg Userindex, "Has recibido " & Gld & " por haber aniquilado a " & strModality(Events(SlotEvent).Modality), FontTypeNames.FONTTYPE_INFO
            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(Events(SlotEvent).Modality) & "» Felicitaciones para " & .Name & " quien derrotó a " & strModality(Events(SlotEvent).Modality), FontTypeNames.FONTTYPE_GUILD)

        End If
        
        WriteUpdateGold Userindex
        
        .Stats.TorneosGanados = .Stats.TorneosGanados + 1
    
    End With
End Sub


' FIN EVENTOS DONDE LOS USUARIOS SE TRANSFORMAN EN CRIATURAS

' ARACNUS #######################################################

Public Sub Aracnus_Veneno(ByVal AttackerIndex As Integer, ByVal Victimindex As Integer)

    ' El personaje transformado en Aracnus, tiene 10% de probabilidad de envenenar a la víctima y dejarla fuera del torneo.

    Const N As Byte = 10
    
    With UserList(AttackerIndex)
        If RandomNumber(1, 100) <= 10 Then
            WriteConsoleMsg Victimindex, "Has sido envenenado por Aracnus, has muerto de inmediato por su veneno letal.", FontTypeNames.FONTTYPE_FIGHT
            Call UserDie(Victimindex)
            
            Transformation_CheckWin Victimindex, .flags.SlotEvent, AttackerIndex
        End If
    
    End With
End Sub

Public Sub Minotauro_Veneno(ByVal AttackerIndex As Integer, ByVal Victimindex As Integer)
    
    ' El personaje transformado en Minotauro, tiene 10% de posibilidad de dar un golpe mortal
    Const N As Byte = 10
    
    With UserList(AttackerIndex)
        If RandomNumber(1, 100) <= 10 Then
            WriteConsoleMsg Victimindex, "¡El minotauro ha logrado paralizar tu cuerpo con su dosis de veneno. Has quedado afuera del evento.", FontTypeNames.FONTTYPE_FIGHT
            Call UserDie(Victimindex)
            
            Transformation_CheckWin Victimindex, .flags.SlotEvent, AttackerIndex
        
        End If
    
    End With
End Sub

' FIN ARACNUS ###################################################

' EVENTO BUSQUEDA '
Private Sub InitBusqueda(ByVal SlotEvent As Byte)

    
    Dim LoopC As Integer
    Dim Pos As WorldPos
    
    With Events(SlotEvent)
        For LoopC = 1 To 20
            Busqueda_CreateObj 216, RandomNumber(20, 80), RandomNumber(20, 80)
        Next LoopC
        
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                Pos.map = 216
                Pos.X = RandomNumber(50, 60)
                Pos.Y = RandomNumber(50, 60)
                
                Call ClosestLegalPos(Pos, Pos)
                Call WarpUserChar(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y, True)
            End If
        Next LoopC
        
        .TimeFinish = 60
    
    End With
End Sub

Private Sub Busqueda_CreateObj(ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)

    ' Creamos un objeto en el mapa de búsqueda.
    
    Dim Pos As WorldPos
    Dim Obj As Obj
    
    Pos.map = map
    Pos.X = X
    Pos.Y = Y
    ClosestStablePos Pos, Pos
    
    Obj.objindex = 1037
    Obj.Amount = 1
    Call MakeObj(Obj, Pos.map, Pos.X, Pos.Y)
    MapData(Pos.map, Pos.X, Pos.Y).ObjEvent = 1
    
End Sub
Private Sub Busqueda_SearchWin(ByVal SlotEvent As Byte)
    With Events(SlotEvent)
        Dim Userindex As Integer
        Userindex = Busqueda_UserRecolectedObj(SlotEvent)
        
        ' vercoso este userindex 0
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Busqueda de objetos» El ganador de la búsqueda de objetos es " & UserList(Userindex).Name & ". Felicitaciones! Se lleva como premio 350.000 monedas de oro", FontTypeNames.FONTTYPE_GUILD)
        
        CloseEvent SlotEvent
        
    End With
End Sub
Private Function Busqueda_UserRecolectedObj(ByVal SlotEvent As Byte) As Integer
    Dim LoopC As Integer
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            
            If .Users(LoopC).Id > 0 Then
                If Busqueda_UserRecolectedObj = 0 Then Busqueda_UserRecolectedObj = LoopC
                If .Users(LoopC).value > .Users(Busqueda_UserRecolectedObj).value Then
                    Busqueda_UserRecolectedObj = LoopC
                End If
            End If
                
        Next LoopC
        
        Busqueda_UserRecolectedObj = .Users(Busqueda_UserRecolectedObj).Id
    End With
End Function

Public Sub Busqueda_GetObj(ByVal SlotEvent As Byte, ByVal SlotUserEvent As Byte)
    With Events(SlotEvent)
        .Users(SlotUserEvent).value = .Users(SlotUserEvent).value + 1
        
        WriteConsoleMsg .Users(SlotUserEvent).Id, "Has recolectado un objeto del piso. En total llevas " & .Users(SlotUserEvent).value & " objetos recolectados. Sigue así!", FontTypeNames.FONTTYPE_INFO
        
        Busqueda_CreateObj 216, RandomNumber(30, 80), RandomNumber(30, 80)
    End With
End Sub

' ENFRENTAMIENTOS ###############################################

Private Sub InitFights(ByVal SlotEvent As Byte)
    
    Fight_SelectedTeam SlotEvent
    Fight_Combate SlotEvent
    
End Sub
Private Sub Fight_SelectedTeam(ByVal SlotEvent As Byte)
    
    ' En los enfrentamientos utilizamos este procedimiento para seleccionar los grupos o bien el usuario queda solo por 1vs1.
    Dim LoopX As Integer
    Dim LoopY As Integer
    Dim Team As Byte
    
    Team = 1
    
    With Events(SlotEvent)
        For LoopX = LBound(.Users()) To UBound(.Users()) Step .TeamCant
            For LoopY = 0 To (.TeamCant - 1)
                .Users(LoopX + LoopY).Team = Team
            Next LoopY
            
            Team = Team + 1
        Next LoopX
    
    End With
End Sub

Private Sub Fight_WarpTeam(ByVal SlotEvent As Byte, _
                                        ByVal ArenaSlot As Byte, _
                                        ByVal TeamEvent As Byte, _
                                        ByVal IsContrincante As Boolean, _
                                        ByRef strTeam As String)

    Dim LoopC As Integer
    Dim strTemp As String, strTemp1 As String, strTemp2 As String
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Team = TeamEvent Then
                If IsContrincante Then
                    Call EventWarpUser(.Users(LoopC).Id, MapEvent.Fight(ArenaSlot).map, MapEvent.Fight(ArenaSlot).X + MAP_TILE_VS, MapEvent.Fight(ArenaSlot).Y + MAP_TILE_VS)
                Else
                    Call EventWarpUser(.Users(LoopC).Id, MapEvent.Fight(ArenaSlot).map, MapEvent.Fight(ArenaSlot).X, MapEvent.Fight(ArenaSlot).Y)
                End If
                
                If strTeam = vbNullString Then
                    strTeam = UserList(.Users(LoopC).Id).Name
                Else
                    strTeam = strTeam & "-" & UserList(.Users(LoopC).Id).Name
                End If
                
                .Users(LoopC).value = 1
                .Users(LoopC).MapFight = ArenaSlot

            End If
        Next LoopC
    End With
End Sub

Private Function Fight_Search_Enfrentamiento(ByVal Userindex As Integer, ByVal SlotEvent As Byte) As Byte
    ' Chequeamos que tengamos contrincante para luchar.
    Dim LoopC As Integer
    
    Fight_Search_Enfrentamiento = 0
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users()) Step .TeamCant
            If .Users(LoopC).Id > 0 And .Users(LoopC).value = 0 Then
                If .Users(LoopC).Id <> Userindex Then
                    Fight_Search_Enfrentamiento = .Users(LoopC).Team
                    Exit For
                End If
            End If
        Next LoopC
    
    End With
End Function
Private Sub Fight_Combate(ByVal SlotEvent As Byte)
    ' Buscamos una arena disponible y mandamos la mayor cantidad de usuarios disponibles.
    Dim LoopC As Integer
    Dim FreeArena As Byte
    Dim OponentTeam As Byte
    Dim strTemp As String
    Dim strTeam1 As String
    Dim strTeam2 As String
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users()) Step .TeamCant
            If .Users(LoopC).Id > 0 And .Users(LoopC).value = 0 Then
                FreeArena = FreeSlotArena()
                
                If FreeArena > 0 Then
                    OponentTeam = Fight_Search_Enfrentamiento(.Users(LoopC).Id, SlotEvent)
                    
                    If OponentTeam > 0 Then
                        Fight_WarpTeam SlotEvent, FreeArena, .Users(LoopC).Team, False, strTeam1
                        Fight_WarpTeam SlotEvent, FreeArena, OponentTeam, True, strTeam2
                        MapEvent.Fight(FreeArena).Run = True
                        
                        strTemp = "Enfrentamientos " & Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant & "» "
                        strTemp = strTemp & strTeam1 & " vs " & strTeam2
                        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strTemp, FontTypeNames.FONTTYPE_GUILD)
                        
                        strTemp = vbNullString
                        strTeam1 = vbNullString
                        strTeam2 = vbNullString
                    End If
                End If
            End If
        Next LoopC
        

        
    End With
End Sub
Private Function CheckTeam_UserDie(ByVal SlotEvent As Integer, ByVal TeamUser As Byte) As Boolean

    Dim LoopC As Integer
    ' Encontramos a uno del Team vivo, significa que no hay terminación del duelo.
    CheckTeam_UserDie = True
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                If .Users(LoopC).Team = TeamUser Then
                    If UserList(.Users(LoopC).Id).flags.Muerto = 0 Then
                        CheckTeam_UserDie = False
                        Exit Function
                    End If
                End If
            End If
        Next LoopC
    
    End With
End Function

Public Sub Fight_UserDie(ByVal SlotEvent As Byte, ByVal SlotUserEvent As Byte, ByVal TeamWin As Byte)
        
    Dim TeamSlot As Byte
    Dim LoopC As Integer
    Dim strTempWin As String
    
    
    With Events(SlotEvent)
        TeamSlot = .Users(SlotUserEvent).Team
        
        If CheckTeam_UserDie(SlotEvent, TeamSlot) = False Then Exit Sub
        
        AbandonateEvent .Users(SlotUserEvent).Id
            
        For LoopC = LBound(.Users()) To UBound(.Users())
            With .Users(LoopC)
                If .Team = TeamWin Then
                    If strTempWin = vbNullString Then
                        strTempWin = UserList(.Id).Name
                    Else
                        strTempWin = strTempWin & "-" & UserList(.Id).Name
                    End If
                    
                    MapEvent.Fight(.MapFight).Run = False
                    .value = 0
                    .MapFight = 0
                    EventWarpUser .Id, 211, 30, 21
                    WriteConsoleMsg .Id, "Felicitaciones. Has ganado el enfrentamiento", FontTypeNames.FONTTYPE_INFO
                    
                End If
            End With
        Next LoopC
        
        If strTempWin <> vbNullString Then SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Enfrentamientos " & Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant & "» Enfrentamiento ganado por " & strTempWin & ".", FontTypeNames.FONTTYPE_GUILD)
        
        'El -1 es ya que luego de este procedimiento, el perdedor abandona el evento.
        If .TeamCant = .Inscribed Then
            Fight_SearchTeamWin SlotEvent
            CloseEvent SlotEvent
        Else
            Fight_Combate SlotEvent
        End If
        
    End With
End Sub

Private Sub Fight_SearchTeamWin(ByVal SlotEvent As Byte)
    Dim LoopC As Integer
    Dim strTemp As String
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                WriteConsoleMsg .Users(LoopC).Id, "Has ganado el evento.", FontTypeNames.FONTTYPE_INFO
                
                If strTemp = vbNullString Then
                    strTemp = UserList(.Users(LoopC).Id).Name
                Else
                    strTemp = strTemp & ", " & UserList(.Users(LoopC).Id).Name
                End If
            End If
        Next LoopC
    
    
    If .TeamCant > 1 Then
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Enfrentamientos " & .TeamCant & "vs" & .TeamCant & "» Ganadores " & strTemp, FontTypeNames.FONTTYPE_GUILD)
    Else
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Enfrentamientos " & .TeamCant & "vs" & .TeamCant & "» Ganador " & strTemp, FontTypeNames.FONTTYPE_GUILD)
    End If
    
    End With
End Sub


' ############################## USUARIO UNSTOPPABLE ###########################################
Public Sub InitUnstoppable(ByVal SlotEvent As Byte)
    Dim LoopC As Integer
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                EventWarpUser .Users(LoopC).Id, 218, RandomNumber(30, 54), RandomNumber(25, 39)
                
            End If
        Next LoopC
        
        .TimeCount = 10
        .TimeFinish = 60 + .TimeCount
    End With
End Sub
Public Sub Unstoppable_Userdie(ByVal SlotEvent As Byte, ByVal VictimSlot As Byte, ByVal AttackerSlot As Byte)
    With Events(SlotEvent)
        With .Users(VictimSlot)
            Call EventWarpUser(.Id, 218, RandomNumber(30, 54), RandomNumber(25, 39))
            Call RevivirUsuario(.Id)
            Call WriteConsoleMsg(.Id, "Has sido aniquilado. Pero no pierdas las esperanzas joven guerrero, reviviste y tu sangre está hambrienta, ve trás el que te asesino y haz justicia!", FontTypeNames.FONTTYPE_FIGHT)
        End With
        
        With .Users(AttackerSlot)
            .value = .value + 1
            WriteConsoleMsg .Id, "Felicitaciones, has sumado una muerte más a tu lista. Actualmente llevas " & .value & " asesinatos. Sigue así y ganarás el evento.", FontTypeNames.FONTTYPE_INFO
        End With
    End With
End Sub

Private Sub Unstoppable_UserWin(ByVal SlotEvent As Byte)
    Dim Userindex As Integer
    
    Event_OrdenateUsersValue SlotEvent
    
    Userindex = Events(SlotEvent).Users(1).Id
    
    With UserList(Userindex)
        WriteConsoleMsg Userindex, "Felicitaciones. Tus " & Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).value & " asesinatos han hecho que ganes el evento. Aquí tienes 500.000 monedas de oro como recompensa.", FontTypeNames.FONTTYPE_INFO
        .Stats.Gld = .Stats.Gld + 500000
        
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Usuario Unstoppable» El ganador del evento es " & .Name & " con " & Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).value & " asesinatos.", FontTypeNames.FONTTYPE_GUILD)
        CloseEvent SlotEvent
    End With
End Sub
Private Function Event_OrdenateUsersValue(ByVal SlotEvent As Byte) As Integer
    ' Utilizados para buscar ganador según VALUE
    Dim LoopX As Integer
    Dim LoopY As Integer
    Dim aux As tUserEvent

    With Events(SlotEvent)
        For LoopX = LBound(.Users()) To UBound(.Users())
            For LoopY = LBound(.Users()) To UBound(.Users()) - 1
                If .Users(LoopY).value < .Users(LoopY + 1).value Then
                    aux = .Users(LoopY)
                    .Users(LoopY) = .Users(LoopY + 1)
                    .Users(LoopY + 1) = aux
                End If
            Next LoopY
        Next LoopX
    End With
End Function

