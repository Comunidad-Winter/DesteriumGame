Attribute VB_Name = "mStreamer"
Option Explicit

' Modulo encargado de tener un GM BOT que vaya siendo sumoneado de manera automática por los diferentes eventos del juego.

' @ Tiempo entre un Summon y OTRO [VALOR DEFAULT]
Public Const STREAMER_TIME_AUTO_WARP  As Long = 10000 ' 6s

' @ Tiempo para que el mismo usuario sea buscado.
Public Const STREAMER_TIME_CAN_SEARCH As Long = 120000 ' 30s

' @ Maximo de usuarios que el stream va a tener en la lista para seguir.
Public Const MAX_STREAM_USERS         As Byte = 50

Public Enum eStreamerMode

    eZonaSegura = 1         ' Busca personajes en zona insegura
    eEventos = 2               ' Eventos automáticos
    eRetos = 3                  ' Retos
    eRetosRapidos = 4       ' Retos Rapido
    eBuscadorAgites = 5     ' Buscador de Agites en ZONA INSEGURA
    eMixed = 6                  ' Realiza un MIXED con orden de prioridad.
    eUserList = 7               ' Busca según la lista de usuarios que decidió ser seguida por el BOT

    e_LAST = 8

End Enum

Public Type tStreamer

    Active As Integer ' Determina el Index del GM BOT (UserIndex)
    InitialPosition As WorldPos
    LastSummon As Long
    LastTarget As String
    UserIndex As Integer
    Last As Long
    Mode As eStreamerMode
        
    Config_TimeWarp As Long
    Config_TimeCanIndex As Long
        
    Users(1 To MAX_STREAM_USERS) As Integer     ' Usuarios que solicitaron al STREAMBOT

End Type

Public Const STREAMER_MAX_BOTS As Byte = 10

Public StreamerBot             As tStreamer

' @ Inicializa al GM BOT, con una posición de respawn general.
Public Sub Streamer_Can(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo Streamer_Can_Err

    '</EhHeader>
    
    If Not EsGm(UserIndex) Then Exit Sub
    
    With UserList(UserIndex)

        If StreamerBot.Active Then
            Streamer_Initial 0, 0, 0, 0
        Else
            Streamer_Initial UserIndex, .Pos.Map, .Pos.X, .Pos.Y

        End If

    End With
    
    '<EhFooter>
    Exit Sub

Streamer_Can_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mStreamer.Streamer_Can " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Streamer_Initial(ByVal UserIndex As Integer, _
                            ByVal Map As Integer, _
                            ByVal X As Byte, _
                            ByVal Y As Byte)

    '<EhHeader>
    On Error GoTo Streamer_Initial_Err

    '</EhHeader>
    
    Dim tUser As Integer
            
    With StreamerBot

        If Map > 0 Then
            If .Active > 0 Then
                Call WriteConsoleMsg(UserIndex, "¡Está siendo utilizado por otro!", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub

            End If

        End If
              
        .Active = UserIndex
        .InitialPosition.Map = Map
        .InitialPosition.X = X
        .InitialPosition.Y = Y

        .Config_TimeWarp = STREAMER_TIME_AUTO_WARP
        .Config_TimeCanIndex = STREAMER_TIME_CAN_SEARCH
        
        If .Active > 0 Then
            Call Streamer_CheckPosition

        End If

    End With

    '<EhFooter>
    Exit Sub

Streamer_Initial_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mStreamer.Streamer_Initial " & "at line " & Erl
        
    '</EhFooter>
End Sub

' @ Busca el UserIndex en EVENTOS AUTOMATICOS
Private Function Streamer_Search_Event(ByVal Time As Long) As Integer

    '<EhHeader>
    On Error GoTo Streamer_Search_Event_Err

    '</EhHeader>

    Dim A          As Long, B As Long

    Dim BestTarget As Integer
        
    For A = 1 To MAX_EVENT_SIMULTANEO

        With Events(A)

            If .Run Then
                    
                For B = LBound(.Users) To UBound(.Users)

                    If .Users(B).ID > 0 Then
                        If (Time - UserList(.Users(B).ID).Counters.TimeGMBOT) >= StreamerBot.Config_TimeCanIndex And UserList(.Users(B).ID).flags.Muerto = 0 Then
                                
                            Streamer_Search_Event = .Users(B).ID
                            StreamerBot.LastSummon = Time
                            UserList(.Users(B).ID).Counters.TimeGMBOT = Time
                            StreamerBot.LastTarget = UCase$(UserList(.Users(B).ID).Name)
                            Exit Function

                        End If

                    End If

                Next B

            End If

        End With
    
    Next A

    '<EhFooter>
    Exit Function

Streamer_Search_Event_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mStreamer.Streamer_Search_Event " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Function

' @ Busca el UserIndex en RETOS
Private Function Streamer_Search_Fight(ByVal Time As Long) As Integer

    '<EhHeader>
    On Error GoTo Streamer_Search_Fight_Err

    '</EhHeader>

    Dim A As Long, B As Long

    For A = 1 To MAX_RETOS_SIMULTANEOS

        With Retos(A)

            If .Run Then

                For B = LBound(.User) To UBound(.User)

                    If .User(B).UserIndex > 0 Then
                        If (Time - UserList(.User(B).UserIndex).Counters.TimeGMBOT) >= StreamerBot.Config_TimeCanIndex And UserList(.User(B).UserIndex).flags.Muerto = 0 Then
                            Streamer_Search_Fight = .User(B).UserIndex
                            StreamerBot.LastSummon = Time
                            UserList(.User(B).UserIndex).Counters.TimeGMBOT = Time
                            StreamerBot.LastTarget = UCase$(UserList(.User(B).UserIndex).Name)
                            Exit Function

                        End If
                    
                    End If

                Next B
            
            End If

        End With
    
    Next A

    '<EhFooter>
    Exit Function

Streamer_Search_Fight_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mStreamer.Streamer_Search_Fight " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Function

' @ Busca el UserIndex en RETOS RAPIDOS
Private Function Streamer_Search_FightFast(ByVal Time As Long) As Integer

    '<EhHeader>
    On Error GoTo Streamer_Search_FightFast_Err

    '</EhHeader>

    Dim A As Long, B As Long
    
    For A = 1 To MAX_RETO_FAST

        With RetoFast(A)

            If .Run Then

                For B = LBound(.Users) To UBound(.Users)
    
                    If .Users(B) > 0 Then
                        If (Time - UserList(.Users(B)).Counters.TimeGMBOT) >= StreamerBot.Config_TimeCanIndex And UserList(.Users(B)).flags.Muerto = 0 Then
                            Streamer_Search_FightFast = .Users(B)
                            StreamerBot.LastSummon = Time
                            UserList(.Users(B)).Counters.TimeGMBOT = Time
                            StreamerBot.LastTarget = UCase$(UserList(.Users(B)).Name)
                            Exit Function
    
                        End If
    
                    End If
    
                Next B

            End If

        End With
    
    Next A

    '<EhFooter>
    Exit Function

Streamer_Search_FightFast_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mStreamer.Streamer_Search_FightFast " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Function

' @ Buscamos usuarios en Zona Segura . @ NO WORKS
Private Function Streamer_Search_Secure(ByVal Time As Long) As Integer

    '<EhHeader>
    On Error GoTo Streamer_Search_Secure_Err

    '</EhHeader>
        
    Dim A As Long
    
    For A = 1 To LastUser

        With UserList(A)

            If (.ConnIDValida) Then
                If .flags.UserLogged Then
                    If Not EsGm(A) Then
                        If (Time - UserList(A).Counters.TimeGMBOT) >= StreamerBot.Config_TimeCanIndex Then
                            If Not MapInfo(.Pos.Map).Pk And UserList(A).flags.Muerto = 0 Then
                                Streamer_Search_Secure = A
                                StreamerBot.LastSummon = Time
                                UserList(A).Counters.TimeGMBOT = Time
                                StreamerBot.LastTarget = UCase$(.Name)
                                Exit Function

                            End If

                        End If

                    End If

                End If

            End If
            
        End With

    Next A
    
    '<EhFooter>
    Exit Function

Streamer_Search_Secure_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mStreamer.Streamer_Search_Secure " & "at line " & Erl
        
    '</EhFooter>
End Function

' @ Buscamos agites
Private Function Streamer_Search_Insecure(ByVal Time As Long) As Integer

    '<EhHeader>
    On Error GoTo Streamer_Search_Insecure_Err
        
    '</EhHeader>
        
    Dim A As Long
    
    For A = 1 To LastUser

        With UserList(A)

            If (.ConnIDValida) Then
                If .flags.UserLogged Then
                    If (Time - UserList(A).Counters.TimeGMBOT) >= StreamerBot.Config_TimeCanIndex Then
                        If MapInfo(.Pos.Map).Pk And MapInfo(.Pos.Map).NumUsers >= 7 And UserList(A).flags.Muerto = 0 Then
                            Streamer_Search_Insecure = A
                            UserList(A).Counters.TimeGMBOT = Time
                            StreamerBot.LastSummon = Time
                            StreamerBot.LastTarget = UCase$(.Name)
                            Exit Function

                        End If

                    End If

                End If

            End If
            
        End With

    Next A
    
    '<EhFooter>
    Exit Function

Streamer_Search_Insecure_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mStreamer.Streamer_Search_Insecure_Err " & "at line " & Erl
        
    '</EhFooter>
End Function

' @ Buscamos uno de los usuarios que haya pedido ser seguido...
Private Function Streamer_Search_Users(ByVal Time As Long) As Integer

    '<EhHeader>
    On Error GoTo Streamer_Search_Users_Err

    '</EhHeader>
        
    Dim A         As Long

    Dim UserIndex As Integer
    
    For A = 1 To MAX_STREAM_USERS
        UserIndex = StreamerBot.Users(A)

        If UserIndex > 0 Then
            
            If (Time - UserList(UserIndex).Counters.TimeGMBOT) >= StreamerBot.Config_TimeCanIndex And UserList(UserIndex).flags.Muerto = 0 Then
            
                Streamer_Search_Users = UserIndex
                UserList(UserIndex).Counters.TimeGMBOT = Time
                StreamerBot.LastSummon = Time
                StreamerBot.LastTarget = UCase$(UserList(UserIndex).Name)
                Exit Function
                    
            End If

        End If
        
    Next A
    
    '<EhFooter>
    Exit Function

Streamer_Search_Users_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mStreamer.Streamer_Search_Users " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Function

' @ Agrega un usuario a la lista
Public Sub Streamer_RequiredBOT(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo Streamer_RequiredBOT_Err

    '</EhHeader>
    
    Dim Slot As Byte
    
    If StreamerBot.Active = 0 Then
        Call WriteConsoleMsg(UserIndex, "El Hamster del CPU está descansando. Solicita nuestra pantalla LITOMANIA en otro momento.", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
    
    With UserList(UserIndex)

        If .flags.BotList > 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Ya te encuentras en la lista de búsqueda del GM! Sal del Juego para no estarlo.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If
        
        Slot = Streamer_Required_Slot
        
        If Slot > 0 Then
            Call Streamer_SetBotList(UserIndex, Slot, False)
            Call WriteConsoleMsg(UserIndex, "Te he agregado a mi lista... podrías ser el próximo ¡Asi que muestra algo o me iré!", FontTypeNames.FONTTYPE_INFOGREEN)
        Else
            Call WriteConsoleMsg(UserIndex, "¡Vaya! Que solicitado soy... Espera un momento que renuevo la lista y vuelve a intentar pronto. Podré seguirte y ni te darás cuenta!", FontTypeNames.FONTTYPE_INFORED)
            
        End If
    
    End With

    '<EhFooter>
    Exit Sub

Streamer_RequiredBOT_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mStreamer.Streamer_RequiredBOT " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Public Sub Streamer_SetBotList(ByVal UserIndex As Integer, _
                               ByVal Slot As Byte, _
                               ByVal Killed As Boolean)

    '<EhHeader>
    On Error GoTo Streamer_SetBotList_Err

    '</EhHeader>

    StreamerBot.Users(Slot) = IIf((Killed = True), 0, UserIndex)
    UserList(UserIndex).flags.BotList = IIf((Killed = True), 0, Slot)
    
    '<EhFooter>
    Exit Sub

Streamer_SetBotList_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mStreamer.Streamer_SetBotList " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

' @ Busca un SLOT libre
Private Function Streamer_Required_Slot() As Byte

    '<EhHeader>
    On Error GoTo Streamer_Required_Slot_Err

    '</EhHeader>
    Dim A As Long
    
    For A = 1 To MAX_STREAM_USERS

        If StreamerBot.Users(A) = 0 Then
            Streamer_Required_Slot = A
            Exit Function

        End If

    Next A
    
    '<EhFooter>
    Exit Function

Streamer_Required_Slot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mStreamer.Streamer_Required_Slot " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Function

' @ Busca al proximo usuario disponible para tomar su posición
Public Function Streamer_Search_UserIndex(ByVal Time As Long) As Integer

    '<EhHeader>
    On Error GoTo Streamer_Search_UserIndex_Err

    '</EhHeader>

    Dim A As Long

    Dim B As Long
    
    With StreamerBot
         
        Select Case .Mode
        
            Case eStreamerMode.eZonaSegura
                Streamer_Search_UserIndex = Streamer_Search_Secure(Time)
                
            Case eStreamerMode.eEventos
                Streamer_Search_UserIndex = Streamer_Search_Event(Time)
                
            Case eStreamerMode.eRetos
                Streamer_Search_UserIndex = Streamer_Search_Fight(Time)
                
            Case eStreamerMode.eRetosRapidos
                Streamer_Search_UserIndex = Streamer_Search_FightFast(Time)
                
            Case eStreamerMode.eBuscadorAgites
                ' @ Realizar comprobaciones de lanzamiento de hechizos y golpes
                Streamer_Search_UserIndex = Streamer_Search_Insecure(Time)
                
            Case eStreamerMode.eUserList
                Streamer_Search_UserIndex = Streamer_Search_Users(Time)
                     
            Case eStreamerMode.eMixed
                ' Ordenar según la prioridad
                Streamer_Search_UserIndex = Streamer_Search_Event(Time): If Streamer_Search_UserIndex > 0 Then Exit Function
                Streamer_Search_UserIndex = Streamer_Search_FightFast(Time): If Streamer_Search_UserIndex > 0 Then Exit Function
                Streamer_Search_UserIndex = Streamer_Search_Fight(Time): If Streamer_Search_UserIndex > 0 Then Exit Function
                Streamer_Search_UserIndex = Streamer_Search_Users(Time): If Streamer_Search_UserIndex > 0 Then Exit Function
                Streamer_Search_UserIndex = Streamer_Search_Insecure(Time): If Streamer_Search_UserIndex > 0 Then Exit Function
                Streamer_Search_UserIndex = Streamer_Search_Secure(Time): If Streamer_Search_UserIndex > 0 Then Exit Function

        End Select

    End With
    
    '<EhFooter>
    Exit Function

Streamer_Search_UserIndex_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mStreamer.Streamer_Search_UserIndex " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub Streamer_Sum(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler
    
    Dim X As Long

    Dim Y As Long
    
    With StreamerBot
    
        X = RandomNumber(UserList(UserIndex).Pos.X - 3, UserList(UserIndex).Pos.X + 3)
        Y = RandomNumber(UserList(UserIndex).Pos.Y - 3, UserList(UserIndex).Pos.Y + 3)
        Call EventWarpUser(.Active, UserList(UserIndex).Pos.Map, X, Y)

    End With
    
    Exit Sub
ErrHandler:

End Sub

' @ Reinicia cuando lo necesite
Public Sub Streamer_CheckUser(ByVal UserIndex As Integer)
    
    On Error GoTo ErrHandler
    
    Dim Time As Long
    
    With StreamerBot
    
        If .Active = 0 Then Exit Sub
        
        Time = GetTime
        
        If .UserIndex = UserIndex Then
            .UserIndex = Streamer_Search_UserIndex(Time)
            
            Call EventWarpUser(.Active, .InitialPosition.Map, .InitialPosition.X, .InitialPosition.Y)
            .LastSummon = Time

        End If
        
    End With
    
    Exit Sub
ErrHandler:
    
End Sub

' @ Comprueba que tan lejos se fue del objetivo
Public Sub Streamer_CheckPosition()

    '<EhHeader>
    On Error GoTo Streamer_CheckPosition_Err

    '</EhHeader>
    Dim Time As Double

    Dim X    As Integer, Y As Integer
              
    Time = GetTime
            
    Static SecondsCheckCercania As Integer
        
    With StreamerBot
    
        ' @ El BOT no está activo.
        If .Active = 0 Then Exit Sub
            
        SecondsCheckCercania = SecondsCheckCercania + 1
            
        If SecondsCheckCercania >= 2 Then
            
            If .UserIndex > 0 Then

                With UserList(.UserIndex)

                    If Distance(UserList(StreamerBot.Active).Pos.X, UserList(StreamerBot.Active).Pos.Y, .Pos.X, .Pos.Y) >= 6 Then
                        Streamer_Sum StreamerBot.UserIndex

                    End If

                End With

            End If
                
            SecondsCheckCercania = 0

        End If
             
        ' @ Segun el Tiempo SETEADO entre WARP & WARP
        If (Time - .LastSummon) < .Config_TimeWarp Then Exit Sub

        ' @ Esto de abajo es llamado respetando cada 40s
        Dim UserIndex As Integer, Pos As WorldPos

        UserIndex = Streamer_Search_UserIndex(Time)
            
        If UserIndex > 0 Then
            .UserIndex = UserIndex
            Call EventWarpUser(.Active, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
            .LastSummon = Time
        Else
            .UserIndex = 0
                 
            If Len(.LastTarget) = 0 Then
                If Distance(.InitialPosition.X, .InitialPosition.Y, UserList(StreamerBot.Active).Pos.X, UserList(StreamerBot.Active).Pos.Y) > 7 Then
                    Call EventWarpUser(.Active, .InitialPosition.Map, .InitialPosition.X, .InitialPosition.Y)
                    .LastSummon = Time

                End If

            End If
            
        End If
   
    End With
    
    '<EhFooter>
    Exit Sub

Streamer_CheckPosition_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mStreamer.Streamer_CheckPosition " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function Streamer_Mode_String(ByRef Mode As eStreamerMode) As String

    Select Case Mode
        
        Case eStreamerMode.eBuscadorAgites
            Streamer_Mode_String = "Agites en Zona Insegura"
            
        Case eStreamerMode.eEventos
            Streamer_Mode_String = "Eventos automáticos"
            
        Case eStreamerMode.eMixed
            Streamer_Mode_String = "Modalidad Mixed. Busca el mejor emparejamiento interno."
                
        Case eStreamerMode.eRetos
            Streamer_Mode_String = "Retos privados"
            
        Case eStreamerMode.eRetosRapidos
            Streamer_Mode_String = "Retos rapidos"
            
        Case eStreamerMode.eZonaSegura
            Streamer_Mode_String = "Usuarios en Zona segura. NO trabajadores."
        
    End Select

End Function
