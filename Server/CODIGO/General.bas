Attribute VB_Name = "General"
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

Option Explicit

Public Running         As Boolean

Private Const MAX_TIME As Double = 2147483647 ' 2^31

Private LastTick       As Double

Private overflowCount  As Long

Global LeerNPCs        As clsIniManager

Function DarCuerpoDesnudo_Genero(ByVal UserGenero As Byte, _
                                 ByVal UserRaza As Byte, _
                                 Optional ByVal Mimetizado As Boolean = False) As Integer

    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 03/14/07
    'Da cuerpo desnudo a un usuario
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************
    '<EhHeader>
    On Error GoTo DarCuerpoDesnudo_Err

    '</EhHeader>

    Dim CuerpoDesnudo As Integer

    Select Case UserGenero

        Case eGenero.Hombre

            Select Case UserRaza

                Case eRaza.Humano
                    CuerpoDesnudo = 21

                Case eRaza.Drow
                    CuerpoDesnudo = 32

                Case eRaza.Elfo
                    CuerpoDesnudo = 21

                Case eRaza.Gnomo
                    CuerpoDesnudo = 53

                Case eRaza.Enano
                    CuerpoDesnudo = 53

            End Select

        Case eGenero.Mujer

            Select Case UserRaza

                Case eRaza.Humano
                    CuerpoDesnudo = 39

                Case eRaza.Drow
                    CuerpoDesnudo = 40

                Case eRaza.Elfo
                    CuerpoDesnudo = 39

                Case eRaza.Gnomo
                    CuerpoDesnudo = 60

                Case eRaza.Enano
                    CuerpoDesnudo = 60

            End Select

    End Select

    DarCuerpoDesnudo_Genero = CuerpoDesnudo
    '<EhFooter>
    Exit Function

DarCuerpoDesnudo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.DarCuerpoDesnudo " & "at line " & Erl

    '</EhFooter>
End Function

Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, _
                     Optional ByVal Mimetizado As Boolean = False)

    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 03/14/07
    'Da cuerpo desnudo a un usuario
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************
    '<EhHeader>
    On Error GoTo DarCuerpoDesnudo_Err

    '</EhHeader>

    Dim CuerpoDesnudo As Integer

    With UserList(UserIndex)

        Select Case .Genero

            Case eGenero.Hombre

                Select Case .Raza

                    Case eRaza.Humano
                        CuerpoDesnudo = 21

                    Case eRaza.Drow
                        CuerpoDesnudo = 32

                    Case eRaza.Elfo
                        CuerpoDesnudo = 21

                    Case eRaza.Gnomo
                        CuerpoDesnudo = 53

                    Case eRaza.Enano
                        CuerpoDesnudo = 53

                End Select

            Case eGenero.Mujer

                Select Case .Raza

                    Case eRaza.Humano
                        CuerpoDesnudo = 39

                    Case eRaza.Drow
                        CuerpoDesnudo = 40

                    Case eRaza.Elfo
                        CuerpoDesnudo = 39

                    Case eRaza.Gnomo
                        CuerpoDesnudo = 60

                    Case eRaza.Enano
                        CuerpoDesnudo = 60

                End Select

        End Select
          
        If Mimetizado Then
            .CharMimetizado.Body = CuerpoDesnudo
        Else
            .Char.Body = CuerpoDesnudo

        End If
          
        .flags.Desnudo = 1

    End With

    '<EhFooter>
    Exit Sub

DarCuerpoDesnudo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.DarCuerpoDesnudo " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub Bloquear(ByVal toMap As Boolean, _
             ByVal sndIndex As Integer, _
             ByVal X As Integer, _
             ByVal Y As Integer, _
             ByVal B As Boolean)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'b ahora es boolean,
    'b=true bloquea el tile en (x,y)
    'b=false desbloquea el tile en (x,y)
    'toMap = true -> Envia los datos a todo el mapa
    'toMap = false -> Envia los datos al user
    'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
    'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s
    '***************************************************
    '<EhHeader>
    On Error GoTo Bloquear_Err

    '</EhHeader>

    If toMap Then
        Call SendData(SendTarget.toMap, sndIndex, PrepareMessageBlockPosition(X, Y, B))
    Else
        Call WriteBlockPosition(sndIndex, X, Y, B)

    End If

    '<EhFooter>
    Exit Sub

Bloquear_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.Bloquear " & "at line " & Erl
        
    '</EhFooter>
End Sub

Function HayAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo HayAgua_Err

    '</EhHeader>

    If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then

        With MapData(Map, X, Y)

            If ((.Graphic(1) >= 1505 And .Graphic(1) <= 1520)) And .Graphic(2) = 0 Then
                
                HayAgua = True
            Else
                HayAgua = False

            End If

        End With

    Else
        HayAgua = False

    End If

    '<EhFooter>
    Exit Function

HayAgua_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.HayAgua " & "at line " & Erl
        
    '</EhFooter>
End Function

Private Function HayLava(ByVal Map As Integer, _
                         ByVal X As Integer, _
                         ByVal Y As Integer) As Boolean

    '<EhHeader>
    On Error GoTo HayLava_Err

    '</EhHeader>

    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 03/12/07
    '***************************************************
    If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
        If (MapData(Map, X, Y).Graphic(1) >= 5837 And MapData(Map, X, Y).Graphic(1) <= 5852) Or MapData(Map, X, Y).trigger = eTrigger.LavaActiva Then
            HayLava = True
        Else
            HayLava = False

        End If

    Else
        HayLava = False

    End If

    '<EhFooter>
    Exit Function

HayLava_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.HayLava " & "at line " & Erl
        
    '</EhFooter>
End Function

Sub LimpiarMundo()

    'SecretitOhs
    '<EhHeader>
    On Error GoTo LimpiarMundo_Err

    '</EhHeader>
    'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpiando Mundo.", FontTypeNames.FONTTYPE_SERVER))

    Dim MapaActual As Long

    Dim Y          As Long

    Dim X          As Long

    Dim bIsExit    As Boolean

    For MapaActual = 1 To NumMaps

        For Y = YMinMapSize To YMaxMapSize

            For X = XMinMapSize To XMaxMapSize

                If MapData(MapaActual, X, Y).ObjInfo.ObjIndex > 0 And MapData(MapaActual, X, Y).Blocked = 0 And MapInfo(MapaActual).Limpieza = 1 Then
                    If (GetTime - MapData(MapaActual, X, Y).Protect) >= 60000 Then
                        
                        If ItemNoEsDeMapa(MapaActual, X, Y, True) Then Call EraseObj(10000, MapaActual, X, Y)

                    End If

                End If

            Next X

        Next Y

    Next MapaActual

    'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Mundo limpiado.", FontTypeNames.FONTTYPE_SERVER))
    '<EhFooter>
    Exit Sub

LimpiarMundo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.LimpiarMundo " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub EnviarSpawnList(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo EnviarSpawnList_Err

    '</EhHeader>

    Dim K          As Long

    Dim npcNames() As String
    
    ReDim npcNames(1 To UBound(SpawnList)) As String
    
    For K = 1 To UBound(SpawnList)
        npcNames(K) = SpawnList(K).NpcName
    Next K
    
    Call WriteSpawnList(UserIndex, npcNames())

    '<EhFooter>
    Exit Sub

EnviarSpawnList_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.EnviarSpawnList " & "at line " & Erl
        
    '</EhFooter>
End Sub

' # Encripta fácilmente según la hora de la PC.
Private Function Encrypt_Value(ByVal Value As Long) As Long
    Encrypt_Value = Value Xor GetTime

End Function

Sub Main()

    '<EhHeader>
    On Error GoTo Main_Err

    '</EhHeader>

    Static variable As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: 15/03/2011
    '15/03/2011: ZaMa - Modularice todo, para que quede mas claro.
    '***************************************************

    ChDir App.Path
    ChDrive App.Path
    
    GlobalActive = True
    Call LoadMotd
    Call BanIpCargar
    Call AutoBan_Initialize
    Call Challenge_SetMap
          
    PacketUseItem = ClientPacketID.UseItem
      
    ReDim ListMails(0) As String
    
    frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    
    ' Start loading..
    frmCargando.Show
    
    ' Constants & vars
    frmCargando.Label1(2).Caption = "Cargando constantes..."
    Call LoadConstants
    DoEvents
    
    ' Arrays
    frmCargando.Label1(2).Caption = "Iniciando Arrays..."
    Call LoadArrays
    
    ' Cargamos base de datps
    'Call MySql_Open
        
    ' Server.ini & Apuestas.dat
    frmCargando.Label1(2).Caption = "Cargando Server.ini"
    Call LoadSini
    Call CargaApuestas
    
    ' Npcs_FilePath
    frmCargando.Label1(2).Caption = "Cargando Criaturas"
    Call CargaNpcsDat

    ' Objs_FilePath
    frmCargando.Label1(2).Caption = "Cargando Objetos"
    Call LoadOBJData

    ' Shop Items
    frmCargando.Label1(2).Caption = "Cargando Shop"
    Call Shop_Load
    Call Shop_Load_Chars
    
    ' Quests
    Call LoadQuests
    
    ' Spell_FilePath
    frmCargando.Label1(2).Caption = "Cargando Hechizos"
    Call CargarHechizos
        
    ' Cargamos el mercado
    Call mMao.Mercader_Load
    
    ' Balance.dat
    frmCargando.Label1(2).Caption = "Cargando Balance.Dat"
    Call LoadBalance
    
    ' Animaciones
    frmCargando.Label1(2).Caption = "Cargando Animaciones"
    Call LoadAnimations
        
    ' Mapas
    If BootDelBackUp Then
        frmCargando.Label1(2).Caption = "Cargando BackUp"
        Call CargarBackUp
    Else
        frmCargando.Label1(2).Caption = "Cargando Mapas"
        Call LoadMapData

    End If
        
    Call DataServer_Generate_ObjData
        
    Call Castle_Load
        
    ' Ruletas
    Call Ruleta_LoadItems
        
    ' Set de comerciantes en mapit
    Call Comerciantes_Load
        
    ' Generamos la info exportable al cliente de mapas
    Call DataServer_Generate_Maps
        
    ' Pathfinding
    Call InitPathFinding
         
    ' Eventos automáticos
    Call LoadMapEvent
    
    ' Load Invocations
    Call LoadInvocaciones
    
    ' Load Global Drops
    Call Drops_Load
            
    ' Cargamos las facciones
    Call LoadFactions
    
    ' Cargamos los RetoFast
    Call LoadRetoFast
        
    ' Eventos AI
    Call Events_Load_PreConfig
        
    ' Pretorianos
    frmCargando.Label1(2).Caption = "Cargando Pretorianos.dat"
    Call LoadPretorianData
    
    ' Map Sounds
    Set SonidosMapas = New SoundMapInfo
    Call SonidosMapas.LoadSoundMapInfo
    
    ' Home distance
    Call generateMatrix(MATRIX_INITIAL_MAP)
    
    ' Connections
    Call ResetUsersConnections
    
    ' Timers
    Call InitMainTimers
    
    ' Sockets
    Call SocketConfig
    
    'Call SocketConfig_Archive
    
    ' End loading..
    Unload frmCargando
    
    'Log start time
    LogServerStartTime
    
    'Ocultar
    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)

    End If
    
    tInicioServer = GetTime
    
    MercaderActivate = True
    
    Running = True

    While (Running)

        Call Server.Poll
        DoEvents
    Wend

    '<EhFooter>
    Exit Sub

Main_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.Main " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub DB_LoadSkills()

    '<EhHeader>
    On Error GoTo DB_LoadSkills_Err

    '</EhHeader>
    Dim Manager As clsIniManager

    Dim A       As Long
    
    Set Manager = New clsIniManager
        
    Manager.Initialize DatPath & "skills.ini"
        
    ' Skills por Nivel que puede ganar el personaje
    For A = 1 To 50
        LevelSkill(A).LevelValue = val(Manager.GetValue("LEVELVALUE", "Lvl" & A))
    Next A

    'NUMSKILLS = val(Manager.GetValue("INIT", "LastSkill"))
    'NUMSKILLSESPECIAL = val(Manager.GetValue("INIT", "LastSkillEspecial"))
            
    '  ReDim InfoSkill(1 To NUMSKILLS) As eInfoSkill
            
    ' Habilidades Cotidianas del Personaje
    For A = 1 To NUMSKILLS
        InfoSkill(A).Name = Manager.GetValue("SK" & A, "Name")
        InfoSkill(A).MaxValue = val(Manager.GetValue("SK" & A, "MaxValue"))
    Next A
    
    ' ReDim InfoSkillEspecial(1 To NUMSKILLSESPECIAL) As eInfoSkill
            
    ' Habilidades Extremas del Personaje
    For A = 1 To NUMSKILLSESPECIAL
        InfoSkillEspecial(A).Name = Manager.GetValue("SKESP" & A, "Name")
        InfoSkillEspecial(A).MaxValue = val(Manager.GetValue("SKESP" & A, "MaxValue"))
    Next A
        
    Set Manager = Nothing
    '<EhFooter>
    Exit Sub

DB_LoadSkills_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.DB_LoadSkills " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Private Sub LoadConstants()

    '<EhHeader>
    On Error GoTo LoadConstants_Err

    '</EhHeader>

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 15/03/2011
    'Loads all constants and general parameters.
    '*****************************************************************
   
    LastBackup = Format(Now, "Short Time")
    Minutos = Format(Now, "Short Time")
    
    ' Paths
    IniPath = App.Path & "\"
    DatPath = App.Path & "\DAT\"
    CharPath = App.Path & "\CHARS\CHARFILE\"
    AccountPath = App.Path & "\CHARS\ACCOUNT\"
    LogPath = App.Path & "\CHARS\LOGS\"
            
    ' Info Skills
    Call DB_LoadSkills
    
    ' Races
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.Drow) = "Drow"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    
    ' Classes
    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
         
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
          
    ' Attributes
    ListaAtributos(eAtributos.Fuerza) = "Fuerza"
    ListaAtributos(eAtributos.Agilidad) = "Agilidad"
    ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
    ListaAtributos(eAtributos.Carisma) = "Carisma"
    ListaAtributos(eAtributos.Constitucion) = "Constitucion"
    
    ' Fishes
    ListaPeces(1) = PECES_POSIBLES.PESCADO1
    ListaPeces(2) = PECES_POSIBLES.PESCADO2
    ListaPeces(3) = PECES_POSIBLES.PESCADO3
    ListaPeces(4) = PECES_POSIBLES.PESCADO4
    ListaPeces(5) = PECES_POSIBLES.PESCADO5
        
    #If Classic = 0 Then
        ListaPeces(6) = PECES_POSIBLES.PESCADO6
        ListaPeces(7) = PECES_POSIBLES.PESCADO7
    #End If
        
    'Bordes del mapa
    MinXBorder = XMinMapSize + (XWindow \ 2)
    MaxXBorder = XMaxMapSize - (XWindow \ 2)
    MinYBorder = YMinMapSize + (YWindow \ 2)
    MaxYBorder = YMaxMapSize - (YWindow \ 2)
    
    Set Denuncias = New cCola
    Denuncias.MaxLenght = MAX_DENOUNCES

    With Prision
        .Map = 21
        .X = 77
        .Y = 15

    End With
    
    With Libertad
        .Map = 21
        .X = 77
        .Y = 29

    End With
            
    MaxUsers = 0

    Set aClon = New clsAntiMassClon
    Set TrashCollector = New Collection

    '<EhFooter>
    Exit Sub

LoadConstants_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.LoadConstants " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub LoadArrays()

    '<EhHeader>
    On Error GoTo LoadArrays_Err

    '</EhHeader>

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 15/03/2011
    'Loads all arrays
    '*****************************************************************

    ' Load Records
    Call LoadRecords
    ' Load guilds info
    Call Guilds_Load
    ' Load spawn list
    Call CargarSpawnList
    ' Load forbidden words
    Call CargarForbidenWords
    ' Load Meditations
    Call Meditation_LoadConfig
    ' Load Ranking
    'Call Load_RankUsers
    'Load Security
    Call Initialize_Security
    ' Cargamos la pesca
    Call Pesca_LoadItems
    ' Invasiones
    Call Invations_Load
    ' Premiums Shop
    'Call Premiums_Load
    Call LoadHelp

    Call BotIntelligence_Load
        
    Call Arenas_Load
    Call MessageSpam_Load
        
    Call CargarFrasesOnFire
    '<EhFooter>
    Exit Sub

LoadArrays_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.LoadArrays " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub ResetUsersConnections()

    '<EhHeader>
    On Error GoTo ResetUsersConnections_Err

    '</EhHeader>

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 15/03/2011
    'Resets Users Connections.
    '*****************************************************************

    Dim LoopC As Long

    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnIDValida = False
    Next LoopC
    
    '<EhFooter>
    Exit Sub

ResetUsersConnections_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.ResetUsersConnections " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub InitMainTimers()

    '<EhHeader>
    On Error GoTo InitMainTimers_Err

    '</EhHeader>

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 15/03/2011
    'Initializes Main Timers.
    '*****************************************************************

    With frmMain
        .TimerGuardarUsuarios.Enabled = True
        .TimerGuardarUsuarios.interval = IntervaloTimerGuardarUsuarios
        .AutoSave.Enabled = True
        .tPiqueteC.Enabled = True
        .GameTimer.Enabled = True
        .FX.Enabled = False
        .Auditoria.Enabled = True
        .KillLog.Enabled = True
        .TIMER_AI.Enabled = True
        .tControlHechizos.Enabled = True
        .tControlHechizos.interval = 60000

    End With
    
    '<EhFooter>
    Exit Sub

InitMainTimers_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.InitMainTimers " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub SocketConfig()

    '<EhHeader>
    On Error GoTo SocketConfig_Err

    '</EhHeader>

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 15/03/2011
    'Sets socket config.
    '*****************************************************************

    Set Writer = New Network.Writer
    Set Server = New Network.Server
    
    Call Server.Attach(AddressOf OnServerConnect, AddressOf OnServerClose, AddressOf OnServerSend, AddressOf OnServerReceive)
    Call Server.Listen(MaxUsers, "0.0.0.0", Puerto)
    
    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    '<EhFooter>
    Exit Sub

SocketConfig_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.SocketConfig " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub LogServerStartTime()

    '<EhHeader>
    On Error GoTo LogServerStartTime_Err

    '</EhHeader>

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 15/03/2011
    'Logs Server Start Time.
    '*****************************************************************
    Dim N As Integer

    N = FreeFile
    Open LogPath & "Main.log" For Append Shared As #N
    Print #N, Date & " " & Time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
    Close #N

    '<EhFooter>
    Exit Sub

LogServerStartTime_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.LogServerStartTime " & "at line " & Erl
        
    '</EhFooter>
End Sub

Function FileExist(ByVal File As String, _
                   Optional FileType As VbFileAttribute = vbNormal) As Boolean

    '*****************************************************************
    'Se fija si existe el archivo
    '*****************************************************************
    '<EhHeader>
    On Error GoTo FileExist_Err

    '</EhHeader>

    FileExist = LenB(dir$(File, FileType)) <> 0
    '<EhFooter>
    Exit Function

FileExist_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.FileExist " & "at line " & Erl
        
    '</EhFooter>
End Function

Function ReadField(ByVal Pos As Integer, _
                   ByRef Text As String, _
                   ByVal SepASCII As Byte) As String

    '*****************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/15/2004
    'Gets a field from a delimited string
    '*****************************************************************
    '<EhHeader>
    On Error GoTo ReadField_Err

    '</EhHeader>

    Dim i          As Long

    Dim lastPos    As Long

    Dim CurrentPos As Long

    Dim delimiter  As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)

    End If

    '<EhFooter>
    Exit Function

ReadField_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.ReadField " & "at line " & Erl
        
    '</EhFooter>
End Function

Function MapaValido(ByVal Map As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo MapaValido_Err

    '</EhHeader>

    MapaValido = Map >= 1 And Map <= NumMaps
    '<EhFooter>
    Exit Function

MapaValido_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.MapaValido " & "at line " & Erl
        
    '</EhFooter>
End Function

Sub MostrarNumUsers()

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo MostrarNumUsers_Err

    '</EhHeader>

    frmMain.txtNumUsers.Text = NumUsers
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageUpdateOnline())
    Call WriteUpdateOnline
    '<EhFooter>
    Exit Sub

MostrarNumUsers_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.MostrarNumUsers " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub Restart()

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo Restart_Err

    '</EhHeader>

    If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."
    
    Dim LoopC As Long

    For LoopC = 1 To MaxUsers
        Protocol.Kick LoopC
    Next

    ReDim UserList(1 To MaxUsers) As User
    
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnIDValida = False
    Next LoopC
    
    LastUser = 0
    NumUsers = 0
    
    Call FreeNPCs
    Call FreeCharIndexes
    
    Call LoadSini
    
    Call LoadOBJData
    
    Call LoadMapData
    
    Call CargarHechizos
    
    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    
    'Log it
    Dim N As Integer

    N = FreeFile
    Open LogPath & "Main.log" For Append Shared As #N
    Print #N, Date & " " & Time & " servidor reiniciado."
    Close #N
    
    'Ocultar
    
    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)

    End If

    '<EhFooter>
    Exit Sub

Restart_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.Restart " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function Intemperie(ByVal UserIndex As Integer) As Boolean

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 15/11/2009
    '15/11/2009: ZaMa - La lluvia no quita stamina en las arenas.
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '**************************************************************
    '<EhHeader>
    On Error GoTo Intemperie_Err

    '</EhHeader>

    With UserList(UserIndex)

        If MapInfo(.Pos.Map).Zona <> "DUNGEON" Then
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> 1 And MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> 2 And MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> 4 Then Intemperie = True
        Else
            Intemperie = False

        End If

    End With
    
    'En las arenas no te afecta la lluvia
    If IsArena(UserIndex) Then Intemperie = False
    '<EhFooter>
    Exit Function

Intemperie_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.Intemperie " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub TiempoInvocacion(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo TiempoInvocacion_Err

    '</EhHeader>

    With UserList(UserIndex)

        If .MascotaIndex > 0 Then
            If Npclist(.MascotaIndex).Contadores.TiempoExistencia > 0 Then
                Npclist(.MascotaIndex).Contadores.TiempoExistencia = Npclist(.MascotaIndex).Contadores.TiempoExistencia - 1

                If Npclist(.MascotaIndex).Contadores.TiempoExistencia = 0 Then Call MuereNpc(.MascotaIndex, 0)

            End If

        End If

    End With

    '<EhFooter>
    Exit Sub

TiempoInvocacion_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.TiempoInvocacion " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub EfectoTransformacion(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo EfectoTransformacion_Err

    '</EhHeader>
    
    With UserList(UserIndex)
        .Stats.MinSta = .Stats.MinSta - 15

        If .Stats.MinSta <= 0 Then .Stats.MinSta = 0
        
        Call WriteUpdateSta(UserIndex)
        
        If .Stats.MinSta <= 0 Then
            Call Transform_User(UserIndex, 0)
            Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    '<EhFooter>
    Exit Sub

EfectoTransformacion_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.EfectoTransformacion " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Function CheckTunicaPolar(ByVal UserIndex As Integer) As Boolean

    With UserList(UserIndex)

        If .Invent.ArmourEqpObjIndex = 0 Then Exit Function
        
        If ObjData(.Invent.ArmourEqpObjIndex).AntiFrio > 0 Then
            CheckTunicaPolar = True 'ObjData(.Invent.ArmourEqpObjIndex).AntiFrio

        End If
    
    End With
    
End Function

Public Sub EfectoFrio(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo EfectoFrio_Err

    '</EhHeader>

    '***************************************************
    'Autor: Unkonwn
    'Last Modification: 23/11/2009
    'If user is naked and it's in a cold map, take health points from him
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************
    Dim modifi     As Integer

    Dim EfectoFrio As Integer
        
    With UserList(UserIndex)
        
        If .Counters.Frio < IntervaloFrio Then
            .Counters.Frio = .Counters.Frio + 1
        Else

            If MapInfo(.Pos.Map).Terreno = eTerrain.terrain_nieve Then
                EfectoFrio = CheckTunicaPolar(UserIndex)
                     
                If .Invent.ArmourEqpObjIndex > 0 Then
                    EfectoFrio = CheckTunicaPolar(UserIndex)

                End If
                    
                If Not CheckTunicaPolar(UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muriendo de frío, abrigate o morirás!!", FontTypeNames.FONTTYPE_INFO)
                    modifi = Porcentaje(.Stats.MaxHp, 10)
                    .Stats.MinHp = .Stats.MinHp - modifi
                    
                    If .Stats.MinHp < 1 Then
                        Call WriteConsoleMsg(UserIndex, "¡¡Has muerto de frío!!", FontTypeNames.FONTTYPE_INFO)
                        .Stats.MinHp = 0
                        Call UserDie(UserIndex)
                    Else
                        Call WriteUpdateHP(UserIndex)

                    End If

                End If

            ElseIf MapInfo(.Pos.Map).Terreno = eTerrain.terrain_bosque Then

                If .flags.Desnudo = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muriendo de frío, abrigate o morirás!!", FontTypeNames.FONTTYPE_INFO)
                    modifi = Porcentaje(.Stats.MaxHp, 5)
                    .Stats.MinHp = .Stats.MinHp - modifi
                    
                    If .Stats.MinHp < 1 Then
                        Call WriteConsoleMsg(UserIndex, "¡¡Has muerto de frío!!", FontTypeNames.FONTTYPE_INFO)
                        .Stats.MinHp = 0
                        Call UserDie(UserIndex)
                    Else
                        Call WriteUpdateHP(UserIndex)

                    End If
                
                End If

            End If
            
            .Counters.Frio = 0

        End If

    End With

    '<EhFooter>
    Exit Sub

EfectoFrio_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.EfectoFrio " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub EfectoLava(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo EfectoLava_Err

    '</EhHeader>

    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 23/11/2009
    'If user is standing on lava, take health points from him
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************
    With UserList(UserIndex)

        If .Counters.Lava < IntervaloFrio Then 'Usamos el mismo intervalo que el del frio
            .Counters.Lava = .Counters.Lava + 1
        Else

            If HayLava(.Pos.Map, .Pos.X, .Pos.Y) Then
                Call WriteConsoleMsg(UserIndex, "¡¡Quitate de la lava, te estás quemando!!", FontTypeNames.FONTTYPE_INFO)
                .Stats.MinHp = .Stats.MinHp - Porcentaje(.Stats.MaxHp, 5)
                
                If .Stats.MinHp < 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Has muerto quemado!!", FontTypeNames.FONTTYPE_INFO)
                    .Stats.MinHp = 0
                    Call UserDie(UserIndex)

                End If
                
                Call WriteUpdateHP(UserIndex)

            End If
            
            .Counters.Lava = 0

        End If

    End With

    '<EhFooter>
    Exit Sub

EfectoLava_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.EfectoLava " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub EfectoAceleracion(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        .Counters.BuffoAceleration = .Counters.BuffoAceleration - 1
        
        If .Counters.BuffoAceleration <= 0 Then
            Call ActualizarVelocidadDeUsuario(UserIndex, False)

        End If
    
    End With

End Sub

''
' Maneja el tiempo de arrivo al hogar
'
' @param UserIndex  El index del usuario a ser afectado por el

'

Public Sub TravelingEffect(ByVal UserIndex As Integer)

    '******************************************************
    'Author: ZaMa
    'Last Update: 01/06/2010 (ZaMa)
    '******************************************************
    '<EhHeader>
    On Error GoTo TravelingEffect_Err

    '</EhHeader>
    
    Dim TiempoTranscurrido As Long

    ' Si ya paso el tiempo de penalizacion
    If IntervaloGoHome(UserIndex) Then
        Call HomeArrival(UserIndex)

    End If

    '<EhFooter>
    Exit Sub

TravelingEffect_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.TravelingEffect " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Maneja el tiempo y el efecto del mimetismo
'
' @param UserIndex  El index del usuario a ser afectado por el mimetismo
'

Public Sub EfectoMimetismo(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo EfectoMimetismo_Err

    '</EhHeader>

    '******************************************************
    'Author: Unknown
    'Last Update: 16/09/2010 (ZaMa)
    '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
    '16/09/2010: ZaMa - Se recupera la apariencia de la barca correspondiente despues de terminado el mimetismo.
    '******************************************************
    Dim Barco As ObjData
    
    With UserList(UserIndex)

        If .flags.Transform > 0 Then Exit Sub
        If .flags.SlotEvent > 0 Then Exit Sub
        If .flags.TransformVIP > 0 Then Exit Sub
        
        If .Counters.Mimetismo < IntervaloInvisible Then
            .Counters.Mimetismo = .Counters.Mimetismo + 1
        Else
            'restore old char
            Call WriteConsoleMsg(UserIndex, "Recuperas tu apariencia normal.", FontTypeNames.FONTTYPE_INFO)
            
            Call Mimetismo_Reset(UserIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

EfectoMimetismo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.EfectoMimetismo " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Mimetismo_Reset(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo Mimetismo_Reset_Err

    '</EhHeader>

    Dim A As Long
        
    With UserList(UserIndex)

        If .flags.Navegando Then
            If .flags.Muerto = 0 Then
                Call ToggleBoatBody(UserIndex)
            Else
                .Char.Body = iFragataFantasmal
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco

                For A = 1 To MAX_AURAS
                    .Char.AuraIndex(A) = NingunAura
                Next A

            End If

        Else
            .Char.Body = .CharMimetizado.Body
            .Char.Head = .CharMimetizado.Head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                  
            For A = 1 To MAX_AURAS
                .Char.AuraIndex(A) = .CharMimetizado.AuraIndex
            Next A
                  
        End If
            
        With .Char
            Call ChangeUserChar(UserIndex, .Body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraIndex)

        End With
            
        If .ShowName = False Then
            .ShowName = True
            Call RefreshCharStatus(UserIndex)

        End If
            
        .Counters.Mimetismo = 0
        .flags.Mimetizado = 0
        ' Se fue el efecto del mimetismo, puede ser atacado por npcs
        .flags.Ignorado = False
    
    End With

    '<EhFooter>
    Exit Sub

Mimetismo_Reset_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.Mimetismo_Reset " & "at line " & Erl
        
    '</EhFooter>
End Sub

'
Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/09/2010 (ZaMa)
    '16/09/2010: ZaMa - Al perder el invi cuando navegas, no se manda el mensaje de sacar invi (ya estas visible).
    '***************************************************
    '<EhHeader>
    On Error GoTo EfectoInvisibilidad_Err

    '</EhHeader>
    Dim TiempoTranscurrido As Long

    With UserList(UserIndex)

        If .Counters.Invisibilidad < IntervaloInvisible Then
            .Counters.Invisibilidad = .Counters.Invisibilidad + 1
            
            TiempoTranscurrido = (.Counters.Invisibilidad * frmMain.GameTimer.interval)
            
            If TiempoTranscurrido Mod 1000 = 0 Or TiempoTranscurrido = 40 Then
                If TiempoTranscurrido = 40 Then
                    Call WriteUpdateGlobalCounter(UserIndex, 1, ((IntervaloInvisible * 40) / 1000))
                Else
                    Call WriteUpdateGlobalCounter(UserIndex, 1, ((IntervaloInvisible * 40) / 1000) - ((.Counters.Invisibilidad * 40) / 1000))

                End If

            End If
            
            If .flags.Navegando = 0 Then
                Call EfectoInvisibilidad_Drawers(UserIndex)

            End If

        Else
            .Counters.Invisibilidad = 0
            .flags.Invisible = 0
            
            Call WriteUpdateGlobalCounter(UserIndex, 1, 0)
            
            If .flags.Oculto = 0 Then
                Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                
                ' Si navega ya esta visible..
                If Not .flags.Navegando = 1 Then

                    'Si está en un oscuro no lo hacemos visible
                    If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> eTrigger.zonaOscura Then
                        Call SetInvisible(UserIndex, .Char.charindex, False)

                    End If

                End If
                
            End If

        End If

    End With

    '<EhFooter>
    Exit Sub

EfectoInvisibilidad_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.EfectoInvisibilidad " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub EfectoInvisibilidad_Drawers(ByVal UserIndex As Integer)

    ' Author: Lautaro Marino
    ' Este procedimiento se encarga de mandar un paquete al cliente para visualizar los clientes invisibles durante un segundo.
    '<EhHeader>
    On Error GoTo EfectoInvisibilidad_Drawers_Err

    '</EhHeader>
    
    With UserList(UserIndex)

        If .Counters.DrawersCount > 0 Then
            .Counters.DrawersCount = .Counters.DrawersCount - 1
            
            If .Counters.DrawersCount = 0 Then
                .Counters.Drawers = RandomNumberPower(7, 15)
                Call SetInvisible(UserIndex, .Char.charindex, .flags.Invisible, True)

                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, .flags.Invisible, True))
            End If
            
        End If
        
        If .Counters.Drawers > 0 Then
            .Counters.Drawers = .Counters.Drawers - 1
        
            If .Counters.Drawers = 0 Then
                .Counters.DrawersCount = RandomNumberPower(1, 200)
                Call SetInvisible(UserIndex, .Char.charindex, .flags.Invisible, False)

                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, .flags.Invisible, False))
            End If

        End If
        
    End With

    '<EhFooter>
    Exit Sub

EfectoInvisibilidad_Drawers_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.EfectoInvisibilidad_Drawers " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo EfectoParalisisNpc_Err

    '</EhHeader>

    With Npclist(NpcIndex)

        If .Contadores.Paralisis > 0 Then
            .Contadores.Paralisis = .Contadores.Paralisis - 1
        Else
            .flags.Paralizado = 0
            .flags.Inmovilizado = 0

        End If

    End With

    '<EhFooter>
    Exit Sub

EfectoParalisisNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.EfectoParalisisNpc " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo EfectoCegueEstu_Err

    '</EhHeader>

    With UserList(UserIndex)

        If .Counters.Ceguera > 0 Then
            .Counters.Ceguera = .Counters.Ceguera - 1
        Else

            If .flags.Ceguera = 1 Then
                .flags.Ceguera = 0
                Call WriteBlindNoMore(UserIndex)

            End If

            If .flags.Estupidez = 1 Then
                .flags.Estupidez = 0
                Call WriteDumbNoMore(UserIndex)

            End If
        
        End If

    End With

    '<EhFooter>
    Exit Sub

EfectoCegueEstu_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.EfectoCegueEstu " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 02/12/2010
    '02/12/2010: ZaMa - Now non-magic clases lose paralisis effect under certain circunstances.
    '***************************************************
    '<EhHeader>
    On Error GoTo EfectoParalisisUser_Err

    '</EhHeader>

    Dim TiempoTranscurrido As Long
    
    With UserList(UserIndex)
    
        If .Counters.Paralisis > 0 Then
        
            Dim CasterIndex As Integer

            CasterIndex = .flags.ParalizedByIndex
        
            ' Only aplies to non-magic clases
            If .Stats.MaxMan = 0 Then

                ' Paralized by user?
                If CasterIndex <> 0 Then
                
                    ' Close? => Remove Paralisis
                    If UserList(CasterIndex).Name <> .flags.ParalizedBy Then
                        Call RemoveParalisis(UserIndex)

                        Exit Sub
                        
                        ' Caster dead? => Remove Paralisis
                    ElseIf UserList(CasterIndex).flags.Muerto = 1 Then
                        Call RemoveParalisis(UserIndex)

                        Exit Sub
                    
                    ElseIf .Counters.Paralisis > IntervaloParalizadoReducido Then

                        ' Out of vision range? => Reduce paralisis counter
                        If Not InVisionRangeAndMap(UserIndex, UserList(CasterIndex).Pos) Then
                            ' Aprox. 1500 ms
                            .Counters.Paralisis = IntervaloParalizadoReducido

                            Exit Sub

                        End If

                    End If
                
                    ' Npc?
                Else
                    CasterIndex = .flags.ParalizedByNpcIndex
                    
                    ' Paralized by npc?
                    If CasterIndex <> 0 Then
                    
                        If .Counters.Paralisis > IntervaloParalizadoReducido Then

                            ' Out of vision range? => Reduce paralisis counter
                            If Not InVisionRangeAndMap(UserIndex, Npclist(CasterIndex).Pos) Then
                                ' Aprox. 1500 ms
                                .Counters.Paralisis = IntervaloParalizadoReducido

                                Exit Sub

                            End If

                        End If

                    End If
                    
                End If

            End If
            
            .Counters.Paralisis = .Counters.Paralisis - 1
            
            TiempoTranscurrido = (.Counters.Paralisis * frmMain.GameTimer.interval)
            
            If TiempoTranscurrido Mod 1000 = 0 Or TiempoTranscurrido = 40 Then
                Call WriteUpdateGlobalCounter(UserIndex, 2, ((.Counters.Paralisis * 40) / 1000))

            End If

        Else
            Call RemoveParalisis(UserIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

EfectoParalisisUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.EfectoParalisisUser " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub RemoveParalisis(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo RemoveParalisis_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    'Removes paralisis effect from user.
    '***************************************************
    With UserList(UserIndex)
        .flags.Paralizado = 0
        .flags.Inmovilizado = 0
        .flags.ParalizedBy = vbNullString
        .flags.ParalizedByIndex = 0
        .flags.ParalizedByNpcIndex = 0
        .Counters.Paralisis = 0
        Call WriteParalizeOK(UserIndex)
        
        WriteUpdateGlobalCounter UserIndex, 2, 0

    End With

    '<EhFooter>
    Exit Sub

RemoveParalisis_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.RemoveParalisis " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, _
                      ByRef EnviarStats As Boolean, _
                      ByVal Intervalo As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo RecStamina_Err

    '</EhHeader>

    With UserList(UserIndex)

        If .Pos.Map = 0 Then Exit Sub
        If .flags.Transform > 0 Then Exit Sub

        Dim massta As Integer
        
        If .flags.Desnudo Then
            If .Stats.MinSta > 0 Then
                If .Counters.STACounter < Intervalo Then
                    .Counters.STACounter = .Counters.STACounter + 1
                Else
                    EnviarStats = True
                    .Counters.STACounter = 0

                    massta = RandomNumber(1, Porcentaje(.Stats.MaxSta, 5))
                    .Stats.MinSta = .Stats.MinSta - massta
                    
                    If .Stats.MinSta <= 0 Then
                        .Stats.MinSta = 0

                    End If

                End If

            End If

        Else
        
            If .Stats.MinSta < .Stats.MaxSta Then
                If .Counters.STACounter < Intervalo Then
                    .Counters.STACounter = .Counters.STACounter + 1
                Else
                    EnviarStats = True
                    .Counters.STACounter = 0
                    'If .flags.Desnudo Then Exit Sub 'Desnudo no sube energía. (ToxicWaste)
                   
                    massta = RandomNumber(1, Porcentaje(.Stats.MaxSta, 10))
                    .Stats.MinSta = .Stats.MinSta + massta

                    If .Stats.MinSta > .Stats.MaxSta Then
                        .Stats.MinSta = .Stats.MaxSta

                    End If

                End If

            End If

        End If

    End With
    
    '<EhFooter>
    Exit Sub

RecStamina_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.RecStamina " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub User_EfectoIncineracion(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo User_EfectoIncineracion_Err

    '</EhHeader>
    Dim N As Integer
    
    With UserList(UserIndex)

        If .Counters.Incinerado < IntervaloFrio Then
            .Counters.Incinerado = .Counters.Incinerado + 1
        Else
            .Counters.Incinerado = 0
            
            Call WriteConsoleMsg(UserIndex, "¡Te estas incinerando, si no te curas morirás!", FontTypeNames.FONTTYPE_VENENO)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(eSound.sFogata, .Pos.X, .Pos.Y, .Char.charindex, True))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, FXIDs.FX_INCINERADO, -1))
            
            N = RandomNumber(1, 50)
            .Stats.MinHp = .Stats.MinHp - N

            If .Stats.MinHp < 1 Then Call UserDie(UserIndex)
            Call WriteUpdateHP(UserIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

User_EfectoIncineracion_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.User_EfectoIncineracion " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Npc_EfectoIncineracion(ByVal NpcIndex As Integer)

    '<EhHeader>
    On Error GoTo Npc_EfectoIncineracion_Err

    '</EhHeader>
    Dim N         As Integer

    Dim UserIndex As Integer
    
    With Npclist(NpcIndex)

        If .Contadores.Incinerado < IntervaloFrio Then
            .Contadores.Incinerado = .Contadores.Incinerado + 1
        Else
            .Contadores.Incinerado = 0
            
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(eSound.sFogata, .Pos.X, .Pos.Y, .Char.charindex, True))
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFX(.Char.charindex, FXIDs.FX_INCINERADO, -1))
            
            N = RandomNumber(1, 50)
            .Stats.MinHp = .Stats.MinHp - N
        
            UserIndex = .Owner

            If .Stats.MinHp < 1 Then Call MuereNpc(NpcIndex, UserIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

Npc_EfectoIncineracion_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.Npc_EfectoIncineracion " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub EfectoVeneno(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo EfectoVeneno_Err

    '</EhHeader>

    Dim N As Integer
    
    With UserList(UserIndex)

        If .Counters.Veneno < IntervaloVeneno Then
            .Counters.Veneno = .Counters.Veneno + 1
        Else
            Call WriteConsoleMsg(UserIndex, "Estás envenenado, si no te curas morirás.", FontTypeNames.FONTTYPE_VENENO)
            .Counters.Veneno = 0
            N = RandomNumber(1, 5)
            .Stats.MinHp = .Stats.MinHp - N

            If .Stats.MinHp < 1 Then Call UserDie(UserIndex)
            Call WriteUpdateHP(UserIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

EfectoVeneno_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.EfectoVeneno " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub DuracionPociones(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ??????
    'Last Modification: 08/06/11 (CHOTS)
    'Le agregué que avise antes cuando se te está por ir
    '
    'Cuando se pierde el efecto de la poción updatea fz y agi (No me gusta que ambos atributos aunque se haya modificado solo uno, pero bueno :p)
    '***************************************************
    '<EhHeader>
    On Error GoTo DuracionPociones_Err

    '</EhHeader>

    Const SEGUNDOS_AVISO   As Byte = 5
        
    Dim Tick               As Long

    Dim TiempoTranscurrido As Long
        
    Tick = GetTime
    'CHOTS | Los segundos antes que se te acabe que te avisa

    With UserList(UserIndex)

        'Controla la duracion de las pociones
        If .flags.DuracionEfecto > 0 Then
            .flags.DuracionEfecto = .flags.DuracionEfecto - 1
                    
            TiempoTranscurrido = (.flags.DuracionEfecto * frmMain.GameTimer.interval)
            
            If TiempoTranscurrido Mod 1000 = 0 Or TiempoTranscurrido = 40 Then
                Call WriteUpdateGlobalCounter(UserIndex, 3, .flags.DuracionEfecto / 40)

            End If
            
            If ((.flags.DuracionEfecto / 25) <= SEGUNDOS_AVISO) Then    'CHOTS | Lo divide por 25 por el intervalo del Timer (40x25=1000=1seg)
                If Tick - .Counters.RuidoDopa > 5000 Then
                    .Counters.RuidoDopa = Tick
                    Call WriteStrDextRunningOut(UserIndex)

                End If

                '  .flags.UltimoMensaje = 221
            End If

            If .flags.DuracionEfecto = 0 Then
                ' .flags.UltimoMensaje = 222
                .flags.TomoPocion = False
                .flags.TipoPocion = 0

                'volvemos los atributos al estado normal
                Dim LoopX As Integer
                
                For LoopX = 1 To NUMATRIBUTOS
                    .Stats.UserAtributos(LoopX) = .Stats.UserAtributosBackUP(LoopX)
                Next LoopX
                
                Call WriteUpdateStrenghtAndDexterity(UserIndex)

            End If

        End If
        
        Dim UpdateMAN As Boolean

        Dim UpdateHP  As Boolean, TempTick As Long
        
        ' Pociones Azules (Clic)
        If .PotionBlue_Clic > 0 Then
            .PotionBlue_Clic_Interval = .PotionBlue_Clic_Interval + 1
            
            If .PotionBlue_Clic_Interval >= TOLERANCE_POTIONBLUE_CLIC Then
                .PotionBlue_Clic = .PotionBlue_Clic - 1
                
                If .PotionBlue_Clic > 0 Then
                    .Stats.MinMan = .Stats.MinMan + Porcentaje(.Stats.MaxMan, 3) + .Stats.Elv \ 2 + 40 / .Stats.Elv
                                
                    If .Stats.MinMan > .Stats.MaxMan Then .Stats.MinMan = .Stats.MaxMan
                    UpdateMAN = True

                End If
                
                .PotionBlue_Clic_Interval = 0
                    
            End If
        
            .PotionRed_Clic = 0
            .PotionRed_U = 0
            .PotionRed_U_Interval = 0
            .PotionRed_Clic_Interval = 0

        End If
        
        ' Pociones Azules (U)
        If .PotionBlue_U > 0 Then
            .PotionBlue_U_Interval = .PotionBlue_U_Interval + 1
            
            If .PotionBlue_U_Interval >= TOLERANCE_POTIONBLUE_U Then
                .PotionBlue_U = .PotionBlue_U - 1
                
                If .PotionBlue_U > 0 Then
                    .Stats.MinMan = .Stats.MinMan + Porcentaje(.Stats.MaxMan, 3) + .Stats.Elv \ 2 + 40 / .Stats.Elv
                                
                    If .Stats.MinMan > .Stats.MaxMan Then .Stats.MinMan = .Stats.MaxMan
                    
                    UpdateMAN = True

                End If
                
                .PotionBlue_U_Interval = 0

            End If
            
            .PotionRed_Clic = 0
            .PotionRed_U = 0
            .PotionRed_U_Interval = 0
            .PotionRed_Clic_Interval = 0

        End If
        
        If UpdateMAN Then Call WriteUpdateMana(UserIndex)
        
        ' Pociones Rojas (Clic)
        If .PotionRed_Clic > 0 Then
            .PotionRed_Clic_Interval = .PotionRed_Clic_Interval + 1
            
            If .PotionRed_Clic_Interval >= TOLERANCE_POTIONRED_CLIC Then
                .PotionRed_Clic = .PotionRed_Clic - 1
                
                If .PotionRed_Clic > 0 Then
                    .Stats.MinHp = .Stats.MinHp + ObjData(38).MaxModificador
    
                    If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
                    UpdateHP = True

                End If
                
                .PotionRed_Clic_Interval = 0

            End If
            
            .PotionBlue_Clic = 0
            .PotionBlue_U = 0
            .PotionBlue_U_Interval = 0
            .PotionBlue_Clic_Interval = 0
        
        End If
        
        ' Pociones Rojas (U)
        If .PotionRed_U > 0 Then
            .PotionRed_U_Interval = .PotionRed_U_Interval + 1
            
            If .PotionRed_U_Interval >= TOLERANCE_POTIONRED_U Then
                .PotionRed_U = .PotionRed_U - 1
                
                If .PotionRed_U > 0 Then
                    .Stats.MinHp = .Stats.MinHp + ObjData(38).MaxModificador
        
                    If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
                    
                    UpdateHP = True

                End If
                
                .PotionRed_U_Interval = 0
                
                .PotionBlue_Clic = 0
                .PotionBlue_U = 0
                .PotionBlue_U_Interval = 0
                .PotionBlue_Clic_Interval = 0

            End If

        End If
        
        If UpdateHP Then
            Call WriteUpdateHP(UserIndex)
            
            If TempTick - .Counters.RuidoPocion > 1000 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                .Counters.RuidoPocion = TempTick

            End If

        End If

    End With

    '<EhFooter>
    Exit Sub

DuracionPociones_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.DuracionPociones " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub HambreYSed(ByVal UserIndex As Integer, ByRef fenviarAyS As Boolean)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo HambreYSed_Err

    '</EhHeader>
  
    With UserList(UserIndex)

        If Not .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Dim cant As Byte
        
        If .Stats.Elv <= 18 Then
            cant = 1
        Else
            cant = 10

        End If
        
        'Sed
        If .Stats.MinAGU > 0 Then
            If .Counters.AGUACounter < IntervaloSed Then
                .Counters.AGUACounter = .Counters.AGUACounter + 1
            Else
                .Counters.AGUACounter = 0
                .Stats.MinAGU = .Stats.MinAGU - cant
                
                If .Stats.MinAGU <= 0 Then
                    .Stats.MinAGU = 0
                    .flags.Sed = 1

                End If
                
                fenviarAyS = True

            End If

        End If
        
        'hambre
        If .Stats.MinHam > 0 Then
            If .Counters.COMCounter < IntervaloHambre Then
                .Counters.COMCounter = .Counters.COMCounter + 1
            Else
                .Counters.COMCounter = 0
                .Stats.MinHam = .Stats.MinHam - cant

                If .Stats.MinHam <= 0 Then
                    .Stats.MinHam = 0
                    .flags.Hambre = 1

                End If

                fenviarAyS = True

            End If

        End If

    End With

    '<EhFooter>
    Exit Sub

HambreYSed_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.HambreYSed " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Sanar(ByVal UserIndex As Integer, _
                 ByRef EnviarStats As Boolean, _
                 ByVal Intervalo As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo Sanar_Err

    '</EhHeader>

    With UserList(UserIndex)

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 1 And MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 2 And MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 4 Then Exit Sub
        
        Dim mashit As Integer

        'con el paso del tiempo va sanando....pero muy lentamente ;-)
        If .Stats.MinHp < .Stats.MaxHp Then
            If .Counters.HPCounter < Intervalo Then
                .Counters.HPCounter = .Counters.HPCounter + 1
            Else
                mashit = RandomNumber(2, Porcentaje(.Stats.MaxSta, 5))
                
                .Counters.HPCounter = 0
                .Stats.MinHp = .Stats.MinHp + mashit

                If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
                Call WriteConsoleMsg(UserIndex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
                EnviarStats = True

            End If

        End If

    End With

    '<EhFooter>
    Exit Sub

Sanar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.Sanar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub CargaNpcsDat()

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo CargaNpcsDat_Err

    '</EhHeader>

    Dim npcfile As String
    
    npcfile = Npcs_FilePath
    Set LeerNPCs = New clsIniManager
    Call LeerNPCs.Initialize(npcfile)

    Call DataServer_Generate_Npcs
    '<EhFooter>
    Exit Sub

CargaNpcsDat_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.CargaNpcsDat " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub CheckCountDown()

    '<EhHeader>
    On Error GoTo CheckCountDown_Err

    '</EhHeader>

    If CountDown_Time = 0 Then Exit Sub
    
    CountDown_Time = CountDown_Time - 1
            
    If CountDown_Map > 0 Then
        Call SendData(SendTarget.toMap, CountDown_Map, PrepareMessageRender_CountDown(CountDown_Time))

        If CountDown_Time = 0 Then
            Call SendData(SendTarget.toMap, CountDown_Map, PrepareMessageConsoleMsg("¡YA!", FontTypeNames.FONTTYPE_FIGHT))

        End If
        
    Else

        Call SendData(SendTarget.toMapSecure, 0, PrepareMessageRender_CountDown(CountDown_Time))
        
        If CountDown_Time = 0 Then
            Call SendData(SendTarget.toMapSecure, 0, PrepareMessageConsoleMsg("¡YA!", FontTypeNames.FONTTYPE_FIGHT))

        End If

    End If

    '<EhFooter>
    Exit Sub

CheckCountDown_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.CheckCountDown " & "at line " & Erl
        
    '</EhFooter>
End Sub

' # Chequea si la zona está caliente
Public Sub Map_CheckFire(ByVal Map As Integer)
    
    On Error GoTo ErrHandler
    
    Const MIN_USER_FIRE As Integer = 6
    
    With MapInfo(Map)

        If .Pk = False Then Exit Sub
        If .NumUsers < MIN_USER_FIRE Then Exit Sub
        
        .OnFire = .OnFire + 1
        
        If .OnFire >= 5 Then ' 3 minutos superando los 6 usuarios
            If FrasesLastMap = Map Then
                .OnFire = 0
                Exit Sub

            End If
            
            ' Selecciona una frase aleatoria
            Dim randomIndex As Integer

            randomIndex = Int((UBound(FrasesOnFire) + 1) * Rnd)

            Dim Mensaje As String

            Mensaje = Replace(FrasesOnFire(randomIndex), "{Mapa}", "**" & .Name & "**")
        
            ' # Envia mensaje a DISCORD de la concentración
            WriteMessageDiscord CHANNEL_ONFIRE, Mensaje & " " & "Players: " & .NumUsers
            
            .OnFire = 0
            FrasesLastMap = Map

        End If

    End With
    
    Exit Sub
ErrHandler:
    Call LogError("Error en checkfire")

End Sub

Sub PasarSegundo()

    '<EhHeader>
    On Error GoTo PasarSegundo_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Long
        
    ' Apertura del servidor
    Call User_Go_Initial_Version
        
    Call Teleports_Loop
        
    Call Invations_MainLoop
    
    ' Respawn de Objetos
    Call ChestLoop

    Call Pretorians_Loop
        
    ' Subasta de Objetos
    Call Auction_Loop
    
    ' Cuenta regresiva
    Call CheckCountDown
    
    ' Sistema de auto baneo
    Call AutoBan_Loop
    
    ' Retos Loop
    Call Retos_Loop
    
    ' Respawn de Npcs
    Call Loop_RespawnNpc
    
    If CountDownLimpieza > 0 Then
        CountDownLimpieza = CountDownLimpieza - 1
        
        If CountDownLimpieza < 4 And CountDownLimpieza > 0 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Limpiando el mundo en " & CountDownLimpieza & " segundo" & IIf((CountDownLimpieza > 1), "s...", "..."), FontTypeNames.FONTTYPE_INFOGREEN))

        End If
        
        If CountDownLimpieza <= 0 Then
            Call LimpiarMundo

        End If

    End If
    
    For i = 1 To LastUser

        With UserList(i)

            If i <> SLOT_TERMINAL_ARCHIVE Then
                UserList(i).Counters.TimeInactive = UserList(i).Counters.TimeInactive + 1
                
                If UserList(i).Counters.TimeInactive >= 60 Then
                    UserList(i).Counters.TimeInactive = 0
                    Call WriteDisconnect(i, True)
                    Call Protocol.Kick(i)

                End If

            End If
        
        End With
                
        If UserList(i).flags.UserLogged Then

            With UserList(i)
                    
                If .Stats.BonusLast > 0 Then Call Reward_Check_User(i)  ' # Bonus del personaje
                    
                If .Counters.goHomeSec > 0 Then
                     
                    .Counters.goHomeSec = .Counters.goHomeSec - 1
                        
                    Call WriteUpdateGlobalCounter(i, 4, .Counters.goHomeSec)

                End If

                If .Counters.ShieldBlocked > 0 Then
                    .Counters.ShieldBlocked = .Counters.ShieldBlocked - 1

                End If
                
                If .Counters.Shield > 0 Then
                    .Counters.Shield = .Counters.Shield - 1
                    
                    If .Counters.Shield = 0 Then
                        Call RefreshCharStatus(i)

                    End If

                End If
                
                If .Counters.FightSend > 0 Then
                    .Counters.FightSend = .Counters.FightSend - 1
                    
                    If .Counters.FightSend = 0 Then
                        Call WriteConsoleMsg(i, "Ya puedes enviar otra invitación de reto.", FontTypeNames.FONTTYPE_INFOGREEN)

                    End If

                End If
                
                If .Counters.ReviveAutomatic > 0 Then
                    .Counters.ReviveAutomatic = .Counters.ReviveAutomatic - 1
                    
                    ' // NUEVO
                    If .Counters.ReviveAutomatic > 0 And .Counters.ReviveAutomatic <= 5 Then
                        Call WriteConsoleMsg(i, "Serás revivido en " & .Counters.ReviveAutomatic & " segundo" & IIf((.Counters.ReviveAutomatic = 1), "s.", "."), FontTypeNames.FONTTYPE_INFO)

                    End If
                    
                    If .Counters.ReviveAutomatic = 0 Then
                        If .flags.Muerto Then Call RevivirUsuario(i)

                    End If

                End If
                
                If .Counters.FightInvitation > 0 Then
                    .Counters.FightInvitation = .Counters.FightInvitation - 1

                End If
                
                If .Counters.TimePublicationMao > 0 Then
                    .Counters.TimePublicationMao = .Counters.TimePublicationMao - 1

                End If
                
                If .Counters.TimeCreateChar > 0 Then .Counters.TimeCreateChar = .Counters.TimeCreateChar - 1
                
                If .Counters.TimeDenounce > 0 Then
                    .Counters.TimeDenounce = .Counters.TimeDenounce - 1

                End If
                
                Call Effect_Loop(i)
                Call AntiFrags_CheckTime(i)
                
                If .flags.Transform Then EfectoTransformacion (i)

                If .Counters.TimeFight > 0 Then
                    .Counters.TimeFight = .Counters.TimeFight - 1
                    
                    ' Cuenta regresiva de retos y eventos
                    If .Counters.TimeFight = 0 Then
                        Call WriteRender_CountDown(i, .Counters.TimeFight)
                        'WriteConsoleMsg i, "Cuenta» ¡YA!", FontTypeNames.FONTTYPE_FIGHT
                                      
                        ' En los duelos desparalizamos el cliente
                        If .flags.SlotEvent > 0 Then
                            If Events(.flags.SlotEvent).Modality = eModalityEvent.Enfrentamientos Then
                                Call WriteUserInEvent(i)

                            End If

                        End If
                                      
                        If .flags.SlotReto > 0 Then
                            Call WriteUserInEvent(i)

                        End If

                    Else
                        Call WriteRender_CountDown(i, .Counters.TimeFight)

                        'WriteConsoleMsg i, .Counters.TimeFight, FontTypeNames.FONTTYPE_GUILD
                    End If

                End If

                If .Counters.TimeTransform > 0 Then
                    .Counters.TimeTransform = .Counters.TimeTransform - 1

                End If
                
                If .Counters.TimeGlobal > 0 Then
                    .Counters.TimeGlobal = .Counters.TimeGlobal - 1

                End If
        
                If .Counters.TimeBono > 0 Then
                    .Counters.TimeBono = .Counters.TimeBono - 1
                    
                End If
            
                ' Tiempo para que el usuario se vaya del mapa
                If .Counters.TimeTelep > 0 Then

                    ' Efecto con objeto
                    If .flags.ObjIndex Then
                        If ObjData(.flags.ObjIndex).TelepMap = .Pos.Map Then
                            .Counters.TimeTelep = .Counters.TimeTelep - 1
                                    
                            If .Counters.TimeTelep = 0 Then
                                Call QuitarObjetos(.flags.ObjIndex, 1, i)
                                .flags.ObjIndex = 0
                                WarpUserChar i, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, True
                                WriteConsoleMsg i, "El efecto de la teletransportación ha terminado.", FontTypeNames.FONTTYPE_INFO
    
                            End If

                        End If
                        
                    Else
                        .Counters.TimeTelep = .Counters.TimeTelep - 1
                        
                        If .Counters.TimeTelep = 0 Then
                            If .flags.SlotEvent Then
                                Call Events_ChangePosition(i, .flags.SlotEvent)

                            End If

                        End If

                    End If

                End If
                
                ' Tiempo para cambiar de apariencia (cada 3 segundos en caso de evento)
                If .Counters.TimeApparience > 0 Then
                    .Counters.TimeApparience = .Counters.TimeApparience - 1
                    
                    If .Counters.TimeApparience = 0 Then
                        Call Events_ChangeApparience(i)

                    End If

                End If

            End With

        End If
        
        With UserList(i)

            'Cerrar usuario
            If .Counters.Saliendo Then
                .Counters.Salir = .Counters.Salir - 1

                If .Counters.Salir <= 0 Then
                                     
                    If .flags.DeslogeandoCuenta Then
                        .flags.DeslogeandoCuenta = False
                        Call WriteDisconnect(i, True)
                        Call FlushBuffer(i)
                        Call Server.Kick(i, True)
                                           
                    Else
                        Call WriteConsoleMsg(i, "Desconectado personaje...", FontTypeNames.FONTTYPE_INFO)
                        Call WriteDisconnect(i)
                        Call CloseSocket(i)
                        Call FlushBuffer(i)

                    End If

                End If

            End If
        
        End With
      
    Next i
        
    Call Streamer_CheckPosition
        
    '<EhFooter>
    Exit Sub

PasarSegundo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.PasarSegundo " & "at line " & Erl
        
    '</EhFooter>
End Sub
 
Public Function ReiniciarAutoUpdate() As Double

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ReiniciarAutoUpdate_Err

    '</EhHeader>

    ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

    '<EhFooter>
    Exit Function

ReiniciarAutoUpdate_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.ReiniciarAutoUpdate " & "at line " & Erl
        
    '</EhFooter>
End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ReiniciarServidor_Err

    '</EhHeader>

    'commit experiencias
    Call mGroup.DistributeExpAndGldGroups
    
    'WorldSave
    Call ES.DoBackUp

    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    Unload frmMain

    '<EhFooter>
    Exit Sub

ReiniciarServidor_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.ReiniciarServidor " & "at line " & Erl

    '</EhFooter>
End Sub
 
Sub GuardarUsuarios(ByVal IsBackup As Boolean)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo GuardarUsuarios_Err

    '</EhHeader>
    
    Dim i             As Integer

    Dim UserGuardados As Long
        
    For i = 1 To LastUser

        If UserList(i).flags.UserLogged Then
            If Not EsGm(i) Then
                Call Power_Search(i)
                        
            End If
                
            If GetTime - UserList(i).Counters.LastSave > IntervaloGuardarUsuarios Then
                Call UpdatePremium(i)
                    
                ' No guarda personajes en eventos.
                If (UserList(i).flags.SlotEvent = 0 And UserList(i).flags.SlotReto = 0) Then
                    Call SaveUser(UserList(i), CharPath & UCase$(UserList(i).Name) & ".chr", False)

                End If
                    
                Call SaveDataAccount(i, UserList(i).Account.Email, UserList(i).IpAddress)

            End If

        End If

    Next i
    
    '<EhFooter>
    Exit Sub

GuardarUsuarios_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.GuardarUsuarios " & "at line " & Erl

    '</EhFooter>
End Sub

Sub GuardarUsuarios_Close()

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo GuardarUsuarios_Err

    '</EhHeader>
    
    Dim i             As Integer

    Dim UserGuardados As Long
        
    For i = 1 To LastUser

        If UserList(i).flags.UserLogged Then
            Call UpdatePremium(i)
                    
            ' No guarda personajes en eventos.
            If (UserList(i).flags.SlotEvent = 0 And UserList(i).flags.SlotReto = 0) Then
                Call SaveUser(UserList(i), CharPath & UCase$(UserList(i).Name) & ".chr", False)

            End If

        End If
            
        If UserList(i).AccountLogged Then
            Call SaveDataAccount(i, UserList(i).Account.Email, UserList(i).IpAddress)

        End If
            
    Next i
    
    '<EhFooter>
    Exit Sub

GuardarUsuarios_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.GuardarUsuarios " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub FreeNPCs()

    '<EhHeader>
    On Error GoTo FreeNPCs_Err

    '</EhHeader>

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Releases all NPC Indexes
    '***************************************************
    Dim LoopC As Long
    
    ' Free all NPC indexes
    For LoopC = 1 To MAXNPCS
        Npclist(LoopC).flags.NPCActive = False
    Next LoopC

    '<EhFooter>
    Exit Sub

FreeNPCs_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.FreeNPCs " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub FreeCharIndexes()

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Releases all char indexes
    '***************************************************
    ' Free all char indexes (set them all to 0)
    '<EhHeader>
    On Error GoTo FreeCharIndexes_Err

    '</EhHeader>
    Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))
    '<EhFooter>
    Exit Sub

FreeCharIndexes_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.FreeCharIndexes " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function Tilde(Data As String) As String

    '<EhHeader>
    On Error GoTo Tilde_Err

    '</EhHeader>
 
    Tilde = Replace(Replace(Replace(Replace(Replace(UCase$(Data), "Á", "A"), "É", "E"), "Í", "I"), "Ó", "O"), "Ú", "U")
 
    '<EhFooter>
    Exit Function

Tilde_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.Tilde " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub WarpPosAnt(ByVal UserIndex As Integer)
    ' • Warpeo del personaje a su posición anterior.
          
    ' // NUEVO
    Dim Pos As WorldPos
          
    On Error GoTo WarpPosAnt_Error

    With UserList(UserIndex)
        Pos.Map = .PosAnt.Map
        Pos.X = .PosAnt.X
        Pos.Y = .PosAnt.Y
                          
        Call ClosestStablePos(Pos, Pos)
        Call WarpUserChar(UserIndex, Pos.Map, Pos.X, Pos.Y, False)
              
        .PosAnt.Map = 0
        .PosAnt.X = 0
        .PosAnt.Y = 0
          
    End With

    Exit Sub

WarpPosAnt_Error:

    LogError "Error " & Err.number & " (" & Err.description & ") in procedure WarpPosAnt of Módulo General in line " & Erl

End Sub

Public Sub Transform_User(ByVal UserIndex As Integer, ByVal BodySelected As Integer)

    '<EhHeader>
    On Error GoTo Transform_User_Err

    '</EhHeader>
            
    Dim A As Long
            
    With UserList(UserIndex)
        
        If .flags.Transform = 0 Then
            .CharMimetizado.Body = .Char.Body
            .CharMimetizado.Head = .Char.Head
            .CharMimetizado.CascoAnim = .Char.CascoAnim
            .CharMimetizado.ShieldAnim = .Char.ShieldAnim
            .CharMimetizado.WeaponAnim = .Char.WeaponAnim
            
            .Char.Body = BodySelected
            '.Char.Head = 0
            '.Char.CascoAnim = 0
            '.Char.ShieldAnim = 0
            '.Char.WeaponAnim = 0
            .flags.Transform = 1
            '.flags.Ignorado = True
            
        Else

            If .flags.Navegando Then
                If .flags.Muerto = 0 Then
                    Call ToggleBoatBody(UserIndex)
                Else
                    .Char.Body = iFragataFantasmal
                    .Char.ShieldAnim = NingunEscudo
                    .Char.WeaponAnim = NingunArma
                    .Char.CascoAnim = NingunCasco

                    For A = 1 To MAX_AURAS
                        .Char.AuraIndex(A) = NingunAura
                    Next A

                End If

            Else
                .Char.Body = .CharMimetizado.Body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim

                For A = 1 To MAX_AURAS
                    .Char.AuraIndex(A) = .CharMimetizado.AuraIndex
                Next A

            End If
            
            .flags.Transform = 0
            '.flags.Ignorado = False
            
        End If
        
        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
        Call RefreshCharStatus(UserIndex)
        
    End With

    '<EhFooter>
    Exit Sub

Transform_User_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.Transform_User " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub TransformVIP_User(ByVal UserIndex As Integer, ByVal BodySelected As Integer)

    '<EhHeader>
    On Error GoTo TransformVIP_User_Err

    '</EhHeader>
             
    Dim A As Long

    With UserList(UserIndex)

        If .flags.TransformVIP = 0 Then
            .CharMimetizado.Body = .Char.Body
            .CharMimetizado.Head = .Char.Head
            .CharMimetizado.CascoAnim = .Char.CascoAnim
            .CharMimetizado.ShieldAnim = .Char.ShieldAnim
            .CharMimetizado.WeaponAnim = .Char.WeaponAnim
            
            .Char.Body = BodySelected
            .Char.Head = 0
            .Char.CascoAnim = 0
            .Char.ShieldAnim = 0
            .Char.WeaponAnim = 0
            .flags.TransformVIP = 1
            '.flags.Ignorado = True
                
            For A = 1 To MAX_AURAS
                .Char.AuraIndex(A) = NingunAura
            Next A
                          
        Else

            If .flags.Navegando Then
                If .flags.Muerto = 0 Then
                    Call ToggleBoatBody(UserIndex)
                Else
                    .Char.Body = iFragataFantasmal
                    .Char.ShieldAnim = NingunEscudo
                    .Char.WeaponAnim = NingunArma
                    .Char.CascoAnim = NingunCasco

                    For A = 1 To MAX_AURAS
                        .Char.AuraIndex(A) = NingunAura
                    Next A

                End If

            Else
                .Char.Body = .CharMimetizado.Body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim

                For A = 1 To MAX_AURAS
                    .Char.AuraIndex(A) = .CharMimetizado.AuraIndex
                Next A

            End If
            
            For A = 1 To MAX_AURAS
                .Char.AuraIndex(A) = NingunAura
            Next A
                          
            .flags.TransformVIP = 0
            '.flags.Mimetizado = 0
            '.flags.Ignorado = False
            
        End If
        
        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
        Call RefreshCharStatus(UserIndex)
        
    End With

    '<EhFooter>
    Exit Sub

TransformVIP_User_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.TransformVIP_User " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Reinicio al deslogear del UserIndex
Public Sub AntiFrags_ResetInfo(ByRef IUser As User)

    '<EhHeader>
    On Error GoTo AntiFrags_ResetInfo_Err

    '</EhHeader>

    Dim A            As Long

    Dim NullAntiFrag As tAntiFrags
    
    For A = 1 To MAX_CONTROL_FRAGS
        IUser.AntiFrags(A) = NullAntiFrag
    Next A

    '<EhFooter>
    Exit Sub

AntiFrags_ResetInfo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.AntiFrags_ResetInfo " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Buscamos al personaje asesinado en la lista.
Public Function AntiFrags_SlotRepeat(ByVal UserIndex As Integer, _
                                     ByVal VictimIndex As Integer) As Byte

    '<EhHeader>
    On Error GoTo AntiFrags_SlotRepeat_Err

    '</EhHeader>

    Dim A          As Long

    Dim VictimName As String
    
    VictimName = UCase$(UserList(VictimIndex).Name)
    
    For A = 1 To MAX_CONTROL_FRAGS

        With UserList(UserIndex).AntiFrags(A)

            If .UserName = VictimName Then
                AntiFrags_SlotRepeat = A

                Exit Function

            End If

        End With

    Next A

    '<EhFooter>
    Exit Function

AntiFrags_SlotRepeat_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.AntiFrags_SlotRepeat " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function AntiFrags_SlotFree(ByVal UserIndex As Integer) As Byte

    '<EhHeader>
    On Error GoTo AntiFrags_SlotFree_Err

    '</EhHeader>

    Dim A As Long
    
    For A = 1 To MAX_CONTROL_FRAGS

        With UserList(UserIndex).AntiFrags(A)
            
            If .Time <= 0 Then
                AntiFrags_SlotFree = A

                Exit For

            End If

        End With

    Next A

    '<EhFooter>
    Exit Function

AntiFrags_SlotFree_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.AntiFrags_SlotFree " & "at line " & Erl
        
    '</EhFooter>
End Function

' Un personaje es asesinado
Public Function AntiFrags_CheckUser(ByVal UserIndex As Integer, _
                                    ByVal VictimIndex As Integer, _
                                    ByVal Time As Long)

    '<EhHeader>
    On Error GoTo AntiFrags_CheckUser_Err

    '</EhHeader>

    Dim Slot       As Integer

    Dim VictimName As String

    VictimName = UCase$(UserList(VictimIndex).Name)
    
    Slot = AntiFrags_SlotRepeat(UserIndex, VictimIndex)
    
    If Slot <= 0 Then
        Slot = AntiFrags_SlotFree(UserIndex)

    End If
    
    If Slot <= 0 Then GoTo AntiFrags_CheckUser_Err
    
    With UserList(UserIndex).AntiFrags(Slot)
        
        ' El personaje ya está en la lista por lo cual no cuenta el Frag.
        If .UserName = UserList(VictimIndex).Name Then
            AntiFrags_CheckUser = False

            Exit Function

        End If
        
        ' El personaje ya está en la lista por lo cual no cuenta el Frag.
        If .Account = UserList(VictimIndex).Account.Email Then
            AntiFrags_CheckUser = False

            Exit Function

        End If
        
        ' El personaje ya está en la lista por lo cual no cuenta el Frag.
        If .IP <> vbNullString Then
            If .IP = UserList(VictimIndex).IpAddress Then
                AntiFrags_CheckUser = False
    
                Exit Function
    
            End If

        End If
        
        .Time = Time
        .UserName = UCase$(VictimName)
        .Account = UserList(VictimIndex).Account.Email
        .IP = UserList(VictimIndex).IpAddress
        
        'Call WriteLogSecurity( "Victima con IP: " & .IP, UserList(UserIndex).Account.Email, .Account, eSubType_Security.eAntiFrags)
    End With
    
    AntiFrags_CheckUser = True

    '<EhFooter>
    Exit Function

AntiFrags_CheckUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.AntiFrags_CheckUser " & "at line " & Erl
        
    '</EhFooter>
End Function

' Descontamos el tiempo del AntiFrags
Public Sub AntiFrags_CheckTime(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo AntiFrags_CheckTime_Err

    '</EhHeader>

    Dim A As Long
    
    For A = 1 To MAX_CONTROL_FRAGS

        With UserList(UserIndex).AntiFrags(A)

            If .Time > 0 Then
                .Time = .Time - 1
                
                If .Time <= 0 Then
                    .Time = 0
                    .UserName = vbNullString
                    .IP = vbNullString
                    .Account = vbNullString
                    .cant = 0

                End If

            End If

        End With

    Next A

    '<EhFooter>
    Exit Sub

AntiFrags_CheckTime_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.AntiFrags_CheckTime " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function CanUse_Inventory(ByVal UserIndex As Integer, _
                                 ByVal ObjIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo CanUse_Inventory_Err

    '</EhHeader>

    If ObjIndex <= 0 Then Exit Function
    
    With UserList(UserIndex)

        Select Case ObjData(ObjIndex).OBJType
            
            Case eOBJType.otPergaminos
                CanUse_Inventory = (InvUsuario.ClasePuedeUsarItem(UserIndex, ObjIndex))
                        
            Case eOBJType.otarmadura, eOBJType.otTransformVIP
                CanUse_Inventory = (InvUsuario.ClasePuedeUsarItem(UserIndex, ObjIndex) And InvUsuario.FaccionPuedeUsarItem(UserIndex, ObjIndex) And InvUsuario.SexoPuedeUsarItem(UserIndex, ObjIndex) And CheckRazaUsaRopa(UserIndex, ObjIndex))
                        
            Case eOBJType.otcasco
                CanUse_Inventory = ClasePuedeUsarItem(UserIndex, ObjIndex) And FaccionPuedeUsarItem(UserIndex, ObjIndex)
                  
            Case eOBJType.otescudo
                CanUse_Inventory = ClasePuedeUsarItem(UserIndex, ObjIndex) And FaccionPuedeUsarItem(UserIndex, ObjIndex)
                  
            Case eOBJType.otWeapon
                CanUse_Inventory = ClasePuedeUsarItem(UserIndex, ObjIndex) And FaccionPuedeUsarItem(UserIndex, ObjIndex)
                  
            Case eOBJType.otAnillo, eOBJType.otMagic
                CanUse_Inventory = ClasePuedeUsarItem(UserIndex, ObjIndex) And FaccionPuedeUsarItem(UserIndex, ObjIndex)
                  
            Case eOBJType.otFlechas
                CanUse_Inventory = ClasePuedeUsarItem(UserIndex, ObjIndex) And FaccionPuedeUsarItem(UserIndex, ObjIndex)
   
            Case Else
                CanUse_Inventory = True

        End Select
        
        If ObjData(ObjIndex).Bronce = 1 And .flags.Bronce = 0 Then
            CanUse_Inventory = False

        End If
        
        If ObjData(ObjIndex).Plata = 1 And .flags.Plata = 0 Then
            CanUse_Inventory = False

        End If
            
        If ObjData(ObjIndex).Oro = 1 And .flags.Oro = 0 Then
            CanUse_Inventory = False

        End If
        
        If ObjData(ObjIndex).Premium = 1 And .flags.Premium = 0 Then
            CanUse_Inventory = False

        End If

    End With

    '<EhFooter>
    Exit Function

CanUse_Inventory_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.CanUse_Inventory " & "at line " & Erl
        
    '</EhFooter>
End Function

' El anti pelotudos que no voy a necesitar porque el gm no necesitara sumonear y poner telep a futuro.
Public Function CanUserTelep(ByVal MapaActual As Integer, _
                             ByVal UserIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo CanUserTelep_Err

    '</EhHeader>

    With UserList(UserIndex)

        If .Counters.Pena > 0 Then Exit Function
        If .flags.SlotReto > 0 Then Exit Function
        If .flags.SlotEvent > 0 Then Exit Function
        If .flags.SlotFast > 0 Then Exit Function
        If .flags.Desafiando > 0 Then Exit Function
        If MapInfo(UserList(UserIndex).Pos.Map).Pk Then Exit Function
        If MapInfo(MapaActual).LvlMin > .Stats.Elv Then Exit Function
              
    End With
    
    CanUserTelep = True
    '<EhFooter>
    Exit Function

CanUserTelep_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.CanUserTelep " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub CheckingOcultation(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo CheckingOcultation_Err

    '</EhHeader>
    
    With UserList(UserIndex)

        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 0 Then
                If .flags.Invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, UserList(UserIndex).Char.charindex, False)
                    Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

    End With
    
    '<EhFooter>
    Exit Sub

CheckingOcultation_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.CheckingOcultation " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function Faction_String(ByVal Faction As eFaction) As String

    '<EhHeader>
    On Error GoTo Faction_String_Err

    '</EhHeader>
    
    Select Case Faction
    
        Case eFaction.fCrim
            Faction_String = "CRIMINAL"

        Case eFaction.fCiu
            Faction_String = "CIUDADANO"

        Case eFaction.fArmada
            Faction_String = "ARMADA REAL"

        Case eFaction.fLegion
            Faction_String = "LEGION OSCURA"

    End Select
    
    '<EhFooter>
    Exit Function

Faction_String_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.Faction_String " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function NpcInventory_GetAnimation(ByVal UserIndex As Integer, _
                                          ByVal ObjIndex As Integer) As Integer

    '<EhHeader>
    On Error GoTo NpcInventory_GetAnimation_Err

    '</EhHeader>

    If ObjIndex = 0 Then Exit Function
    
    With ObjData(ObjIndex)

        Select Case .OBJType

            Case eOBJType.otarmadura

                If .RopajeEnano <> 0 And (UserList(UserIndex).Raza = eRaza.Enano Or UserList(UserIndex).Raza = eRaza.Gnomo) Then
                    NpcInventory_GetAnimation = .RopajeEnano
                Else
                    NpcInventory_GetAnimation = .Ropaje

                End If

            Case eOBJType.otescudo
                NpcInventory_GetAnimation = .ShieldAnim
                
            Case eOBJType.otcasco
                NpcInventory_GetAnimation = .CascoAnim
                
            Case eOBJType.otWeapon

                If .WeaponRazaEnanaAnim <> 0 And (UserList(UserIndex).Raza = eRaza.Enano Or UserList(UserIndex).Raza = eRaza.Gnomo) Then
                    
                    NpcInventory_GetAnimation = .WeaponRazaEnanaAnim
                Else
                    NpcInventory_GetAnimation = .WeaponAnim

                End If

        End Select

    End With
    
    '<EhFooter>
    Exit Function

NpcInventory_GetAnimation_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.NpcInventory_GetAnimation " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function Respawn_Npc_Free(ByVal NpcIndex As Integer, _
                                 ByVal Map As Integer, _
                                 ByVal Time As Long, _
                                 ByVal CastleIndex As Integer, _
                                 ByRef OrigPos As WorldPos) As Boolean

    '<EhHeader>
    On Error GoTo Respawn_Npc_Free_Err

    '</EhHeader>

    Dim A As Long
    
    For A = 1 To RESPAWN_MAX

        With Respawn_Npc(A)

            If .Time = 0 Then
                .Map = Map
                .OrigPos = OrigPos
                .NpcIndex = NpcIndex
                .Time = Time
                .CastleIndex = CastleIndex
                    
                Respawn_Npc_Free = True

                Exit Function

            End If

        End With

    Next
    
    '<EhFooter>
    Exit Function

Respawn_Npc_Free_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.Respawn_Npc_Free " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub Loop_RespawnNpc()

    '<EhHeader>
    On Error GoTo Loop_RespawnNpc_Err

    '</EhHeader>

    Dim A       As Long

    Dim OrigPos As WorldPos
    
    For A = 1 To RESPAWN_MAX

        With Respawn_Npc(A)

            If .Time > 0 Then
                .Time = .Time - 1
                
                If .Time = 0 Then

                    Dim Npc As Integer

                    Npc = CrearNPC(.NpcIndex, .Map, .OrigPos)
                    
                    If Npc Then
                        Npclist(Npc).CastleIndex = .CastleIndex
                            
                        If Npclist(Npc).CastleIndex > 0 Then
                            Call mCastle.Castle_Close(Npclist(Npc).CastleIndex)

                        End If
                            
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡" & Npclist(Npc).Name & " en " & MapInfo(Npclist(Npc).Pos.Map).Name & "!", FontTypeNames.FONTTYPE_USERBRONCE))
                        
                        ' # Envia un mensaje a discord
                        Dim TextDiscord As String

                        TextDiscord = "--------------------"
                        TextDiscord = TextDiscord & vbCrLf & "¡**" & Npclist(Npc).Name & "** en **" & MapInfo(Npclist(Npc).Pos.Map).Name & "**!"
                            
                        If Npclist(Npc).NroDrops > 0 Or Npclist(Npc).Invent.NroItems > 0 Then
                            TextDiscord = TextDiscord & vbCrLf & Npclist(Npc).TempDrops

                        End If
                            
                        TextDiscord = TextDiscord & vbCrLf & "--------------------"
                            
                        WriteMessageDiscord CHANNEL_BOSSES, TextDiscord

                    End If
                    
                    .NpcIndex = 0
                    .Time = 0
                    .Map = 0
                    .CastleIndex = 0
                    .OrigPos.Map = 0
                    .OrigPos.X = 0
                    .OrigPos.Y = 0

                End If

            End If
        
        End With

    Next A

    '<EhFooter>
    Exit Sub

Loop_RespawnNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.Loop_RespawnNpc " & "at line " & Erl

    '</EhFooter>
End Sub

Public Function Is_Map_valid(ByVal UserIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo Is_Map_valid_Err

    '</EhHeader>

    With UserList(UserIndex)

        If .Pos.Map >= 74 And .Pos.Map <= 87 Then Exit Function
        If .Pos.Map = 24 Then Exit Function

    End With
    
    Is_Map_valid = True
    '<EhFooter>
    Exit Function

Is_Map_valid_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.Is_Map_valid " & "at line " & Erl
        
    '</EhFooter>
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean

    '<EhHeader>
    On Error GoTo CheckMailString_Err

    '</EhHeader>

    Dim lPos As Long

    Dim lX   As Long

    Dim iAsc As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")

    If (lPos <> 0) Then

        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1

            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))

                If Not CMSValidateChar_(iAsc) Then Exit Function

            End If

        Next lX
        
        'Finale
        CheckMailString = True

    End If

    '<EhFooter>
    Exit Function

CheckMailString_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.CheckMailString " & "at line " & Erl
        
    '</EhFooter>
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean

    '<EhHeader>
    On Error GoTo CMSValidateChar__Err

    '</EhHeader>
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 90) Or (iAsc >= 97 And iAsc <= 122) Or (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
    '<EhFooter>
    Exit Function

CMSValidateChar__Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.CMSValidateChar_ " & "at line " & Erl
        
    '</EhFooter>
End Function

' Chequeamos si los personajes estan en un mapa determinado.
Public Sub Checking_UsersInMap(ByVal Map As Integer)

    '<EhHeader>
    On Error GoTo Checking_UsersInMap_Err

    '</EhHeader>

    Dim X As Long, Y As Long
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(Map, X, Y).UserIndex <> 0 Then
                Call WarpUserChar(MapData(Map, X, Y).UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, False)

            End If
            
        Next X
    Next Y

    '<EhFooter>
    Exit Sub

Checking_UsersInMap_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.Checking_UsersInMap " & "at line " & Erl
        
    '</EhFooter>
End Sub

' # Elimina todos los objetos de un mapa que no se encuentren bloqueados.
Public Sub DeleteObjectMap(ByVal Map As Integer)

    '<EhHeader>
    On Error GoTo DeleteObjectMap_Err

    '</EhHeader>

    Dim LoopX As Long, LoopY As Long
       
    For LoopX = XMinMapSize To XMaxMapSize
        For LoopY = YMinMapSize To YMaxMapSize
    
            If InMapBounds(Map, LoopX, LoopY) Then
                    
                If MapData(Map, LoopX, LoopY).Blocked = 0 Then
                    EraseObj 10000, Map, LoopX, LoopY

                End If
                    
            End If
    
        Next LoopY
    Next LoopX

    '<EhFooter>
    Exit Sub

DeleteObjectMap_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.DeleteObjectMap " & "at line " & Erl
        
    '</EhFooter>
End Sub

' # Creación de un objeto en un mapa determinado
Public Sub Create_ObjectMap(ByVal ObjIndex As Integer, _
                            ByVal Amount As Integer, _
                            ByVal Map As Integer, _
                            ByVal X As Byte, _
                            ByVal Y As Byte, _
                            ByVal ObjEvent As Byte)

    '<EhHeader>
    On Error GoTo Create_ObjectMap_Err

    '</EhHeader>
          
    Dim Pos As WorldPos

    Dim Obj As Obj
          
    Pos.Map = Map
    Pos.X = X
    Pos.Y = Y
    
    Obj.ObjIndex = ObjIndex
    Obj.Amount = Amount
    
    Call TirarItemAlPiso(Pos, Obj)
    MapData(Pos.Map, Pos.X, Pos.Y).ObjEvent = ObjEvent

    '<EhFooter>
    Exit Sub

Create_ObjectMap_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.Create_ObjectMap " & "at line " & Erl
        
    '</EhFooter>
End Sub

' # Quitamos las criaturas de un Mapa

Public Sub Remove_All_Map(ByVal Map As Integer, _
                          ByVal NpcIndex As Integer, _
                          ByVal ObjIndex As Integer)

    '<EhHeader>
    On Error GoTo Remove_All_Map_Err

    '</EhHeader>

    Dim A As Long, B As Long
    
    For A = YMinMapSize To YMaxMapSize
        For B = XMinMapSize To XMaxMapSize

            If InMapBounds(Map, A, B) Then
                If NpcIndex Then
                    If MapData(Map, A, B).NpcIndex > 0 Then
                        If Npclist(MapData(Map, A, B).NpcIndex).Attackable = 1 Then
                            Call QuitarNPC(MapData(Map, A, B).NpcIndex)

                        End If

                    End If

                End If

                If ObjIndex Then
                    If MapData(Map, A, B).ObjInfo.ObjIndex > 0 Then
                        If ItemNoEsDeMapa(Map, A, B, False) Then
                            EraseObj 10000, Map, A, B

                        End If

                    End If

                End If

            End If

        Next B
    Next A
          
    '<EhFooter>
    Exit Sub

Remove_All_Map_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.Remove_All_Map " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function GetTime() As Double

    On Error GoTo ErrHandler

    Dim CurrentTick As Double

    CurrentTick = timeGetTime() And &H7FFFFFFF
    
    If CurrentTick < LastTick Then
        overflowCount = overflowCount + 1

    End If
    
    LastTick = CurrentTick
    
    ' Time since last overflow plus overflows times MAX_TIME
    GetTime = CurrentTick + overflowCount * MAX_TIME
    Exit Function

ErrHandler:
    GetTime = 0
    Call LogError("E$rror gettime")

End Function

' # Mapas válidos para que los Game Master puedan hacer sus eventos y sumonear usuarios
Public Function EventMaster_CheckMapvalid(ByVal UserIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo EventMaster_CheckMapvalid_Err

    '</EhHeader>
    
    With UserList(UserIndex)

        If (.Pos.Map = 74 Or .Pos.Map = 76 Or .Pos.Map = 77 Or .Pos.Map = 79 Or .Pos.Map = 80 Or .Pos.Map = 81 Or .Pos.Map = 82 Or .Pos.Map = 83 Or .Pos.Map = 84 Or .Pos.Map = 87) Then
            
            EventMaster_CheckMapvalid = True
            Exit Function
            
        End If
    
    End With
    
    '<EhFooter>
    Exit Function

EventMaster_CheckMapvalid_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.EventMaster_CheckMapvalid " & "at line " & Erl
        
    '</EhFooter>
End Function

' Determina si tiene algun item de los digitados
Public Function User_TieneObjetos_Especiales(ByVal UserIndex As Integer, _
                                             ByVal Bronce As Byte, _
                                             ByVal Plata As Byte, _
                                             ByVal Oro As Byte, _
                                             ByVal Premium) As Boolean

    '<EhHeader>
    On Error GoTo User_TieneObjetos_Especiales_Err

    '</EhHeader>

    Dim A        As Long

    Dim ObjIndex As Integer
    
    Dim Total    As Long

    For A = 1 To UserList(UserIndex).CurrentInventorySlots
        ObjIndex = UserList(UserIndex).Invent.Object(A).ObjIndex
        
        If ObjIndex > 0 Then
            If Bronce = 0 And ObjData(ObjIndex).Bronce = 1 Then
                User_TieneObjetos_Especiales = True
                Exit Function

            End If
                
            If Plata = 0 And ObjData(ObjIndex).Plata = 1 Then
                User_TieneObjetos_Especiales = True
                Exit Function

            End If
                
            If Oro = 0 And ObjData(ObjIndex).Oro = 1 Then
                User_TieneObjetos_Especiales = True
                Exit Function

            End If
            
            If Premium = 0 And ObjData(ObjIndex).Premium = 1 Then
                User_TieneObjetos_Especiales = True
                Exit Function

            End If

        End If

    Next A
    
    '<EhFooter>
    Exit Function

User_TieneObjetos_Especiales_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.User_TieneObjetos_Especiales " & "at line " & Erl
        
    '</EhFooter>
End Function

' Transforma los segundos en un tiempo determinado en (Horas Minutos & Segundos)
Public Function SecondsToHMS(ByVal Seconds As Long) As String

    '<EhHeader>
    On Error GoTo SecondsToHMS_Err

    '</EhHeader>

    Dim HR As Integer
    
    Dim MS As Integer
    
    Dim SS As Integer
        
    Dim DS As Integer
        
    DS = (Seconds \ 3600) \ 24
        
    HR = (Seconds \ 3600) Mod 24
    
    MS = (Seconds Mod 3600) \ 60
    
    SS = (Seconds Mod 3600) Mod 60
        
    SecondsToHMS = IIf(DS > 0, DS & " días ", vbNullString) & IIf(HR > 0, HR & " horas ", vbNullString) & IIf(MS > 0, MS & " minutos ", vbNullString) & IIf(SS > 0, SS & " segundos", vbNullString)

    '<EhFooter>
    Exit Function

SecondsToHMS_Err:
    LogError Err.description & vbCrLf & "in SecondsToHMS " & "at line " & Erl

    '</EhFooter>
End Function

Public Sub CheckHappyHour()

    '<EhHeader>
    On Error GoTo CheckHappyHour_Err

    '</EhHeader>
    HappyHour = Not HappyHour
    
    If HappyHour Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡HappyHour Activado! Exp x2 ¡Entrená tu personaje!", FontTypeNames.FONTTYPE_USERBRONCE))
    Else
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡HappyHour Desactivado!", FontTypeNames.FONTTYPE_USERBRONCE))

    End If

    '<EhFooter>
    Exit Sub

CheckHappyHour_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.CheckHappyHour " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub CheckPartyTime()

    '<EhHeader>
    On Error GoTo CheckPartyTime_Err

    '</EhHeader>
    PartyTime = Not PartyTime
    
    If PartyTime Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("PartyTime» Los miembros de la party reciben 25% de experiencia extra.", FontTypeNames.FONTTYPE_INVASION))
    Else
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("PartyTime» Desactivado!", FontTypeNames.FONTTYPE_INVASION))

    End If

    '<EhFooter>
    Exit Sub

CheckPartyTime_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.CheckPartyTime " & "at line " & Erl
        
    '</EhFooter>
End Sub

' @ Metodo Burbuja para ordenar arrays
Function BubbleSort(ByRef vIn As Variant, _
                    bAscending As Boolean, _
                    Optional vRet As Variant) As Boolean

    ' Sorts the single dimension list array, ascending or descending
    ' Returns sorted list in vRet if supplied, otherwise in vIn modified
    '<EhHeader>
    On Error GoTo BubbleSort_Err

    '</EhHeader>
        
    Dim First As Long, Last As Long

    Dim i     As Long, j As Long, bWasMissing As Boolean

    Dim Temp  As Variant, vW As Variant
    
    First = LBound(vIn)
    Last = UBound(vIn)
    
    ReDim vW(First To Last, 1)
    vW = vIn
    
    If bAscending = True Then

        For i = First To Last - 1
            For j = i + 1 To Last

                If vW(i) > vW(j) Then
                    Temp = vW(j)
                    vW(j) = vW(i)
                    vW(i) = Temp

                End If

            Next j
        Next i

    Else 'descending sort

        For i = First To Last - 1
            For j = i + 1 To Last

                If vW(i) < vW(j) Then
                    Temp = vW(j)
                    vW(j) = vW(i)
                    vW(i) = Temp

                End If

            Next j
        Next i

    End If
  
    'find whether optional vRet was initially missing
    bWasMissing = IsMissing(vRet)
   
    ' transfers
    If bWasMissing Then
        vIn = vW  'return in input array
    Else
        ReDim vRet(First To Last, 1)
        vRet = vW 'return with input unchanged

    End If
   
    BubbleSort = True

    '<EhFooter>
    Exit Function

BubbleSort_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.BubbleSort " & "at line " & Erl
    BubbleSort = False
        
    '</EhFooter>
End Function

Public Function Email_Is_Testing_Pro(ByVal Email As String) As Boolean

    '<EhHeader>
    On Error GoTo Email_Is_Testing_Pro_Err

    '</EhHeader>

    If Email = "marinolauta@gmail.com" Or Email = "montiel.marcoseze@gmail.com" Or Email = "gabi.barrantes.94@gmail.com" Or Email = "hogarcasa1991@gmail.com" Or Email = "chontecito@gmail.com" Or Email = "nuria_sabrina@hotmail.com" Or Email = "dreamlotao@gmail.com" Then
       
        Email_Is_Testing_Pro = True

    End If

    '<EhFooter>
    Exit Function

Email_Is_Testing_Pro_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.General.Email_Is_Testing_Pro " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Function

Public Function Detect_FirstDayNext()

    Dim currentDate As Date

    Dim newDate     As Date

    Dim currentHour As Date
    
    ' Get the current date and time
    currentDate = Now
    
    ' Get the current hour and minute portion
    currentHour = TimeValue(currentDate)
    
    ' Add one month to the current date
    newDate = DateAdd("m", 1, currentDate)
    
    ' Set the day to 1
    'newDate = DateSerial(Year(newDate), Month(newDate), 1)
    
    ' Combine the new date with the current hour and minute
    newDate = DateValue(newDate) + currentHour
    
    ' Display the new date and time in a MsgBox (for verification)
    Detect_FirstDayNext = newDate

End Function

Function CalcularPorcentajeBonificación(Exp As Long, usuariosOnline As Long) As Double

    Dim Porcentaje As Double
    
    If usuariosOnline > 100 Then
        ' A partir de 100 onlines: 3% por cada online
        Porcentaje = (usuariosOnline * 0.3) / 100
    Else
        ' Antes de 100 onlines: 1.5% por cada online
        Porcentaje = (usuariosOnline * 0.15) / 100

    End If

    ' Aplicar la bonificación al valor de 'exp'
    Dim bonificación As Double

    bonificación = Exp * Porcentaje

    CalcularPorcentajeBonificación = bonificación

End Function

Function CalcularPorcentajeBonificacion(ByVal Exp As Long) As Double

    On Error GoTo ErrHandler

    Dim Porcentaje As Double
    
    Porcentaje = (NumUsers + UsersBot) * 0.002
    
    ' Aplicar la bonificación al valor de 'exp'
    Dim bonificación As Double

    bonificación = Exp * Porcentaje

    CalcularPorcentajeBonificacion = bonificación
    
    Exit Function
ErrHandler:
    
End Function

' # Cargamos las frases al azar
Public Sub CargarFrasesOnFire()

    Dim FilePath As String
    
    FilePath = DatPath & "frases_on_fire.txt"
    FrasesOnFire = LeerFrasesDesdeArchivo(FilePath)

End Sub

' Función para leer las frases desde un archivo y almacenarlas en un array
Private Function LeerFrasesDesdeArchivo(ByVal rutaArchivo As String) As String()

    Dim frases()  As String

    Dim Contenido As String
    
    On Error Resume Next

    Open rutaArchivo For Binary As #1

    If Err.number = 0 Then
        Contenido = InputB(LOF(1), #1)
        Close #1
    Else
        ' Manejar el error, por ejemplo, mostrar un mensaje o registrar el error
        MsgBox "Error al leer el archivo de frases: " & Err.description, vbExclamation
        Exit Function

    End If

    On Error GoTo 0
    
    ' Convierte los bytes a una cadena Unicode
    Contenido = StrConv(Contenido, vbUnicode)
    
    ' Divide la cadena en frases utilizando vbCrLf
    frases = Split(Contenido, vbCrLf)
    
    LeerFrasesDesdeArchivo = frases

End Function

