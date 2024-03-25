Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martín Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20060517
' Reparado por Lorwik

Option Explicit

Public RequestCount    As Integer

Public LastRequestTime As Double

Public Type tImageData

    Bytes() As Byte

End Type

Public ImageData             As tImageData

Public SLOT_TERMINAL_ARCHIVE As Integer ' Connection Index: Programa externo encargado de la manipulación de archivos.

Public Enum eSearchData

    eMac = 1
    eDisk = 2
    eIpAddress = 3

End Enum

Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long

Public Declare Function lstrlen _
               Lib "kernel32" _
               Alias "lstrlenA" (ByVal lpString As Any) As Long

Public Declare Sub MemCopy _
               Lib "kernel32" _
               Alias "RtlMoveMemory" (Dest As Any, _
                                      Src As Any, _
                                      ByVal cb As Long)

Private Const CLIENT_XOR_KEY As Long = 192

Private Const SERVER_XOR_KEY As Long = 128

Private Declare Function ntohl Lib "ws2_32" (ByVal netlong As Long) As Long

'We'll pass Long host address values in lieu of this struct:
Private Type in_addr

    s_b1 As Byte
    s_b2 As Byte
    s_b3 As Byte
    s_b4 As Byte

End Type

Private Declare Function RtlIpv4AddressToString _
                Lib "ntdll" _
                Alias "RtlIpv4AddressToStringW" (ByRef Addr As Any, _
                                                 ByVal pS As Long) As Long

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Public Const SEPARATOR As String * 1 = vbNullChar

Public Enum eMessageType

    Info = 0
    Admin = 1
    Guild = 2
    Party = 3
    Combate = 4
    Trabajo = 5
    m_MOTD = 6
    cEvents_Curso = 7
    cEvents_General = 8

End Enum

Private Enum ServerPacketID
    
    Connected
    loggedaccount
    LoggedAccountBatle
    AccountInfo
    logged                  ' LOGGED
    LoggedRemoveChar
    LoggedAccount_DataChar
    
    SendIntervals
    
    Mercader_List
    Mercader_ListOffer
    Mercader_ListInfo

    MiniMap_InfoCriature
    
    Render_CountDown
    
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU
    UserCommerceEnd         ' FINCOMUSUOK
    UserOfferConfirm
    CommerceChat
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateDsp               ' PACKETDSP
    UpdateBankGold
    UpdateExp               ' ASE
    ChangeMap               ' CM
    PosUpdate               ' PU
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat               ' |+
    ShowMessageBox          ' !!
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterChangeHeading  ' CCH
    CharacterRemove         ' BP
    CharacterChangeNick
    CharacterMove           ' MP, +, * and _ '
    CharacterAttackMovement
    CharacterAttackNpc
    ForceCharMove
    CharacterChange         ' CP
    ObjectCreate            ' HO
    ObjectDelete            ' BO
    BlockPosition           ' BQ
    PlayMusic               ' TM
    PlayWave              ' TW
    StopWaveMap
    PauseToggle             ' BKW
    CreateFX                ' CFX
    UpdateUserStats         ' EST
    ChangeInventorySlot     ' CSI
    ChangeBankSlot          ' SBO
    ChangeBankSlot_Account
    ChangeSpellSlot         ' SHS
    Atributes               ' ATR
    RestOK                  ' DOK
    ErrorMsg                ' ERR
    Blind                   ' CEGU
    Dumb                    ' DUMB
    ChangeNPCInventorySlot  ' NPCI
    UpdateHungerAndThirst   ' EHYS
    MiniStats               ' MEST
    LevelUp                 ' SUNI
    SetInvisible            ' NOVER
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    TradeOK                 ' TRANSOK
    BankOK                  ' BANCOOK
    ChangeUserTradeSlot     ' COMUSUINV          '
    Pong
    UpdateTagAndStatus
    
    'GM messages
    SpawnList               ' SPL
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
    ShowDenounces
    RecordList
    RecordDetails
    
    UpdateStrenghtAndDexterity
    UpdateStrenght
    UpdateDexterity
    AddSlots
    MultiMessage
    CancelOfferItem
    ShowMenu
    StrDextRunningOut
    ChatPersonalizado
    GroupPrincipal
    GroupUpdateExp
    
    UserInEvent
    SendInfoRetos
    
    MontateToggle
    SolicitaCapProc
    UpdateListSecurity
    CreateDamage
    
    ClickVesA
    
    UpdateControlPotas
    
    UpdateInfoIntervals
    UpdateGroupIndex

    ' Clanes
    Guild_List
    Guild_Info
    Guild_InfoUsers
    
    Fight_PanelAccept
    
    UpdateEffectPoison
    CreateFXMap
    RenderConsole
    ViewListQuest
    UpdateUserDead
    QuestInfo
    UpdateGlobalCounter
    SendInfoNpc
    UpdatePosGuild
    UpdateLevelGuild
    UpdateStatusMAO
    UpdateOnline
    UpdateEvento
    UpdateMeditation
    SendShopChars
    UpdateFinishQuest
    UpdateDataSkin
    RequiredMoveChar
    UpdateBar
    UpdateBarTerrain
    VelocidadToggle
    SpeedToChar
    UpdateUserTrabajo
    TournamentList
    
    StatsUser
    StatsUser_Inventory
    StatsUser_Spells
    StatsUser_Bank
    StatsUser_Skills
    StatsUser_Bonos
    StatsUser_Penas
    StatsUser_Skins
    StatsUser_Logros
    
    UpdateClient

End Enum

Public Enum ClientPacketID

    LoginAccount
    LoginChar
    LoginCharNew
    LoginRemove
    LoginName
    ChangeClass
   
    DragToggle
    RequestAtributes        'ATR
  
    Train                   'ENTR
    CommerceBuy             'COMP
    BankExtractItem         'RETI
    CommerceSell            'VEND
    moveItem
    RightClick
    UserEditation
    
    ' Paquetes exclusivos de sistema de ROL (Desterium AO)
    PartyClient
    GroupChangePorc
    SendReply               'SendReply
    AcceptReply             'AcceptReply
    AbandonateReply         'AbandonateReply
    Entrardesafio
    SetPanelClient
    ChatGlobal
    LearnMeditation
    InfoEvento
    
    DragToPos
    Enlist
    Reward
    
    Fianza
    Home
    
    AbandonateFaction
    SendListSecurity
    BankDeposit             'DEPO
    MoveSpell               'DESPHE
    MoveBank
    UserCommerceOffer       'OFRECER
    Online                  '/ONLINE
    Quit                    '/SALIR
    Meditate                '/MEDITAR
    Resucitate              '/RESUCITAR
    Heal                    '/CURAR
    Help                    '/AYUDA
    RequestStats            '/EST
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    PartyMessage            '/PMSG
    CouncilMessage          '/BMSG
    ChangeDescription       '/DESC
    Punishments             '/PENAS
    Gamble                  '/APOSTAR
    BankGold
    Denounce                '/DENUNCIAR
    Ping                    '/PING
    GmCommands
    InitCrafting
    ShareNpc                '/COMPARTIR
    StopSharingNpc
    Consultation
    Event_Participe
    
    RequestSkills           'ESKI
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    UserCommerceConfirm
    CommerceChat
    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    Drop                    'TI
    CastSpell               'LH
    LeftClick               'LC
    DoubleClick             'RC
    Work                    'UK
    UseItem                 'USA
    UseItemTwo
    CraftBlacksmith         'CNS
    WorkLeftClick           'WLC
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    ModifySkills            'SKSE
    
    UpdateInactive
    
    Retos_RewardObj
    Mercader_New
    Mercader_Required
    
    Map_RequiredInfo
    Forgive_Faction
    WherePower
    
    Auction_New
    Auction_Info
    Auction_Offer
    
    GoInvation
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle

    Guilds_Required
    Guilds_Found
    Guilds_Invitation
    Guilds_Online
    Guilds_Kick
    Guilds_Abandonate
    Guilds_Talk
    
    Fight_CancelInvitation
    Events_DonateObject
    QuestRequired
    ModoStreamer
    StreamerSetLink
    ChangeNick
    ConfirmTransaccion
    ConfirmItem
    ConfirmTier
    RequiredShopChars
    ConfirmChar
    ConfirmQuest
    RequiredSkins
    RequiredLive
    AcelerationChar
    AlquilarComerciante
    TirarRuleta
    CastleInfo
    RequiredStatsUser
    CentralServer = 249
    [PacketCount]

End Enum

Public PacketUseItem  As ClientPacketID

Public PacketWorkLeft As ClientPacketID

Public Enum FontTypeNames

    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFORED
    FONTTYPE_INFOGREEN
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
    FONTTYPE_CONSE
    FONTTYPE_DIOS
    FONTTYPE_EVENT
    FONTTYPE_USERGOLD
    FONTTYPE_USERPREMIUM
    FONTTYPE_USERBRONCE
    FONTTYPE_USERPLATA
    FONTTYPE_ANGEL
    FONTTYPE_DEMONIO
    FONTTYPE_GLOBAL
    FONTTYPE_ADMIN
    FONTTYPE_CRITICO
    FONTTYPE_INFORETOS
    FONTTYPE_INVASION
    FONTTYPE_PODER
    FONTTYPE_DESAFIOS
    FONTTYPE_STREAM
    FONTTYPE_RMSG

End Enum

Public Enum eEditOptions

    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
    eo_addGold
    eo_Vida
    eo_Poss

End Enum

Public Server As Network.Server

Public Writer As Network.Writer

Public Reader As Network.Reader

Public Function IsRequestAllowed() As Boolean

    IsRequestAllowed = True
    Exit Function
    
    ' Configura el límite de solicitudes por segundo
    Const RequestLimitPerSecond As Integer = 10 ' Cambia este valor según tus necesidades

    ' Obtiene el tiempo actual en segundos con precisión de milisegundos
    Dim currentTime             As Double

    currentTime = CDbl(Timer)

    ' Calcula el tiempo transcurrido desde la última solicitud
    Dim ElapsedSeconds As Double

    ElapsedSeconds = currentTime - LastRequestTime

    ' Si ha pasado más de 1 segundo, reinicia el contador de solicitudes
    If ElapsedSeconds >= 1 Then
        RequestCount = 0
        LastRequestTime = currentTime

    End If

    ' Verifica si se ha alcanzado el límite de solicitudes
    If RequestCount >= RequestLimitPerSecond Then
        ' La solicitud actual supera el límite
        IsRequestAllowed = False
    Else
        ' La solicitud está permitida
        RequestCount = RequestCount + 1
        IsRequestAllowed = True

    End If

End Function

Private Function verifyTimeStamp(ByVal ActualCount As Long, _
                                 ByRef LastCount As Long, _
                                 ByRef LastTick As Long, _
                                 ByRef Iterations, _
                                 ByVal UserIndex As Integer, _
                                 ByVal PacketName As String, _
                                 Optional ByVal DeltaThreshold As Long = 100, _
                                 Optional ByVal MaxIterations As Long = 5, _
                                 Optional ByVal CloseClient As Boolean = False) As Boolean
    
    Dim Ticks As Long, Delta As Long

    Ticks = GetTime
    
    Delta = (Ticks - LastTick)
    LastTick = Ticks

    'Controlamos secuencia para ver que no haya paquetes duplicados.
    If ActualCount <= LastCount Then
        Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg("Paquete grabado: " & PacketName & " | Cuenta: " & UserList(UserIndex).Account.Email & " | Ip: " & UserList(UserIndex).IpAddress & ". ", FontTypeNames.FONTTYPE_INFOBOLD))
        Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, "Paquete grabado: " & PacketName & " | Cuenta: " & UserList(UserIndex).Account.Email & " | Ip: " & UserList(UserIndex).IpAddress & ". ")
        LastCount = ActualCount
        ' Call CloseSocket(UserIndex)
        Exit Function

    End If
    
    'controlamos speedhack/macro
    If Delta < DeltaThreshold Then
        Iterations = Iterations + 1

        If Iterations >= MaxIterations Then
            'Call WriteShowMessageBox(UserIndex, "Relajate andá a tomarte un té con Gulfas.")
            verifyTimeStamp = False
            'Call LogMacroServidor("El usuario " & UserList(UserIndex).name & " iteró el paquete " & PacketName & " " & MaxIterations & " veces.")
            Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg("Control de macro---> El usuario " & UserList(UserIndex).Name & "| Revisar --> " & PacketName & " (Envíos: " & Iterations & ").", FontTypeNames.FONTTYPE_INFOBOLD))
            Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, "Control de macro---> El usuario " & UserList(UserIndex).Name & "| Revisar --> " & PacketName & " (Envíos: " & Iterations & ").")
            'Call WriteCerrarleCliente(UserIndex)
            'Call CloseSocket(UserIndex)
            LastCount = ActualCount
            Iterations = 0
            Debug.Print "CIERRO CLIENTE"

        End If

        'Exit Function
    Else
        Iterations = 0

    End If
        
    verifyTimeStamp = True
    LastCount = ActualCount

End Function

Public Sub Kick(ByVal Connection As Long, Optional ByVal Message As String = vbNullString)

    '<EhHeader>
    On Error GoTo Kick_Err

    '</EhHeader>
    
    If (Message <> vbNullString) Then
        Call WriteErrorMsg(Connection, Message)
        Call Server.Flush(Connection)

    End If
    
    If UserList(Connection).flags.UserLogged Then
        Call CloseSocket(Connection)

    End If
    
    Call Server.Flush(Connection)
    Call Server.Kick(Connection, True)
    '<EhFooter>
    Exit Sub

Kick_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.Kick " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function Ipv4NetAtoS(ByVal NetAddrLong As Long) As String

    On Error GoTo ErrHandler

    Dim pS   As Long

    Dim pEnd As Long

    Ipv4NetAtoS = Space$(15)
    pS = StrPtr(Ipv4NetAtoS)
    pEnd = RtlIpv4AddressToString(ntohl(NetAddrLong), pS)
    Ipv4NetAtoS = Left$(Ipv4NetAtoS, ((pEnd Xor &H80000000) - (pS Xor &H80000000)) \ 2)
    Exit Function
ErrHandler:
    Ipv4NetAtoS = "255.255.255.0"

End Function

Function NextOpenUser() As Integer
        
    On Error GoTo NextOpenUser_Err

    Dim LoopC As Long
   
    For LoopC = 1 To MaxUsers + 1

        If LoopC > MaxUsers Then Exit For
        If (Not UserList(LoopC).ConnIDValida And UserList(LoopC).flags.UserLogged = False) Then Exit For
    Next LoopC
   
    NextOpenUser = LoopC
        
    Exit Function

NextOpenUser_Err:
        
End Function

Public Sub OnServerConnect(ByVal Connection As Long, ByVal Address As String)

    '<EhHeader>
    On Error GoTo OnServerConnect_Err

    '</EhHeader>

    Dim FreeUser As Long
            
    If Connection <= MaxUsers Then
        FreeUser = NextOpenUser()
            
        UserList(FreeUser).ConnIDValida = True
        UserList(FreeUser).IpAddress = Address

        If FreeUser >= LastUser Then LastUser = FreeUser
                    
        Dim Server As Byte

        Server = 0
                    
        Call WriteConnectedMessage(Connection, Server)
    Else
        Call Protocol.Kick(Connection, "El servidor se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")

    End If
    
    '<EhFooter>
    Exit Sub

OnServerConnect_Err:
    Call Kick(Connection)
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.OnServerConnect " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub OnServerClose(ByVal Connection As Long)

    On Error GoTo OnServerClose_Error
    
    If Not (Connection = SLOT_TERMINAL_ARCHIVE) Then
        If UserList(Connection).AccountLogged Then
            Call mAccount.DisconnectAccount(Connection)

        End If

    Else
        SLOT_TERMINAL_ARCHIVE = 0

    End If

    UserList(Connection).ConnIDValida = False
    UserList(Connection).IpAddress = vbNullString
    
    Call FreeSlot(Connection)
    
    Exit Sub

OnServerClose_Error:

    Call LogError("OnServerClose: " + Err.description)
    
End Sub

Public Sub OnServerSend(ByVal Connection As Long, ByVal Message As Network.Reader)

    '<EhHeader>
    On Error GoTo OnServerSend_Err

    '</EhHeader>

    '<EhFooter>
    Exit Sub

OnServerSend_Err:
    Call Kick(Connection)
    
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.OnServerSend " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub OnServerReceive(ByVal Connection As Long, ByVal Message As Network.Reader)

    '<EhHeader>
    On Error GoTo OnServerReceive_Err

    '</EhHeader>

    'Debug.Print "OnServerReceive"

    ' Dim BufferRef() As Byte
    ' Call message.GetData(BufferRef)
    
    '  Dim i As Long
    ' For i = 0 To UBound(BufferRef)
    ' BufferRef(i) = BufferRef(i) Xor CLIENT_XOR_KEY
    ' Next i
    
    '   Set Reader = message
    
    '   While (message.GetAvailable() > 0)

    Call HandleIncomingData(Connection, Message)

    'Wend
    
    ' Set Reader = Nothing
    '<EhFooter>
    Exit Sub

OnServerReceive_Err:
    Call Kick(Connection)
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.OnServerReceive " & "at line " & Erl

    '</EhFooter>
End Sub

''
' Handles incoming data.
'
' @param    userIndex The index of the user sending the message.

Public Function HandleIncomingData(ByVal UserIndex As Integer, _
                                   ByVal Message As Network.Reader) As Boolean

    On Error Resume Next

    Set Reader = Message
        
    Dim PacketID As Long
    
    PacketID = Reader.ReadInt
    
    Dim Time As Long
    
    Time = GetTime()
        
    If Time - UserList(UserIndex).Counters.TimeLastReset >= 5000 Then
        UserList(UserIndex).Counters.TimeLastReset = Time
        UserList(UserIndex).Counters.PacketCount = 0

    End If
    
    UserList(UserIndex).Counters.PacketCount = UserList(UserIndex).Counters.PacketCount + 1
    
    If UserList(UserIndex).Counters.PacketCount > 100 Then
        Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, "Control de paquetes -> La cuenta " & UserList(UserIndex).Account.Email & " en personaje " & UserList(UserIndex).Name & " | IP: " & UserList(UserIndex).Account.Sec.IP_Address & " | Iteración paquetes | Último paquete: " & PacketID & ".")
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Control de paquetes -> La cuenta " & UserList(UserIndex).Account.Email & " en personaje " & UserList(UserIndex).Name & "  | Iteración paquetes | Último paquete: " & PacketID & ".", FontTypeNames.FONTTYPE_FIGHT))
        UserList(UserIndex).Counters.PacketCount = 0
        '  Exit Function

    End If
    
    If PacketID < 0 Or PacketID >= ClientPacketID.PacketCount Then
        'Call Logs_Security(eSecurity, eAntiHack, "La cuenta " & UserList(UserIndex).Account.Email & " con IP: " & UserList(UserIndex).IpAddress & " mando fake paquet " & PacketID)

        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("La cuenta " & UserList(UserIndex).Account.Email & " con IP: " & UserList(UserIndex).IpAddress & " mando fake paquet " & PacketID, FontTypeNames.FONTTYPE_SERVER))
        'Call Protocol.Kick(UserIndex)
            
        ' Exit Function

    End If
    
    'Does the packet requires a logged user??
    If Not (PacketID = ClientPacketID.LoginChar Or PacketID = ClientPacketID.LoginCharNew Or PacketID = ClientPacketID.LoginName Or PacketID = ClientPacketID.LoginAccount Or PacketID = ClientPacketID.LoginRemove Or PacketID = ClientPacketID.CentralServer) Then

        ' Si no está logeado en la cuenta no se permite enviar paquetes
        If Not UserList(UserIndex).AccountLogged Then
            'Call Logs_Security(eSecurity, eAntiHack, "La IP: " & UserList(UserIndex).IpAddress & " mando fake paquet " & PacketID)
            Call Protocol.Kick(UserIndex)
            Exit Function

        End If
    
        If Not (PacketID = ClientPacketID.UpdateInactive Or PacketID = ClientPacketID.Mercader_Required Or PacketID = ClientPacketID.Mercader_New) Then
                
            'Is the user actually logged?
            If Not UserList(UserIndex).flags.UserLogged Then
    
                Call CloseSocket(UserIndex)
                Exit Function
    
                'He is logged. Reset idle counter if id is valid.
            ElseIf PacketID <= ClientPacketID.[PacketCount] Then
                UserList(UserIndex).Counters.IdleCount = 0

            End If

        End If
        
    ElseIf PacketID <= ClientPacketID.[PacketCount] Then
        UserList(UserIndex).Counters.IdleCount = 0

        'Is the user logged?
            
        If UserList(UserIndex).flags.UserLogged Then
            Call CloseSocket(UserIndex)

            Exit Function

        End If

    End If

    ' Ante cualquier paquete, pierde la proteccion de ser atacado.
    UserList(UserIndex).flags.NoPuedeSerAtacado = False
    
    'Lorwik> Se lo copie a Hide :P
    frmMain.lstDebug.AddItem " < [" & UserIndex & "] PacketID: " & PacketID
    
    Select Case PacketID
            
        Case ClientPacketID.RequiredStatsUser
            Call HandleRequiredStatsUser(UserIndex)
                
        Case ClientPacketID.CastleInfo
            Call HandleCastleInfo(UserIndex)
                
        Case ClientPacketID.TirarRuleta
            Call HandleTirarRuleta(UserIndex)
                
        Case ClientPacketID.AlquilarComerciante
            Call HandleAlquilarComerciante(UserIndex)
            
        Case ClientPacketID.AcelerationChar
            Call HandleAcelerationChar(UserIndex)
                
        Case ClientPacketID.RequiredLive
            Call HandleRequiredLive(UserIndex)
                
        Case ClientPacketID.RequiredSkins
            Call HandleRequiredSkin(UserIndex)
                
        Case ClientPacketID.ConfirmQuest
            Call HandleConfirmQuest(UserIndex)
                
        Case ClientPacketID.ConfirmChar
            Call HandleConfirmChar(UserIndex)
                
        Case ClientPacketID.RequiredShopChars
            Call HandleRequiredShopChars(UserIndex)
                
        Case ClientPacketID.ConfirmTier
            Call HandleConfirmTier(UserIndex)
                
        Case ClientPacketID.ConfirmItem
            Call HandleConfirmItem(UserIndex)
                
        Case ClientPacketID.ConfirmTransaccion
            Call HandleConfirmTransaccion(UserIndex)
                
        Case ClientPacketID.ChangeNick
            Call HandleChangeNick(UserIndex)
                
        Case ClientPacketID.StreamerSetLink
            Call HandleStreamerSetLink(UserIndex)
                
        Case ClientPacketID.ModoStreamer
            Call HandleModoStreamer(UserIndex)
                
        Case ClientPacketID.ChangeClass
            Call HandleChangeClass(UserIndex)
                    
        Case ClientPacketID.LoginName
            Call HandleLoginName(UserIndex)
                
        Case ClientPacketID.CentralServer
            Call HandleCentralServer(UserIndex)

        Case ClientPacketID.Fight_CancelInvitation
            Call HandleFight_CancelInvitation(UserIndex)
            
        Case ClientPacketID.Guilds_Talk
            Call HandleGuilds_Talk(UserIndex)
            
        Case ClientPacketID.Guilds_Abandonate
            Call HandleGuilds_Abandonate(UserIndex)
            
        Case ClientPacketID.Guilds_Kick
            Call HandleGuilds_Kick(UserIndex)
            
        Case ClientPacketID.Guilds_Online
            Call HandleGuilds_Online(UserIndex)
        
        Case ClientPacketID.Guilds_Invitation
            Call HandleGuilds_Invitation(UserIndex)
            
        Case ClientPacketID.Guilds_Found
            Call HandleGuilds_Found(UserIndex)
            
        Case ClientPacketID.Guilds_Required
            Call HandleGuilds_Required(UserIndex)
            
        Case ClientPacketID.Retos_RewardObj
            Call HandleRetos_RewardObj(UserIndex)
            
        Case ClientPacketID.UpdateInactive
            Call HandleUpdateInactive(UserIndex)
            
        Case ClientPacketID.SendReply
            Call HandleSendReply(UserIndex)
            
        Case ClientPacketID.AcceptReply
            Call HandleAcceptReply(UserIndex)
            
        Case ClientPacketID.AbandonateReply
            Call HandleAbandonateReply(UserIndex)
        
        Case ClientPacketID.SendListSecurity
            Call HandleSendListSecurity(UserIndex)
            
        Case ClientPacketID.Event_Participe
            Call HandleEvent_Participe(UserIndex)
            
        Case ClientPacketID.AbandonateFaction
            Call HandleAbandonateFaction(UserIndex)

        Case ClientPacketID.LoginRemove
            Call HandleLoginRemove(UserIndex)
            
        Case ClientPacketID.LoginAccount
            Call HandleLoginAccount(UserIndex)
        
        Case ClientPacketID.Mercader_New
            Call HandleMercader_New(UserIndex)
            
        Case ClientPacketID.Mercader_Required
            Call HandleMercader_Required(UserIndex)
            
        Case ClientPacketID.Forgive_Faction
            Call HandleForgive_Faction(UserIndex)
        
        Case ClientPacketID.WherePower
            Call HandleWherePower(UserIndex)
        
        Case ClientPacketID.Auction_New
            Call HandleAuction_New(UserIndex)
        
        Case ClientPacketID.Auction_Info
            Call HandleAuction_Info(UserIndex)
            
        Case ClientPacketID.Auction_Offer
            Call HandleAuction_Offer(UserIndex)
            
        Case ClientPacketID.GoInvation
            Call HandleGoInvation(UserIndex)
            
        Case ClientPacketID.Map_RequiredInfo
            Call HandleMap_RequiredInfo(UserIndex)
            
        Case ClientPacketID.LoginChar
            Call HandleLoginChar(UserIndex)
            
        Case ClientPacketID.LoginCharNew
            Call HandleLoginCharNew(UserIndex)
            
        Case ClientPacketID.Entrardesafio
            Call HandleEntrarDesafio(UserIndex)
            
        Case ClientPacketID.SetPanelClient
            Call HandleSetPanelClient(UserIndex)
            
        Case ClientPacketID.GroupChangePorc
            Call HandleGroupChangePorc(UserIndex)
            
        Case ClientPacketID.PartyClient
            Call HandlePartyClient(UserIndex)
        
        Case ClientPacketID.Talk                    ';
            Call HandleTalk(UserIndex)
        
        Case ClientPacketID.Yell                    '-
            Call HandleYell(UserIndex)
        
        Case ClientPacketID.Whisper                 '\
            Call HandleWhisper(UserIndex)
        
        Case ClientPacketID.Walk                    'M
            Call HandleWalk(UserIndex)
        
        Case ClientPacketID.RequestPositionUpdate   'RPU
            Call HandleRequestPositionUpdate(UserIndex)
        
        Case ClientPacketID.Attack                  'AT
            Call HandleAttack(UserIndex)
        
        Case ClientPacketID.PickUp                  'AG
            Call HandlePickUp(UserIndex)
        
        Case ClientPacketID.SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
            Call HandleSafeToggle(UserIndex)
        
        Case ClientPacketID.ResuscitationSafeToggle
            Call HandleResuscitationToggle(UserIndex)
            
        Case ClientPacketID.DragToggle
            Call HandleDragToggle(UserIndex)
        
        Case ClientPacketID.RequestAtributes        'ATR
            Call HandleRequestAtributes(UserIndex)
        
        Case ClientPacketID.RequestSkills           'ESKI
            Call HandleRequestSkills(UserIndex)
        
        Case ClientPacketID.RequestMiniStats        'FEST
            Call HandleRequestMiniStats(UserIndex)
        
        Case ClientPacketID.CommerceEnd             'FINCOM
            Call HandleCommerceEnd(UserIndex)
            
        Case ClientPacketID.CommerceChat
            Call HandleCommerceChat(UserIndex)
        
        Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
            Call HandleUserCommerceEnd(UserIndex)
            
        Case ClientPacketID.UserCommerceConfirm
            Call HandleUserCommerceConfirm(UserIndex)
        
        Case ClientPacketID.BankEnd                 'FINBAN
            Call HandleBankEnd(UserIndex)
        
        Case ClientPacketID.UserCommerceOk          'COMUSUOK
            Call HandleUserCommerceOk(UserIndex)
        
        Case ClientPacketID.UserCommerceReject      'COMUSUNO
            Call HandleUserCommerceReject(UserIndex)
        
        Case ClientPacketID.Drop                    'TI
            Call HandleDrop(UserIndex)
        
        Case ClientPacketID.CastSpell               'LH
            Call HandleCastSpell(UserIndex)
        
        Case ClientPacketID.LeftClick               'LC
            Call HandleLeftClick(UserIndex)
        
        Case ClientPacketID.DoubleClick             'RC
            Call HandleDoubleClick(UserIndex)
        
        Case ClientPacketID.Work                    'UK
            Call HandleWork(UserIndex)
        
        Case ClientPacketID.UseItem                 'USA
            Call HandleUseItem(UserIndex)
        
        Case ClientPacketID.UseItemTwo                 'USA
            Call HandleUseItemTwo(UserIndex)

        Case ClientPacketID.CraftBlacksmith         'CNS
            Call HandleCraftBlacksmith(UserIndex)
        
        Case ClientPacketID.WorkLeftClick           'WLC
            Call HandleWorkLeftClick(UserIndex)
        
        Case ClientPacketID.SpellInfo               'INFS
            Call HandleSpellInfo(UserIndex)
        
        Case ClientPacketID.EquipItem               'EQUI
            Call HandleEquipItem(UserIndex)
        
        Case ClientPacketID.ChangeHeading           'CHEA
            Call HandleChangeHeading(UserIndex)
        
        Case ClientPacketID.ModifySkills            'SKSE
            Call HandleModifySkills(UserIndex)
        
        Case ClientPacketID.Train                   'ENTR
            Call HandleTrain(UserIndex)
        
        Case ClientPacketID.CommerceBuy             'COMP
            Call HandleCommerceBuy(UserIndex)
        
        Case ClientPacketID.BankExtractItem         'RETI
            Call HandleBankExtractItem(UserIndex)
        
        Case ClientPacketID.CommerceSell            'VEND
            Call HandleCommerceSell(UserIndex)
        
        Case ClientPacketID.BankDeposit             'DEPO
            Call HandleBankDeposit(UserIndex)
        
        Case ClientPacketID.MoveSpell               'DESPHE
            Call HandleMoveSpell(UserIndex)
            
        Case ClientPacketID.MoveBank
            Call HandleMoveBank(UserIndex)
        
        Case ClientPacketID.UserCommerceOffer       'OFRECER
            Call HandleUserCommerceOffer(UserIndex)
         
        Case ClientPacketID.Online                  '/ONLINE
            Call HandleOnline(UserIndex)
        
        Case ClientPacketID.Quit                    '/SALIR
            Call HandleQuit(UserIndex)
        
        Case ClientPacketID.Meditate                '/MEDITAR
            Call HandleMeditate(UserIndex)
        
        Case ClientPacketID.Resucitate              '/RESUCITAR
            Call HandleResucitate(UserIndex)
        
        Case ClientPacketID.Heal                    '/CURAR
            Call HandleHeal(UserIndex)
        
        Case ClientPacketID.Help                    '/AYUDA
            Call HandleHelp(UserIndex)
        
        Case ClientPacketID.RequestStats            '/EST
            Call HandleRequestStats(UserIndex)
        
        Case ClientPacketID.CommerceStart           '/COMERCIAR
            Call HandleCommerceStart(UserIndex)
        
        Case ClientPacketID.BankStart               '/BOVEDA
            Call HandleBankStart(UserIndex)
        
        Case ClientPacketID.PartyMessage            '/PMSG
            Call HandlePartyMessage(UserIndex)
        
        Case ClientPacketID.CouncilMessage          '/BMSG
            Call HandleCouncilMessage(UserIndex)
        
        Case ClientPacketID.ChangeDescription       '/DESC
            Call HandleChangeDescription(UserIndex)
        
        Case ClientPacketID.Punishments             '/PENAS
            Call HandlePunishments(UserIndex)
        
        Case ClientPacketID.Gamble                  '/APOSTAR
            Call HandleGamble(UserIndex)
        
        Case ClientPacketID.BankGold
            Call HandleBankGold(UserIndex)
            
        Case ClientPacketID.Denounce                '/DENUNCIAR
            Call HandleDenounce(UserIndex)
        
        Case ClientPacketID.Ping                    '/PING
            Call HandlePing(UserIndex)
        
        Case ClientPacketID.GmCommands              'GM Messages
            Call HandleGMCommands(UserIndex)
            
        Case ClientPacketID.InitCrafting
            Call HandleInitCrafting(UserIndex)
            
        Case ClientPacketID.ShareNpc                '/COMPARTIR
            Call HandleShareNpc(UserIndex)
            
        Case ClientPacketID.StopSharingNpc
            Call HandleStopSharingNpc(UserIndex)
            
        Case ClientPacketID.Consultation
            Call HandleConsultation(UserIndex)
        
        Case ClientPacketID.moveItem
            Call HandleMoveItem(UserIndex)
            
        Case ClientPacketID.RightClick
            Call HandleRightClick(UserIndex)
            
        Case ClientPacketID.UserEditation
            Call HandleUserEditation(UserIndex)
            
        Case ClientPacketID.ChatGlobal
            Call HandleChatGlobal(UserIndex)
            
        Case ClientPacketID.LearnMeditation
            Call HandleLearnMeditation(UserIndex)
            
        Case ClientPacketID.InfoEvento
            Call HandleInfoEvento(UserIndex)
        
        Case ClientPacketID.DragToPos
            Call HandleDragToPos(UserIndex)
            
        Case ClientPacketID.Enlist
            Call HandleEnlist(UserIndex)
            
        Case ClientPacketID.Reward
            Call HandleReward(UserIndex)
            
        Case ClientPacketID.Fianza
            Call HandleFianza(UserIndex)
            
        Case ClientPacketID.Home
            'Call HandleHome(UserIndex)
        
        Case ClientPacketID.Events_DonateObject
            Call HandleEvents_DonateObject(UserIndex)
            
        Case ClientPacketID.QuestRequired
            Call HandleQuestRequired(UserIndex)
                 
        Case Else
            Err.Raise -1, "Invalid Message"

    End Select
        
    If (Message.GetAvailable() > 0) Then

        '  Err.Raise &HDEADBEEF, "HandleIncomingData", "El paquete '" & PacketID & "' se encuentra en mal estado con '" & message.GetAvailable() & "' bytes de mas por el usuario '" & UserList(UserIndex).Name & "'"
    End If
    
HandleIncomingData_Err:
    
    Set Reader = Nothing

    If Err.number <> 0 Then
        Call LogError(Err.description & vbNewLine & "PackedID: " & PacketID & vbNewLine & IIf(UserList(UserIndex).flags.UserLogged, "Usuario: " & UserList(UserIndex).Name, "UserIndex: " & UserIndex & " con IP: " & UserList(UserIndex).IpAddress & " Email: " & UserList(UserIndex).Account.Email))
        'Call CloseSocket(UserIndex)
        
        HandleIncomingData = False

    End If

End Function

Public Sub WriteMultiMessage(ByVal UserIndex As Integer, _
                             ByVal MessageIndex As Integer, _
                             Optional ByVal Arg1 As Long, _
                             Optional ByVal Arg2 As Long, _
                             Optional ByVal Arg3 As Long, _
                             Optional ByVal StringArg1 As String)

    '<EhHeader>
    On Error GoTo WriteMultiMessage_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.MultiMessage)
    Call Writer.WriteInt(MessageIndex)
        
    Select Case MessageIndex

        Case eMessages.DontSeeAnything, eMessages.NPCSwing, eMessages.NPCKillUser, eMessages.BlockedWithShieldUser, eMessages.BlockedWithShieldother, eMessages.UserSwing, eMessages.SafeModeOn, eMessages.SafeModeOff, eMessages.ResuscitationSafeOff, eMessages.ResuscitationSafeOn, eMessages.NobilityLost, eMessages.CantUseWhileMeditating, eMessages.CancelHome, eMessages.FinishHome
            
        Case eMessages.NPCHitUser
            Call Writer.WriteInt(Arg1) 'Target
            Call Writer.WriteInt(Arg2) 'damage
                
        Case eMessages.UserHitNPC
            Call Writer.WriteInt(Arg1) 'damage
                
        Case eMessages.UserAttackedSwing
            Call Writer.WriteInt(UserList(Arg1).Char.charindex)
                
        Case eMessages.UserHittedByUser
            Call Writer.WriteInt(Arg1) 'AttackerIndex
            Call Writer.WriteInt(Arg2) 'Target
            Call Writer.WriteInt(Arg3) 'damage
                
        Case eMessages.UserHittedUser
            Call Writer.WriteInt(Arg1) 'AttackerIndex
            Call Writer.WriteInt(Arg2) 'Target
            Call Writer.WriteInt(Arg3) 'damage
                
        Case eMessages.WorkRequestTarget
            Call Writer.WriteInt(Arg1) 'skill
            
        Case eMessages.HaveKilledUser '"Has matado a " & UserList(VictimIndex).name & "!" "Has ganado " & DaExp & " puntos de experiencia."
            Call Writer.WriteInt(UserList(Arg1).Char.charindex) 'VictimIndex
            Call Writer.WriteInt(Arg2) 'Expe
            
        Case eMessages.UserKill '"¡" & .name & " te ha matado!"
            Call Writer.WriteInt(UserList(Arg1).Char.charindex) 'AttackerIndex
            
        Case eMessages.Home
            Call Writer.WriteInt(CByte(Arg1))
            Call Writer.WriteInt(CInt(Arg2))
            'El cliente no conoce nada sobre nombre de mapas y hogares, por lo tanto _
             hasta que no se pasen los dats e .INFs al cliente, esto queda así.
            Call Writer.WriteString8(StringArg1) 'Call .Writer.WriteInt(CByte(Arg2))
        
        Case eMessages.EarnExp

    End Select

    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteMultiMessage_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteMultiMessage " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleGMCommands(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleGMCommands_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Command As Byte

    With UserList(UserIndex)
    
        Command = Reader.ReadInt
    
        Select Case Command
                
            Case eGMCommands.GMMessage                '/GMSG
                Call HandleGMMessage(UserIndex)
        
            Case eGMCommands.ShowName                '/SHOWNAME
                Call HandleShowName(UserIndex)
        
            Case eGMCommands.serverTime              '/HORA
                Call HandleServerTime(UserIndex)
        
            Case eGMCommands.Where                   '/DONDE
                Call HandleWhere(UserIndex)
        
            Case eGMCommands.CreaturesInMap          '/NENE
                Call HandleCreaturesInMap(UserIndex)
        
            Case eGMCommands.WarpChar                '/TELEP
                Call HandleWarpChar(UserIndex)
        
            Case eGMCommands.Silence                 '/SILENCIAR
                Call HandleSilence(UserIndex)
        
            Case eGMCommands.GoToChar                '/IRA
                Call HandleGoToChar(UserIndex)
        
            Case eGMCommands.Invisible               '/INVISIBLE
                Call HandleInvisible(UserIndex)
        
            Case eGMCommands.GMPanel                 '/PANELGM
                Call HandleGMPanel(UserIndex)
        
            Case eGMCommands.RequestUserList         'LISTUSU
                Call HandleRequestUserList(UserIndex)
        
            Case eGMCommands.Jail                    '/CARCEL
                Call HandleJail(UserIndex)
        
            Case eGMCommands.KillNPC                 '/RMATA
                Call HandleKillNPC(UserIndex)
        
            Case eGMCommands.WarnUser                '/ADVERTENCIA
                Call HandleWarnUser(UserIndex)
        
            Case eGMCommands.RequestCharInfo         '/INFO
                Call HandleRequestCharInfo(UserIndex)
        
            Case eGMCommands.RequestCharInventory    '/INV
                Call HandleRequestCharInventory(UserIndex)
        
            Case eGMCommands.RequestCharBank         '/BOV
                Call HandleRequestCharBank(UserIndex)
        
            Case eGMCommands.ReviveChar              '/REVIVIR
                Call HandleReviveChar(UserIndex)
        
            Case eGMCommands.OnlineGM                '/ONLINEGM
                Call HandleOnlineGM(UserIndex)
        
            Case eGMCommands.OnlineMap               '/ONLINEMAP
                Call HandleOnlineMap(UserIndex)
        
            Case eGMCommands.Forgive                 '/PERDON
                Call HandleForgive(UserIndex)
        
            Case eGMCommands.Kick                    '/ECHAR
                Call HandleKick(UserIndex)
        
            Case eGMCommands.Execute                 '/EJECUTAR
                Call HandleExecute(UserIndex)
        
            Case eGMCommands.BanChar                 '/BAN
                Call HandleBanChar(UserIndex)
        
            Case eGMCommands.UnbanChar               '/UNBAN
                Call HandleUnbanChar(UserIndex)
        
            Case eGMCommands.NPCFollow               '/SEGUIR
                Call HandleNPCFollow(UserIndex)
        
            Case eGMCommands.SummonChar              '/SUM
                Call HandleSummonChar(UserIndex)
        
            Case eGMCommands.SpawnListRequest        '/CC
                Call HandleSpawnListRequest(UserIndex)
        
            Case eGMCommands.SpawnCreature           'SPA
                Call HandleSpawnCreature(UserIndex)
        
            Case eGMCommands.ResetNPCInventory       '/RESETINV
                Call HandleResetNPCInventory(UserIndex)
        
            Case eGMCommands.CleanWorld              '/LIMPIAR
                Call HandleCleanWorld(UserIndex)
        
            Case eGMCommands.ServerMessage           '/RMSG
                Call HandleServerMessage(UserIndex)
        
            Case eGMCommands.MapMessage              '/MAPMSG
                Call HandleMapMessage(UserIndex)
            
            Case eGMCommands.NickToIP                '/NICK2IP
                Call HandleNickToIP(UserIndex)
        
            Case eGMCommands.IpToNick                '/IP2NICK
                Call HandleIPToNick(UserIndex)
        
            Case eGMCommands.TeleportCreate          '/CT
                Call HandleTeleportCreate(UserIndex)
        
            Case eGMCommands.TeleportDestroy         '/DT
                Call HandleTeleportDestroy(UserIndex)
        
            Case eGMCommands.ForceMIDIToMap          '/FORCEMIDIMAP
                Call HanldeForceMIDIToMap(UserIndex)
        
            Case eGMCommands.ForceWAVEToMap          '/FORCEWAVMAP
                Call HandleForceWAVEToMap(UserIndex)
        
            Case eGMCommands.RoyalArmyMessage        '/REALMSG
                Call HandleRoyalArmyMessage(UserIndex)
        
            Case eGMCommands.ChaosLegionMessage      '/CAOSMSG
                Call HandleChaosLegionMessage(UserIndex)
        
            Case eGMCommands.TalkAsNPC               '/TALKAS
                Call HandleTalkAsNPC(UserIndex)
        
            Case eGMCommands.DestroyAllItemsInArea   '/MASSDEST
                Call HandleDestroyAllItemsInArea(UserIndex)
        
            Case eGMCommands.AcceptRoyalCouncilMember '/ACEPTCONSE
                Call HandleAcceptRoyalCouncilMember(UserIndex)
        
            Case eGMCommands.AcceptChaosCouncilMember '/ACEPTCONSECAOS
                Call HandleAcceptChaosCouncilMember(UserIndex)
        
            Case eGMCommands.ItemsInTheFloor         '/PISO
                Call HandleItemsInTheFloor(UserIndex)
        
            Case eGMCommands.CouncilKick             '/KICKCONSE
                Call HandleCouncilKick(UserIndex)
        
            Case eGMCommands.SetTrigger              '/TRIGGER
                Call HandleSetTrigger(UserIndex)
        
            Case eGMCommands.AskTrigger              '/TRIGGER with no args
                Call HandleAskTrigger(UserIndex)
        
            Case eGMCommands.BannedIPList            '/BANIPLIST
                Call HandleBannedIPList(UserIndex)
        
            Case eGMCommands.BannedIPReload          '/BANIPRELOAD
                Call HandleBannedIPReload(UserIndex)
        
            Case eGMCommands.BanIP                   '/BANIP
                Call HandleBanIP(UserIndex)
        
            Case eGMCommands.UnbanIP                 '/UNBANIP
                Call HandleUnbanIP(UserIndex)
        
            Case eGMCommands.CreateItem              '/CI
                Call HandleCreateItem(UserIndex)
        
            Case eGMCommands.DestroyItems            '/DEST
                Call HandleDestroyItems(UserIndex)
        
            Case eGMCommands.ChaosLegionKick         '/NOCAOS
                Call HandleChaosLegionKick(UserIndex)
        
            Case eGMCommands.RoyalArmyKick           '/NOREAL
                Call HandleRoyalArmyKick(UserIndex)
        
            Case eGMCommands.ForceMIDIAll            '/FORCEMIDI
                Call HandleForceMIDIAll(UserIndex)
        
            Case eGMCommands.ForceWAVEAll            '/FORCEWAV
                Call HandleForceWAVEAll(UserIndex)
        
            Case eGMCommands.TileBlockedToggle       '/BLOQ
                Call HandleTileBlockedToggle(UserIndex)
        
            Case eGMCommands.KillNPCNoRespawn        '/MATA
                Call HandleKillNPCNoRespawn(UserIndex)
        
            Case eGMCommands.KillAllNearbyNPCs       '/MASSKILL
                Call HandleKillAllNearbyNPCs(UserIndex)
        
            Case eGMCommands.LastIP                  '/LASTIP
                Call HandleLastIP(UserIndex)
        
            Case eGMCommands.SystemMessage           '/SMSG
                Call HandleSystemMessage(UserIndex)
        
            Case eGMCommands.CreateNPC               '/ACC
                Call HandleCreateNPC(UserIndex)
        
            Case eGMCommands.CreateNPCWithRespawn    '/RACC
                Call HandleCreateNPCWithRespawn(UserIndex)
        
            Case eGMCommands.ServerOpenToUsersToggle '/HABILITAR
                Call HandleServerOpenToUsersToggle(UserIndex)
        
            Case eGMCommands.TurnOffServer           '/APAGAR
                Call HandleTurnOffServer(UserIndex)
        
            Case eGMCommands.TurnCriminal            '/CONDEN
                Call HandleTurnCriminal(UserIndex)
        
            Case eGMCommands.ResetFactions           '/RAJAR
                Call HandleResetFactions(UserIndex)
        
            Case Declaraciones.eGMCommands.DoBackUp               '/DOBACKUP
                Call HandleDoBackUp(UserIndex)
        
            Case eGMCommands.SaveMap                 '/GUARDAMAPA
                Call HandleSaveMap(UserIndex)
        
            Case eGMCommands.ChangeMapInfoPK         '/MODMAPINFO PK
                Call HandleChangeMapInfoPK(UserIndex)
            
            Case eGMCommands.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
                Call HandleChangeMapInfoBackup(UserIndex)
        
            Case eGMCommands.ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
                Call HandleChangeMapInfoRestricted(UserIndex)
            
            Case eGMCommands.ChangeMapInfoLvl
                Call HandleChangeMapInfoLvl(UserIndex)
            
            Case eGMCommands.ChangeMapInfoLimpieza
                Call HandleChangeMapInfoLimpieza(UserIndex)
            
            Case eGMCommands.ChangeMapInfoItems
                Call HandleChangeMapInfoItems(UserIndex)

            Case eGMCommands.ChangeMapExp
                Call HandleChangeMapInfoExp(UserIndex)
                    
            Case eGMCommands.ChangeMapInfoAttack
                Call HandleChangeMapInfoAttack(UserIndex)
                    
            Case eGMCommands.ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
                Call HandleChangeMapInfoNoMagic(UserIndex)
        
            Case eGMCommands.ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
                Call HandleChangeMapInfoNoInvi(UserIndex)
        
            Case eGMCommands.ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
                Call HandleChangeMapInfoNoResu(UserIndex)
        
            Case eGMCommands.ChangeMapInfoLand       '/MODMAPINFO TERRENO
                Call HandleChangeMapInfoLand(UserIndex)
        
            Case eGMCommands.ChangeMapInfoZone       '/MODMAPINFO ZONA
                Call HandleChangeMapInfoZone(UserIndex)
        
            Case eGMCommands.ChangeMapInfoStealNpc   '/MODMAPINFO ROBONPC
                Call HandleChangeMapInfoStealNpc(UserIndex)
            
            Case eGMCommands.ChangeMapInfoNoOcultar  '/MODMAPINFO OCULTARSINEFECTO
                Call HandleChangeMapInfoNoOcultar(UserIndex)
            
            Case eGMCommands.ChangeMapInfoNoInvocar  '/MODMAPINFO INVOCARSINEFECTO
                Call HandleChangeMapInfoNoInvocar(UserIndex)
            
            Case eGMCommands.SaveChars               '/GRABAR
                Call HandleSaveChars(UserIndex)
        
            Case eGMCommands.ChatColor               '/CHATCOLOR
                Call HandleChatColor(UserIndex)
        
            Case eGMCommands.Ignored                 '/IGNORADO
                Call HandleIgnored(UserIndex)
            
            Case eGMCommands.CreatePretorianClan     '/CREARPRETORIANOS
                Call HandleCreatePretorianClan(UserIndex)
         
            Case eGMCommands.RemovePretorianClan     '/ELIMINARPRETORIANOS
                Call HandleDeletePretorianClan(UserIndex)
                
            Case eGMCommands.EnableDenounces         '/DENUNCIAS
                Call HandleEnableDenounces(UserIndex)
            
            Case eGMCommands.ShowDenouncesList       '/SHOW DENUNCIAS
                Call HandleShowDenouncesList(UserIndex)
        
            Case eGMCommands.SetDialog               '/SETDIALOG
                Call HandleSetDialog(UserIndex)
            
            Case eGMCommands.Impersonate             '/IMPERSONAR
                Call HandleImpersonate(UserIndex)
            
            Case eGMCommands.Imitate                 '/MIMETIZAR
                Call HandleImitate(UserIndex)
            
            Case eGMCommands.RecordAdd
                Call HandleRecordAdd(UserIndex)
            
            Case eGMCommands.RecordAddObs
                Call HandleRecordAddObs(UserIndex)
            
            Case eGMCommands.RecordRemove
                Call HandleRecordRemove(UserIndex)
            
            Case eGMCommands.RecordListRequest
                Call HandleRecordListRequest(UserIndex)
            
            Case eGMCommands.RecordDetailsRequest
                Call HandleRecordDetailsRequest(UserIndex)
            
            Case eGMCommands.SearchObj
                Call HandleSearchObj(UserIndex)
            
            Case eGMCommands.SolicitaSeguridad
                Call HandleSolicitaSeguridad(UserIndex)
            
            Case eGMCommands.CheckingGlobal
                Call HandleCheckingGlobal(UserIndex)
            
            Case eGMCommands.CountDown
                Call HandleCountDown(UserIndex)
            
            Case eGMCommands.GiveBackUser
                Call HandleGiveBackUser(UserIndex)
            
            Case eGMCommands.Pro_Seguimiento
                Call HandlePro_Seguimiento(UserIndex)
            
            Case eGMCommands.Events_KickUser
                Call HandleEvents_KickUser(UserIndex)
            
            Case eGMCommands.SendDataUser
                Call HandleSendDataUser(UserIndex)
                
            Case eGMCommands.SearchDataUser
                Call HandleSearchDataUser(UserIndex)
                
            Case eGMCommands.ChangeModoArgentum
                Call HandleChangeModoArgentum(UserIndex)
                        
            Case eGMCommands.StreamerBotSetting
                Call HandleStreamerBotSetting(UserIndex)
                    
            Case eGMCommands.LotteryNew
                Call HandleLotteryNew(UserIndex)

        End Select

    End With

    Exit Sub

    Call LogError("Error en GmCommands. Error: " & Err.number & " - " & Err.description & ". Paquete: " & Command)

    '<EhFooter>
    Exit Sub

HandleGMCommands_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGMCommands " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Talk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleTalk_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010
    '15/07/2009: ZaMa - Now invisible admins talk by console.
    '23/09/2009: ZaMa - Now invisible admins can't send empty chat.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '***************************************************

    With UserList(UserIndex)
        
        Dim chat          As String

        Dim ValidChat     As Boolean

        Dim PacketCounter As Long

        Dim Packet_ID     As Long
         
        ValidChat = True
        
        chat = Reader.ReadString16()
             
        PacketCounter = Reader.ReadInt32
        Packet_ID = PacketNames.Talk
            
        Call verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "Talk", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID))
                      
        ValidChat = Interval_Message(UserIndex)
                
        If Len(chat) >= 300 Then Exit Sub
             
        If EsGm(UserIndex) Then
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eChat, "Dijo: " & chat)

        End If
        
        Call CheckingOcultation(UserIndex)
        
        If .flags.SlotEvent > 0 Then
            If Events(.flags.SlotEvent).Modality = eModalityEvent.DeathMatch Then Exit Sub

        End If
        
        If .flags.Silenciado = 1 Then
            ValidChat = False

        End If
        
        If Not PalabraPermitida(LCase$(chat)) Then
            Call WriteConsoleMsg(UserIndex, "Según la detección automática de insultos, podrías haber insultado. Recuerda que si te sacan una 'FotoDenuncia' irás a la carcel y depende la gravedad podrás recibir baneo completo de cuenta.", FontTypeNames.FONTTYPE_GMMSG)

        End If
        
        If LenB(chat) <> 0 And ValidChat Then
            
            If Not (.flags.AdminInvisible = 1) Then
                If .flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.charindex, CHAT_COLOR_DEAD_CHAR))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatPersonalizado(chat, .Char.charindex, 1))

                End If
                    
                If Len(chat) >= 3 Then
                    Call WriteAnalyzeText(.Name, chat)

                End If
                    
            Else

                If RTrim(chat) <> "" Then
                    Call SendData(SendTarget.ToGM, UserIndex, PrepareMessageConsoleMsg("Gm '" & .Name & "'> " & chat, FontTypeNames.FONTTYPE_GM))

                End If

            End If

        End If
        
    End With

    '<EhFooter>
    Exit Sub

HandleTalk_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleTalk " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Yell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleYell(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleYell_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010 (ZaMa)
    '15/07/2009: ZaMa - Now invisible admins yell by console.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '***************************************************

    With UserList(UserIndex)
        
        Dim chat      As String

        Dim UserKey   As Integer

        Dim ValidChat As Boolean

        ValidChat = True
        
        chat = Reader.ReadString16()

        ValidChat = Interval_Message(UserIndex)

        If EsGm(UserIndex) Then
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eChat, "Grito: " & chat)

        End If
            
        Call CheckingOcultation(UserIndex)

        If .flags.SlotEvent > 0 Then
            If Events(.flags.SlotEvent).Modality = eModalityEvent.DeathMatch Then Exit Sub

        End If
        
        If .flags.Silenciado = 1 Then
            ValidChat = False

        End If
        
        If ValidChat Then
            If .flags.Privilegios And PlayerType.User Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.charindex, CHAT_COLOR_DEAD_CHAR))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatPersonalizado(chat, .Char.charindex, 4))

                End If

            Else

                If Not (.flags.AdminInvisible = 1) Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.charindex, CHAT_COLOR_GM_YELL))
                Else
                    Call SendData(SendTarget.ToGM, UserIndex, PrepareMessageConsoleMsg("Gms> " & chat, FontTypeNames.FONTTYPE_GM))

                End If

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleYell_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleYell " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Whisper" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleWhisper_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 03/12/2010
    '28/05/2009: ZaMa - Now it doesn't appear any message when private talking to an invisible admin
    '15/07/2009: ZaMa - Now invisible admins wisper by console.
    '03/12/2010: Enanoh - Agregué susurro a Admins en modo consulta y Los Dioses pueden susurrar en ciertos casos.
    '***************************************************

    With UserList(UserIndex)
        
        Dim chat            As String

        Dim targetUserIndex As Integer

        Dim TargetPriv      As PlayerType

        Dim UserPriv        As PlayerType

        Dim TargetName      As String

        Dim ValidChat       As Boolean

        ValidChat = True
        
        TargetName = Reader.ReadString8()
        chat = Reader.ReadString16()
        
        ValidChat = Interval_Message(UserIndex)
        
        UserPriv = .flags.Privilegios
        
        If .flags.SlotEvent > 0 Then Exit Sub
                
        If .flags.Muerto Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. ", FontTypeNames.FONTTYPE_INFO)
        Else
            ' Offline?
            targetUserIndex = NameIndex(TargetName)

            If targetUserIndex = 0 Then

                ' Admin?
                If EsGmChar(TargetName) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                    ' Whisperer admin? (Else say nothing)
                ElseIf (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If
                
                ' Online
            Else
                ' Privilegios
                TargetPriv = UserList(targetUserIndex).flags.Privilegios
                
                ' Semis y usuarios no pueden susurrar a dioses (Salvo en consulta)
                If (TargetPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 And (UserPriv And (PlayerType.User Or PlayerType.SemiDios)) <> 0 And Not .flags.EnConsulta Then
                    
                    ' No puede
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)

                    ' Usuarios no pueden susurrar a semis o conses (Salvo en consulta)
                ElseIf (UserPriv And PlayerType.User) <> 0 And (Not TargetPriv And PlayerType.User) <> 0 And Not .flags.EnConsulta Then
                    
                    ' No puede
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                
                    ' En rango? (Los dioses pueden susurrar a distancia)
                ElseIf Not EstaPCarea(UserIndex, targetUserIndex) And (UserPriv And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios)) = 0 Then
                    
                    ' No se puede susurrar a admins fuera de su rango
                    If (TargetPriv And (PlayerType.User)) = 0 And (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) = 0 Then
                        Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                    
                        ' Whisperer admin? (Else say nothing)
                    ElseIf (UserPriv And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios)) <> 0 Then
                        Call WriteConsoleMsg(UserIndex, "Estás muy lejos del usuario.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else

                    '[GMs]
                    If UserPriv And (PlayerType.SemiDios) Then
                        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eChat, "Le susurro a '" & UserList(targetUserIndex).Name & "' " & chat)
                        
                        ' Usuarios a administradores
                    ElseIf (UserPriv And PlayerType.User) <> 0 And (TargetPriv And PlayerType.User) = 0 Then
                        Call Logs_User(UserList(targetUserIndex).Name, eLog.eGm, eLogDescUser.eChat, .Name & " le susurro en consulta: " & chat)

                    End If

                    If .flags.SlotEvent > 0 Then
                        If Events(.flags.SlotEvent).Modality = eModalityEvent.DeathMatch Then ValidChat = False

                    End If
        
                    If .flags.Silenciado = 1 Then
                        ValidChat = False

                    End If
                    
                    If LenB(chat) <> 0 And ValidChat Then

                        ' Dios susurrando a distancia
                        If Not EstaPCarea(UserIndex, targetUserIndex) And (UserPriv And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios)) <> 0 Then
                            
                            Call WriteConsoleMsg(UserIndex, "Susurraste> " & chat, FontTypeNames.FONTTYPE_GM)

                            Call WriteConsoleMsg(targetUserIndex, "Gm susurra> " & chat, FontTypeNames.FONTTYPE_GM)
                            
                        ElseIf Not (.flags.AdminInvisible = 1) Then
                            Call WriteChatPersonalizado(UserIndex, chat, .Char.charindex, 6)
                            Call WriteChatPersonalizado(targetUserIndex, chat, .Char.charindex, 6)
                            Call FlushBuffer(targetUserIndex)
                            
                            '[CDT 17-02-2004]
                            If .flags.Privilegios And (PlayerType.User) Then
                                Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageChatOverHead("A " & UserList(targetUserIndex).Name & "> " & chat, .Char.charindex, vbYellow))

                            End If

                        Else
                            Call WriteConsoleMsg(UserIndex, "Susurraste> " & chat, FontTypeNames.FONTTYPE_GM)

                            If UserIndex <> targetUserIndex Then Call WriteConsoleMsg(targetUserIndex, "Gm susurra> " & chat, FontTypeNames.FONTTYPE_GM)
                            
                            If .flags.Privilegios And (PlayerType.User) Then
                                Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageConsoleMsg("Gm dijo a " & UserList(targetUserIndex).Name & "> " & chat, FontTypeNames.FONTTYPE_GM))

                            End If

                        End If

                    End If

                End If

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleWhisper_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleWhisper " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Walk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010 (ZaMa)
    '11/19/09 Pato - Now the class bandit can walk hidden.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '***************************************************
    
    Dim dummy       As Long

    Dim TempTick    As Long

    Dim Heading     As eHeading
    
    Dim MaxTimeWalk As Integer
        
    Dim PacketCount As Long
        
    With UserList(UserIndex)
        
        Heading = Reader.ReadInt()
        PacketCount = Reader.ReadInt32
            
        Call verifyTimeStamp(PacketCount, .PacketCounters(PacketNames.Walk), .PacketTimers(PacketNames.Walk), .MacroIterations(PacketNames.Walk), UserIndex, "Walk", PacketTimerThreshold(PacketNames.Walk), MacroIterations(PacketNames.Walk))
            
        If .flags.Muerto Then
            MaxTimeWalk = 36
        Else
            MaxTimeWalk = 30

        End If
        
        If .flags.Paralizado = 0 Then
            
            If .flags.Meditando Then
                
                ' Probabilidad de subir un % de maná al moverse
                If RandomNumber(1, 100) <= 20 Then

                    Dim Mana As Long

                    Mana = Porcentaje(.Stats.MaxMan, Porcentaje(Balance.PorcentajeRecuperoMana, 50 + .Stats.UserSkills(eSkill.Magia) * 0.5))

                    If Mana <= 0 Then Mana = 1
                    
                    If .Stats.MinMan + Mana >= .Stats.MaxMan Then
                        .Stats.MinMan = .Stats.MaxMan
                    Else
                        .Stats.MinMan = .Stats.MinMan + Mana

                    End If
                    
                    Call WriteUpdateMana(UserIndex)

                End If
                
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                UserList(UserIndex).Char.FX = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.charindex, 0))

            End If
            
            Dim CurrentTick As Long

            CurrentTick = GetTime
        
            'Prevent SpeedHack (refactored by WyroX)
            If .Char.speeding > 0 Then

                Dim ElapsedTimeStep As Long, MinTimeStep As Long, DeltaStep As Single

                ElapsedTimeStep = CurrentTick - .Counters.LastStep
                MinTimeStep = IntervaloCaminar / .Char.speeding
                DeltaStep = (MinTimeStep - ElapsedTimeStep) / MinTimeStep

                If DeltaStep > 0 Then
                
                    .Counters.SpeedHackCounter = .Counters.SpeedHackCounter + DeltaStep
                
                    If .Counters.SpeedHackCounter > MaximoSpeedHack Then
                        'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Administración Â» Posible uso de SpeedHack del usuario " & .name & ".", e_FontTypeNames.FONTTYPE_SERVER))
                        Call WritePosUpdate(UserIndex)
                        Exit Sub

                    End If

                Else
                
                    .Counters.SpeedHackCounter = .Counters.SpeedHackCounter + DeltaStep * 5

                    If .Counters.SpeedHackCounter < 0 Then .Counters.SpeedHackCounter = 0

                End If

            End If
            
            ' @ En la daga rusa no te podes mover [El chequeo está en el cliente, pero al empezar se mueven como retrasados]
            If .flags.SlotEvent > 0 Then
                If Events(.flags.SlotEvent).Modality = eModalityEvent.DagaRusa Then
                    If Events(.flags.SlotEvent).Run Then Exit Sub

                End If

            End If
            
            'Move user
            If MoveUserChar(UserIndex, Heading) Then
                ' Save current step for anti-sh
                .Counters.LastStep = CurrentTick
                         
                If UserIndex <> StreamerBot.Active And StreamerBot.Active > 0 Then
                    ' If StrComp(StreamerBot.LastTarget, UCase$(.Name)) = 0 Then
                       
                    'End If
                    
                    ' If StrComp(StreamerBot.LastTarget, UCase$(.Name)) = 0 Then
                    ' If MoveUserChar(StreamerBot.Active, Heading) Then
                    '   Call WriteForceCharMove(StreamerBot.Active, Heading)
                    'End If
                    'End If

                End If

                'Stop resting if needed
                If .flags.Descansar Then
                    .flags.Descansar = False
                    
                    Call WriteRestOK(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)

                End If

                'If exiting, cancel
                Call CancelExit(UserIndex)
                
                'Esta usando el /HOGAR, no se puede mover
                If .flags.Traveling = 1 Then
                    .flags.Traveling = 0
                    .Counters.goHome = 0
                    Call WriteConsoleMsg(UserIndex, "Has cancelado el viaje a casa.", FontTypeNames.FONTTYPE_INFO)

                End If
            
            Else
                .Counters.LastStep = 0
                Call WritePosUpdate(UserIndex)

            End If
            
        Else    'paralized

            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                
                Call WriteConsoleMsg(UserIndex, "No puedes moverte porque estás paralizado.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'Can't move while hidden except he is a thief
        If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
                            
            Dim HunterInPosValid As Boolean
                
            ' @ Cazadores y sus capuchas weonas
            '  If .Clase = eClass.Hunter And .Stats.UserSkills(eSkill.Ocultarse) > 90 Then
            If .Invent.CascoEqpObjIndex > 0 Then
                If ObjData(.Invent.CascoEqpObjIndex).Oculto > 0 Then

                    ' Si está en el rango permitido desde que se ocultó, puede moverse libre.
                    ' Esta dentro del rango permitido
                    If Distance(.Pos.X, .Pos.Y, .PosOculto.X, .PosOculto.Y) <= ObjData(.Invent.CascoEqpObjIndex).Oculto Then
                        HunterInPosValid = True

                    End If
                
                End If

            End If

            ' End If
                
            If .Clase <> eClass.Thief And Not HunterInPosValid Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
            
                If .flags.Navegando = 0 Then

                    'If not under a spell effect, show char
                    If .flags.Invisible = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                        Call UsUaRiOs.SetInvisible(UserIndex, .Char.charindex, False)

                    End If

                End If

            End If

        End If
        
        .Counters.PiqueteC = 0
        Call Guilds_UpdatePosition(UserIndex)

    End With

End Sub

Public Function Check_UserBlocked(ByVal Map As Integer, _
                                  ByVal X As Integer, _
                                  ByVal Y As Integer) As Boolean

    On Error GoTo ErrHandler
    
    If MapData(Map, X - 1, Y).Blocked = 0 And MapData(Map, X - 1, Y).NpcIndex = 0 And MapData(Map, X - 1, Y).UserIndex = 0 And MapData(Map, X - 1, Y).TileExit.Map = 0 Then
            
        Check_UserBlocked = False
        Exit Function

    End If
    
    If MapData(Map, X + 1, Y).Blocked = 0 And MapData(Map, X + 1, Y).NpcIndex = 0 And MapData(Map, X + 1, Y).UserIndex = 0 And MapData(Map, X + 1, Y).TileExit.Map = 0 Then
            
        Check_UserBlocked = False
        Exit Function

    End If
    
    If MapData(Map, X, Y - 1).Blocked = 0 And MapData(Map, X, Y - 1).NpcIndex = 0 And MapData(Map, X, Y - 1).UserIndex = 0 And MapData(Map, X, Y - 1).TileExit.Map = 0 Then
        Check_UserBlocked = False
        Exit Function

    End If
    
    If MapData(Map, X, Y + 1).Blocked = 0 And MapData(Map, X, Y + 1).NpcIndex = 0 And MapData(Map, X, Y + 1).UserIndex = 0 And MapData(Map, X, Y + 1).TileExit.Map = 0 Then
            
        Check_UserBlocked = False
        Exit Function

    End If
    
    Check_UserBlocked = True
    Exit Function
ErrHandler:
    
End Function

''
' Handles the "RequestPositionUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRequestPositionUpdate_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 29/10/2021
    '
    '***************************************************
    
    Dim Pos    As WorldPos

    Dim OldPos As WorldPos
    
    With UserList(UserIndex)
        Pos = .Pos
        
        If .flags.SlotReto = 0 And .flags.SlotEvent = 0 And .flags.SlotFast = 0 And .flags.Desafiando = 0 And Not MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = ZONAPELEA Then
            
            If Check_UserBlocked(Pos.Map, Pos.X, Pos.Y) Then
                Call ClosestStablePos(Pos, Pos)

                If Pos.X <> 0 And .Pos.Y <> 0 Then
                    Call WarpUserChar(UserIndex, Pos.Map, Pos.X, Pos.Y, True)

                End If

            End If

        End If
        
    End With
    
    Call WritePosUpdate(UserIndex)
    '<EhFooter>
    Exit Sub

HandleRequestPositionUpdate_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRequestPositionUpdate " & "at line " & Erl

    '</EhFooter>
End Sub

''
' Handles the "Attack" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAttack(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleAttack_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010
    'Last Modified By: ZaMa
    '10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo.
    '13/11/2009: ZaMa - Se cancela el estado no atacable al atcar.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '***************************************************
    
    With UserList(UserIndex)
          
        Dim PacketCounter As Long

        PacketCounter = Reader.ReadInt32
                        
        Dim Packet_ID As Long

        Packet_ID = PacketNames.Attack
            
        Call verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "Attack", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID))
            
        If .flags.GmSeguidor > 0 Then

            Dim Temp As Long, TiempoActual As Long

            TiempoActual = GetTime
            Temp = TiempoActual - .interval(0).IAttack
                    
            Call WriteUpdateInfoIntervals(.flags.GmSeguidor, 4, Temp, .flags.MenuCliente)
                    
            .interval(0).IAttack = TiempoActual

        End If
        
        'If dead, can't attack
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        'If user meditates, can't attack
        If .flags.Meditando Then

            Exit Sub

        End If
        
        'If equiped weapon is ranged, can't attack this way
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes usar así este arma.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If

        End If
        
        If UserList(UserIndex).flags.Meditando Then
            UserList(UserIndex).flags.Meditando = False
            UserList(UserIndex).Char.FX = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.charindex, 0))

        End If
            
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Attack!
        Call UsuarioAtaca(UserIndex)
        
        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False
        
        'I see you...
        If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 0 Then

                If .flags.Invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.charindex, False)
                    Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
     
    End With

    '<EhFooter>
    Exit Sub

HandleAttack_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleAttack " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "PickUp" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePickUp(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandlePickUp_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 07/25/09
    '02/26/2006: Marco - Agregué un checkeo por si el usuario trata de agarrar un item mientras comercia.
    '***************************************************
    
    With UserList(UserIndex)
        
        'If dead, it can't pick up objects
        If .flags.Muerto = 1 Then Exit Sub
        
        'If user is trading items and attempts to pickup an item, he's cheating, so we kick him.
        If .flags.Comerciando Then Exit Sub
        
        Call GetObj(UserIndex)

    End With

    '<EhFooter>
    Exit Sub

HandlePickUp_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandlePickUp " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "SafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSafeToggle(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleSafeToggle_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
            
        If .Faction.Status = r_Armada Then
            Call WriteConsoleMsg(UserIndex, "Tu facción no te permite quitar el seguro. Por favor dirigete al Rey de Banderbill y abandona la facción.", FONTTYPE_WARNING)
            Exit Sub

        End If
            
        If .flags.Seguro Then
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOff) 'Call WriteSafeModeOff(UserIndex)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn) 'Call WriteSafeModeOn(UserIndex)

        End If
        
        .flags.Seguro = Not .flags.Seguro

    End With

    '<EhFooter>
    Exit Sub

HandleSafeToggle_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleSafeToggle " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ResuscitationSafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResuscitationToggle(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleResuscitationToggle_Err

    '</EhHeader>

    '***************************************************
    'Author: Rapsodius
    'Creation Date: 10/10/07
    '***************************************************
    With UserList(UserIndex)
        
        .flags.SeguroResu = Not .flags.SeguroResu
        
        If .flags.SeguroResu Then
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn) 'Call WriteResuscitationSafeOn(UserIndex)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff) 'Call WriteResuscitationSafeOff(UserIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleResuscitationToggle_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleResuscitationToggle " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleDragToggle(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleDragToggle_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        .flags.DragBlocked = Not .flags.DragBlocked
        
        If .flags.DragBlocked Then
            Call WriteMultiMessage(UserIndex, eMessages.DragSafeOn) 'Call WriteResuscitationSafeOn(UserIndex)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.DragSafeOff) 'Call WriteResuscitationSafeOff(UserIndex)

        End If
    
    End With

    '<EhFooter>
    Exit Sub

HandleDragToggle_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleDragToggle " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAtributes(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRequestAtributes_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call WriteAttributes(UserIndex)
    '<EhFooter>
    Exit Sub

HandleRequestAtributes_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRequestAtributes " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRequestSkills_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call WriteSendSkills(UserIndex)
    '<EhFooter>
    Exit Sub

HandleRequestSkills_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRequestSkills " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMiniStats(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRequestMiniStats_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call WriteMiniStats(UserIndex, UserIndex)
    '<EhFooter>
    Exit Sub

HandleRequestMiniStats_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRequestMiniStats " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleCommerceEnd_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    'User quits commerce mode
    UserList(UserIndex).flags.Comerciando = False
    Call WriteCommerceEnd(UserIndex)
    '<EhFooter>
    Exit Sub

HandleCommerceEnd_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCommerceEnd " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleUserCommerceEnd_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 11/03/2010
    '11/03/2010: ZaMa - Le avisa por consola al que cencela que dejo de comerciar.
    '***************************************************
    With UserList(UserIndex)
        
        'Quits commerce mode with user
        If .ComUsu.DestUsu > 0 Then
            If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call WriteConsoleMsg(.ComUsu.DestUsu, .Name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(.ComUsu.DestUsu)
                
                'Send data in the outgoing buffer of the other user
                Call FlushBuffer(.ComUsu.DestUsu)

            End If

        End If
        
        Call FinComerciarUsu(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has dejado de comerciar.", FontTypeNames.FONTTYPE_TALK)

    End With

    '<EhFooter>
    Exit Sub

HandleUserCommerceEnd_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleUserCommerceEnd " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "UserCommerceConfirm" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUserCommerceConfirm(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleUserCommerceConfirm_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    '
    '***************************************************

    'Validate the commerce
    If PuedeSeguirComerciando(UserIndex) Then
        'Tell the other user the confirmation of the offer
        Call WriteUserOfferConfirm(UserList(UserIndex).ComUsu.DestUsu)
        UserList(UserIndex).ComUsu.Confirmo = True

    End If
    
    '<EhFooter>
    Exit Sub

HandleUserCommerceConfirm_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleUserCommerceConfirm " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleCommerceChat(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleCommerceChat_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 03/12/2009
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim chat As String
        
        chat = Reader.ReadString8()
        
        If LenB(chat) <> 0 Then
            If PuedeSeguirComerciando(UserIndex) Then
                
                chat = UserList(UserIndex).Name & "> " & chat
                Call WriteCommerceChat(UserIndex, chat, FontTypeNames.FONTTYPE_PARTY)
                Call WriteCommerceChat(UserList(UserIndex).ComUsu.DestUsu, chat, FontTypeNames.FONTTYPE_PARTY)

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleCommerceChat_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCommerceChat " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "BankEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleBankEnd_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        'User exits banking mode
        .flags.Comerciando = False
        Call WriteBankEnd(UserIndex)

    End With

    '<EhFooter>
    Exit Sub

HandleBankEnd_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleBankEnd " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOk(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleUserCommerceOk_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    'Trade accepted
    Call AceptarComercioUsu(UserIndex)
    '<EhFooter>
    Exit Sub

HandleUserCommerceOk_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleUserCommerceOk " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleUserCommerceReject_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim otherUser As Integer
    
    With UserList(UserIndex)
        
        otherUser = .ComUsu.DestUsu
        
        'Offer rejected
        If otherUser > 0 Then
            If UserList(otherUser).flags.UserLogged Then
                Call WriteConsoleMsg(otherUser, .Name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(otherUser)
                
                'Send data in the outgoing buffer of the other user
                Call FlushBuffer(otherUser)

            End If

        End If
        
        Call WriteConsoleMsg(UserIndex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
        Call FinComerciarUsu(UserIndex)

    End With

    '<EhFooter>
    Exit Sub

HandleUserCommerceReject_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleUserCommerceReject " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Drop" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleDrop_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 07/25/09
    '07/25/09: Marco - Agregué un checkeo para patear a los usuarios que tiran items mientras comercian.
    '***************************************************
    
    Dim Slot   As Byte

    Dim Amount As Integer
    
    With UserList(UserIndex)

        Slot = Reader.ReadInt()
        Amount = Reader.ReadInt()

        If Not Interval_Drop(UserIndex) Then Exit Sub
        
        'low rank admins can't drop item. Neither can the dead nor those sailing.
        If .flags.Navegando = 1 Or .flags.Muerto = 1 Or .flags.Montando = 1 Or .flags.SlotEvent > 0 Or .flags.SlotReto > 0 Then Exit Sub

        'If the user is trading, he can't drop items => He's cheating, we kick him.
        If .flags.Comerciando Then Exit Sub
        
        If Slot = FLAGORO + 1 Then

            Exit Sub

            'If Amount > 10000 Then Exit Sub 'Don't drop too much gold
            'If (.Stats.Eldhir - Amount) < 0 Then Exit Sub
            
            'Dim Pos As WorldPos
            'Dim Obj As Obj
            
            'Obj.ObjIndex = 1246
            'Obj.Amount = Amount
            
            'TirarItemAlPiso .Pos, Obj
            ' .Stats.Eldhir = .Stats.Eldhir - Amount
            'Call WriteUpdateDsp(UserIndex)
            
            'Are we dropping gold or other items??
        ElseIf Slot = FLAGORO Then

            If Amount > 10000 Then Exit Sub 'Don't drop too much gold
            If (.Stats.Gld - Amount) < 0 Then Exit Sub
            
            Dim Pos As WorldPos

            Dim Obj As Obj
            
            Obj.ObjIndex = iORO
            Obj.Amount = Amount
            
            TirarItemAlPiso .Pos, Obj
            .Stats.Gld = .Stats.Gld - Amount
            Call WriteUpdateGold(UserIndex)
            Exit Sub
        Else

            'Only drop valid slots
            If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
                If .Invent.Object(Slot).ObjIndex = 0 Then

                    Exit Sub

                End If
                
                Call DropObj(UserIndex, Slot, Amount, .Pos.Map, .Pos.X, .Pos.Y)

            End If

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleDrop_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleDrop " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "CastSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCastSpell(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleCastSpell_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '13/11/2009: ZaMa - Ahora los npcs pueden atacar al usuario si quizo castear un hechizo
    '***************************************************
    With UserList(UserIndex)
        
        Dim Spell As Byte
        
        Spell = Reader.ReadInt()
        Reader.ReadInt16
        Reader.ReadInt8
        
        If Not IntervaloPermiteCastear(UserIndex, True) Then Exit Sub    'Nuevo intervalo de casteo.
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If .flags.MenuCliente <> 255 And .flags.MenuCliente <> 1 Then

            Exit Sub

        End If
        
        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False
        
        If Spell < 1 Then
            .flags.Hechizo = 0

            Exit Sub

        ElseIf Spell > MAXUSERHECHIZOS Then
            .flags.Hechizo = 0

            Exit Sub

        End If
        
        .flags.Hechizo = .Stats.UserHechizos(Spell)

        If Hechizos(.flags.Hechizo).AutoLanzar = 1 Then

            'Check bow's interval
            If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
                
            'Check attack-spell interval
            If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub
                
            'Check Magic interval
            If Not IntervaloPermiteLanzarSpell(UserIndex) Then Exit Sub

            .Counters.controlHechizos.HechizosTotales = .Counters.controlHechizos.HechizosTotales + 1
            Call LanzarHechizo(.flags.Hechizo, UserIndex)
            .flags.Hechizo = 0

        End If
            
    End With

    '<EhFooter>
    Exit Sub

HandleCastSpell_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCastSpell " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

''
' Handles the "LeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleLeftClick_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
        
    Dim X As Byte

    Dim Y As Byte
        
    X = Reader.ReadInt()
    Y = Reader.ReadInt()
        
    Dim PacketCounter As Long

    PacketCounter = Reader.ReadInt32
                        
    Dim Packet_ID As Long

    Packet_ID = PacketNames.LeftClick
            
    With UserList(UserIndex)
        Call verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "LeftClick", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID))
            
        Call LookatTile(UserIndex, .Pos.Map, X, Y)
            
    End With
 
    '<EhFooter>
    Exit Sub

HandleLeftClick_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleLeftClick " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "DoubleClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDoubleClick(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleDoubleClick_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
        
    Dim X    As Byte

    Dim Y    As Byte
        
    Dim Tipo As Byte
        
    X = Reader.ReadInt8()
    Y = Reader.ReadInt8()
    Tipo = Reader.ReadInt8
          
    Call Accion(UserIndex, UserList(UserIndex).Pos.Map, X, Y, Tipo)

    '<EhFooter>
    Exit Sub

HandleDoubleClick_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleDoubleClick " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "RightClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRightClick(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRightClick_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 10/05/2011
    '
    '***************************************************
        
    Dim X       As Byte

    Dim Y       As Byte
        
    Dim MouseX  As Long

    Dim MouseY  As Long

    Dim UserKey As Integer
        
    X = Reader.ReadInt8()
    Y = Reader.ReadInt8()
        
    MouseX = Reader.ReadInt32()
    MouseY = Reader.ReadInt32()
            
    Call Extra.ShowMenu(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
    
    '<EhFooter>
    Exit Sub

HandleRightClick_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRightClick " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Work" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWork(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleWork_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010 (ZaMa)
    '13/01/2010: ZaMa - El pirata se puede ocultar en barca
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Skill   As eSkill

        Dim UserKey As Integer
        
        Skill = Reader.ReadInt()
            
        Dim PacketCounter As Long

        PacketCounter = Reader.ReadInt32
                        
        Dim Packet_ID As Long

        Packet_ID = PacketNames.Work
            
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill
        
            Case Robar, Magia, Domar
                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, Skill)
                
            Case Ocultarse
                Call verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "Ocultar", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID))
                    
                ' Verifico si se peude ocultar en este mapa
                If MapInfo(.Pos.Map).OcultarSinEfecto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡Ocultarse no funciona aquí!", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                    
                If .flags.SlotFast > 0 Then
                    If RetoFast(.flags.SlotFast).ConfigVale <> ValeTodo Then
                        Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite el ocultamiento.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                        Exit Sub
    
                    End If

                End If
                
                If .flags.SlotEvent > 0 Then
                    If Events(.flags.SlotEvent).config(eConfigEvent.eOcultar) = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Ocultar no está permitido aquí! Retirate de la Zona del Evento si deseas esconderte entre las sombras.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                End If
                
                If .flags.EnConsulta Then
                    Call WriteConsoleMsg(UserIndex, "No puedes ocultarte si estás en consulta.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                
                'If .Stats.MaxMan > 0 Then
                ' Call WriteConsoleMsg(UserIndex, "No tienes el conocimiento para ocultarte entre las sombras.", FontTypeNames.FONTTYPE_INFO)
                
                '   Exit Sub
                '  End If
                
                If Power.UserIndex = UserIndex Then
                    Call WriteConsoleMsg(UserIndex, "¿A que seguro eres un cazador ah!? ¡Plantate!", FontTypeNames.FONTTYPE_INFO)
                
                    Exit Sub

                End If
                
                If .flags.Navegando = 1 Or .flags.Montando = 1 Or .flags.Mimetizado = 1 Or .flags.Invisible = 1 Then

                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 3 Then
                        Call WriteConsoleMsg(UserIndex, "No puedes ocultarte en este momento .", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 3

                    End If

                    '[/CDT]
                    Exit Sub

                End If
                
                If .flags.Oculto = 1 Then

                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 2 Then
                        Call WriteConsoleMsg(UserIndex, "Ya estás oculto.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 2

                    End If

                    '[/CDT]
                    Exit Sub

                End If
                
                Call DoOcultarse(UserIndex)
                
        End Select
        
    End With

    '<EhFooter>
    Exit Sub

HandleWork_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleWork " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "InitCrafting" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInitCrafting(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleInitCrafting_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/01/2010
    '
    '***************************************************
    
    Dim TotalItems    As Long

    Dim ItemsPorCiclo As Integer
    
    With UserList(UserIndex)
        
        TotalItems = Reader.ReadInt
        ItemsPorCiclo = Reader.ReadInt
        
        If TotalItems > 0 Then
            
            .Construir.cantidad = TotalItems
            .Construir.PorCiclo = MinimoInt(MaxItemsConstruibles(UserIndex), ItemsPorCiclo)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleInitCrafting_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleInitCrafting " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "UseItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleUseItem_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Slot           As Byte

        Dim SecondaryClick As Byte

        Dim Value          As Long

        Dim UserKey        As Integer
        
        Dim Key            As Integer

        Dim PacketCounter  As Long

        Dim Packet_ID      As Long

        Slot = Reader.ReadInt8()
        SecondaryClick = Reader.ReadInt8()
        Value = Reader.ReadInt32()
     
        PacketCounter = Reader.ReadInt32
            
        If Slot <= .CurrentInventorySlots And Slot > 0 Then
            If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
        Else
            Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, .Name & " con IP: " & .IpAddress & " hizo algo raro al usar objetos")
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[ANTI-CHEAT]: " & .Name & " hizo algo raro al usar objetos", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
            Exit Sub

        End If
            
        If .flags.Meditando Then Exit Sub
        If .flags.Comerciando Then Exit Sub
        
        If SecondaryClick And .flags.MenuCliente = 1 Then Exit Sub
        If .flags.LastSlotClient <> 255 And Slot <> .flags.LastSlotClient Then Exit Sub
              
        If SecondaryClick Then
 
            If (GetTime - .TimeUseClicInitial) >= 1000 Then
                If .TimeUseClic >= 5 Then
                    Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, .Name & " con IP: " & .IpAddress & " está utilizando más de 4 doble-clics")
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[ANTI-CHEAT]: " & .Name & " está utilizando más de 4 doble-clics", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))

                End If
                        
                .TimeUseClic = 0
                .TimeUseClicInitial = GetTime
                Exit Sub
            Else
                .TimeUseClic = .TimeUseClic + 1

            End If
                
            Packet_ID = PacketNames.UseItem
            Call verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "UseItem", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID))
        Else
            Packet_ID = PacketNames.UseItemU
            Call verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "UseItemU", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID))

        End If
            
        Call UseInvItem(UserIndex, Slot, SecondaryClick, Value)

    End With
            
    '<EhFooter>
    Exit Sub

HandleUseItem_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleUseItem " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "UseItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseItemTwo(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleUseItemTwo_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Slot           As Byte

        Dim SecondaryClick As Byte

        Dim Value          As Long

        Dim UserKey        As Integer
        
        Dim Key            As Integer

        Slot = Reader.ReadInt8()
        SecondaryClick = Reader.ReadInt8()
        Value = Reader.ReadInt32()

        If Slot <= .CurrentInventorySlots And Slot > 0 Then
            If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
        Else
            Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, .Name & " con IP: " & .IpAddress & " hizo algo raro al usar objetos")
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[ANTI-CHEAT]: " & .Name & " hizo algo raro al usar objetos", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
            Exit Sub

        End If
        
        If .flags.Meditando Then Exit Sub
        If .flags.Comerciando Then Exit Sub
        
        If SecondaryClick And .flags.MenuCliente = 1 Then Exit Sub
        If .flags.LastSlotClient <> 255 And Slot <> .flags.LastSlotClient Then Exit Sub
            
        If PacketUseItem <> ClientPacketID.UseItemTwo Then
            Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, .Name & " con IP: " & .IpAddress & " utilizo un paquete guardado")
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[ANTI-CHEAT]: " & .Name & " utilizo un paquete guardado", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))

        End If
            
        Call UseInvItem(UserIndex, Slot, SecondaryClick, Value)

    End With
            
    '<EhFooter>
    Exit Sub

HandleUseItemTwo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleUseItemTwo " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "CraftBlacksmith" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftBlacksmith(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleCraftBlacksmith_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
        
    Dim QuestIndex As Integer
    
    QuestIndex = Reader.ReadInt16()
    
    If QuestIndex < 1 Or QuestIndex > NumQuests Then Exit Sub
    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

    '<EhFooter>
    Exit Sub

HandleCraftBlacksmith_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCraftBlacksmith " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleWorkLeftClick_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 14/01/2010 (ZaMa)
    '16/11/2009: ZaMa - Agregada la posibilidad de extraer madera elfica.
    '12/01/2010: ZaMa - Ahora se admiten armas arrojadizas (proyectiles sin municiones).
    '14/01/2010: ZaMa - Ya no se pierden municiones al atacar npcs con dueño.
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim X           As Byte

        Dim Y           As Byte

        Dim Skill       As eSkill

        Dim DummyInt    As Integer

        Dim tU          As Integer   'Target user

        Dim tN          As Integer   'Target NPC
        
        Dim WeaponIndex As Integer
        
        Dim Key         As Integer
            
        Dim MouseX      As Long
            
        Dim MouseY      As Long
            
        X = Reader.ReadInt8()
        Y = Reader.ReadInt8()
        
        Skill = Reader.ReadInt8()
        MouseX = Reader.ReadInt8
        MouseY = Reader.ReadInt16
              
        Dim PacketCounter As Long

        PacketCounter = Reader.ReadInt32
                        
        Dim Packet_ID As Long

        Packet_ID = PacketNames.WorkLeftClick
            
        If (.flags.Muerto = 1 And Skill <> TeleportInvoker) Or .flags.Descansar Or .flags.Meditando Or Not InMapBounds(.Pos.Map, X, Y) Then Exit Sub
                
        '  If .Clase <> eClass.Worker Then
        ' UpdatePointer UserIndex, .flags.MenuCliente, X, Y, "Click to Win"

        '  End If
              
        If Not InRangoVision(UserIndex, X, Y) Then
            Call WritePosUpdate(UserIndex)

            Exit Sub

        End If
            
        If .flags.Meditando Then
            .flags.Meditando = False
            .Char.FX = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))

        End If
            
        Call verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "WorkLeftClick", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID))
            
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill

            Case eSkill.Proyectiles
                
                'Check attack interval
                If Not IntervaloPermiteMagiaGolpe(UserIndex, False) Then Exit Sub

                'Check Magic interval
                If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub

                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                
                Call LanzarProyectil(UserIndex, X, Y)
                            
            Case eSkill.Magia

                'Check the map allows spells to be casted.
                If MapInfo(.Pos.Map).MagiaSinEfecto > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Una fuerza oscura te impide canalizar tu energía.", FontTypeNames.FONTTYPE_FIGHT)

                    Exit Sub

                End If
                
                'Target whatever is in that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                'If it's outside range log it and exit
                If Abs(.Pos.X - X) > RANGO_VISION_x Or Abs(.Pos.Y - Y) > RANGO_VISION_y Then
                    Call LogCheating("Ataque fuera de rango de " & .Name & "(" & .Pos.Map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .IpAddress & " a la posición (" & .Pos.Map & "/" & X & "/" & Y & ")")

                    Exit Sub

                End If
                
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
                
                'Check attack-spell interval
                If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub
                
                'Check Magic interval
                If Not IntervaloPermiteLanzarSpell(UserIndex) Then Exit Sub
                 
                'Check intervals and cast
                If .flags.Hechizo > 0 Then
                    If Hechizos(.flags.Hechizo).AutoLanzar = 1 Then Exit Sub ' Anti hack
                    .Counters.controlHechizos.HechizosTotales = .Counters.controlHechizos.HechizosTotales + 1
                    Call LanzarHechizo(.flags.Hechizo, UserIndex)
                    .flags.Hechizo = 0
                Else
                    Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eSkill.Robar

                If .Clase <> eClass.Thief Then
                    Call WriteConsoleMsg(UserIndex, "¡Tu no puedes robar!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                        
                'Does the map allow us to steal here?
                If MapInfo(.Pos.Map).Pk Then
                    
                    'Check interval
                    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                    
                    'Target whatever is in that tile
                    Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
                    tU = .flags.TargetUser
                    
                    If tU > 0 And tU <> UserIndex Then

                        'Can't steal administrative players
                        If UserList(tU).flags.Privilegios And PlayerType.User Then
                            If UserList(tU).flags.Muerto = 0 Then
                                If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 4 Then
                                    Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

                                    Exit Sub

                                End If
                                 
                                '17/09/02
                                'Check the trigger
                                If MapData(UserList(tU).Pos.Map, X, Y).trigger = eTrigger.ZONASEGURA Then
                                    Call WriteConsoleMsg(UserIndex, "No puedes robar aquí.", FontTypeNames.FONTTYPE_WARNING)

                                    Exit Sub

                                End If
                                 
                                If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                    Call WriteConsoleMsg(UserIndex, "No puedes robar aquí.", FontTypeNames.FONTTYPE_WARNING)

                                    Exit Sub

                                End If
                                 
                                Call DoRobar(UserIndex, tU)

                            End If

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "¡No hay a quien robarle!", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "¡No puedes robar en zonas seguras!", FontTypeNames.FONTTYPE_INFO)

                End If

            Case eSkill.Domar
                'Modificado 25/11/02
                'Optimizado y solucionado el bug de la doma de
                'criaturas hostiles.
                
                'Target whatever is that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                tN = .flags.TargetNPC
                
                If tN > 0 Then
                    If Npclist(tN).flags.Domable > 0 Then
                        If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

                            Exit Sub

                        End If
                        
                        If LenB(Npclist(tN).flags.AttackedBy) <> 0 Then
                            Call WriteConsoleMsg(UserIndex, "No puedes domar una criatura que está luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)

                            Exit Sub

                        End If
                        
                        'mMascotas.Mascotas_AddNew UserIndex, tN
                        'Call DoDomar(UserIndex, tN)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "¡No hay ninguna criatura allí!", FontTypeNames.FONTTYPE_INFO)

                End If
           
            Case eSkill.Pesca
                
                WeaponIndex = .Invent.WeaponEqpObjIndex

                If WeaponIndex = 0 Then Exit Sub
                
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                'Basado en la idea de Barrin
                'Comentario por Barrin: jah, "basado", caradura ! ^^
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 1 Then
                    Call WriteConsoleMsg(UserIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                
                If HayAgua(.Pos.Map, .flags.TargetX, .flags.TargetY) Then
                    If Abs(.Pos.X - .flags.TargetX) + Abs(.Pos.Y - .flags.TargetY) > 6 Then
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para sacar peces.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                                 
                    Select Case WeaponIndex

                        Case CAÑA_PESCA, RED_PESCA
                            Call DoPescar(UserIndex, WeaponIndex)
 
                        Case Else

                            Exit Sub    'Invalid item!

                    End Select
                    
                    'Play sound!
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_PESCAR, .Pos.X, .Pos.Y, .Char.charindex))
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, río o mar.", FontTypeNames.FONTTYPE_INFO)

                End If

            Case eSkill.Mineria

                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                                
                WeaponIndex = .Invent.WeaponEqpObjIndex
                                
                If WeaponIndex = 0 Then Exit Sub
                
                If (WeaponIndex <> PIQUETE_MINERO) Then

                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub

                End If
                
                'Target whatever is in the tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                
                If DummyInt > 0 Then

                    'Check distance
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                    
                    '¿Hay un yacimiento donde clickeo?
                    If ObjData(DummyInt).OBJType = eOBJType.otYacimiento Then
                        Call DoMineria(UserIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)

                End If
                  
            Case TeleportInvoker 'UGLY!!! This is a constant, not a skill!!

                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                          
                'Validate other items
                If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > .CurrentInventorySlots Then

                    Exit Sub

                End If
                    
                If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <= 0 Then Exit Sub ' @@ No se si puede pasar
                    
                Call Teleports_AddNew(UserIndex, .Invent.Object(.flags.TargetObjInvSlot).ObjIndex, .Pos.Map, X, Y)

            Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
                
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                'Check there is a proper item there
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then

                        'Validate other items
                        If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > .CurrentInventorySlots Then

                            Exit Sub

                        End If
                        
                        ''chequeamos que no se zarpe duplicando oro
                        If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                            If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).Amount = 0 Then
                                Call WriteConsoleMsg(UserIndex, "No tienes más minerales.", FontTypeNames.FONTTYPE_INFO)

                                Exit Sub

                            End If
                            
                            ''FUISTE
                            Call Protocol.Kick(UserIndex)

                            Exit Sub

                        End If

                        If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales Then
                            Call FundirMineral(UserIndex)
                        ElseIf ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otWeapon Then

                            ' Call FundirArmas(UserIndex)
                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)

                End If

            Case eSkill.Talar

                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                WeaponIndex = .Invent.WeaponEqpObjIndex
                
                If WeaponIndex = 0 Then
                    
                    Call WriteConsoleMsg(UserIndex, "Deberías equiparte el hacha.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                
                If WeaponIndex <> HACHA_LEÑADOR Then

                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub

                End If
                
                DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                
                If DummyInt > 0 Then
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                    
                    'Barrin 29/9/03
                    If .Pos.X = X And .Pos.Y = Y Then
                        Call WriteConsoleMsg(UserIndex, "No puedes talar desde allí.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                    
                    '¿Hay un arbol donde clickeo?
                    If ObjData(DummyInt).OBJType = eOBJType.otArboles Then
                        If WeaponIndex = HACHA_LEÑADOR Then
                            
                            Dim Objeto As Integer

                            Objeto = ObjData(DummyInt).ArbolItem
                            
                            If Objeto = 0 Then
                                Call WriteConsoleMsg(UserIndex, "El árbol no posee leños suficientes para poder arrojar.", FontTypeNames.FONTTYPE_INFO)
                            Else
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_TALAR, .Pos.X, .Pos.Y, .Char.charindex))
                                Call DoTalar(UserIndex, Objeto)

                            End If
                            
                        Else
                            Call WriteConsoleMsg(UserIndex, "No has podido extraer leña. Comprueba los conocimientos necesarios.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "No hay ningún árbol ahí.", FontTypeNames.FONTTYPE_INFO)

                End If
            
        End Select

    End With

    '<EhFooter>
    Exit Sub

HandleWorkLeftClick_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleWorkLeftClick " & "at line " & Erl

    '</EhFooter>
End Sub

''
' Handles the "SpellInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpellInfo(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleSpellInfo_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim spellSlot As Byte

        Dim Spell     As Integer
        
        spellSlot = Reader.ReadInt()
        
        'Validate slot
        If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        'Validate spell in the slot
        Spell = .Stats.UserHechizos(spellSlot)

        If Spell > 0 And Spell < NumeroHechizos + 1 Then

            With Hechizos(Spell)
                'Send information
                Call WriteConsoleMsg(UserIndex, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & vbCrLf & "Nombre:" & .Nombre & vbCrLf & "Descripción:" & .Desc & vbCrLf & "Skill requerido: " & .MinSkill & " de magia." & vbCrLf & "Maná necesario: " & .ManaRequerido & vbCrLf & "Energía necesaria: " & .StaRequerido & vbCrLf & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%", FontTypeNames.FONTTYPE_INFO)

            End With

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleSpellInfo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleSpellInfo " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "EquipItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleEquipItem_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim itemSlot As Byte
        
        itemSlot = Reader.ReadInt()
        
        'Dead users can't equip items
        If .flags.Muerto = 1 Then Exit Sub
        
        'Validate item slot
        If itemSlot > .CurrentInventorySlots Or itemSlot < 1 Then Exit Sub
        
        If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
        
        If Not Interval_Equipped(UserIndex) Then Exit Sub
        Call EquiparInvItem(UserIndex, itemSlot)

    End With

    '<EhFooter>
    Exit Sub

HandleEquipItem_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleEquipItem " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 06/28/2008
    'Last Modified By: NicoNZ
    ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
    ' 06/28/2008: NicoNZ - Sólo se puede cambiar si está inmovilizado.
    '***************************************************

    With UserList(UserIndex)
        
        Dim Heading       As eHeading

        Dim posX          As Integer

        Dim posY          As Integer

        Dim PacketCounter As Long

        Heading = Reader.ReadInt()
        PacketCounter = Reader.ReadInt32
                        
        Dim Packet_ID As Long

        Packet_ID = PacketNames.ChangeHeading
            
        If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "ChangeHeading", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        
        ' Las clases con maná no se pueden mover.
        If .flags.Paralizado = 1 And .flags.Inmovilizado = 1 Then
            If .Stats.MaxMan <> 0 Then Exit Sub
            
        Else

            If LegalPos(.Pos.Map, .Pos.X + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then

                Exit Sub

            End If

        End If
               
        'If .flags.Paralizado = 1 And .flags.Inmovilizado = 1 Then

        'Select Case Heading

        'Case eHeading.NORTH
        '  posY = -1

        ' Case eHeading.EAST
        '   posX = 1

        ' Case eHeading.SOUTH
        '   posY = 1

        ' Case eHeading.WEST
        '    posX = -1
        ' End Select
            
        ' If LegalPos(.Pos.Map, .Pos.X + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then

        '   Exit Sub

        ' End If
        ' End If
        
        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If Heading > 0 And Heading < 5 Then
            .Char.Heading = Heading
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeHeading(.Char.charindex, .Char.Heading))

        End If

    End With

End Sub

''
' Handles the "ModifySkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleModifySkills(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleModifySkills_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 11/19/09
    '11/19/09: Pato - Adapting to new skills system.
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim i                      As Long

        Dim Count                  As Integer

        Dim Points(1 To NUMSKILLS) As Byte
        
        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        For i = 1 To NUMSKILLS
            Points(i) = Reader.ReadInt()
            
            If Points(i) < 0 Then
                Call LogHackAttemp(.Name & " IP:" & .IpAddress & " trató de hackear los skills.")
                .Stats.SkillPts = 0
                Call Protocol.Kick(UserIndex)

                Exit Sub

            End If
            
            Count = Count + Points(i)
        Next i
        
        If Count > .Stats.SkillPts Then
            Call LogHackAttemp(.Name & " IP:" & .IpAddress & " trató de hackear los skills.")
            Call Protocol.Kick(UserIndex)
            Exit Sub

        End If
        
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        .Counters.AsignedSkills = MinimoInt(10, .Counters.AsignedSkills + Count)
        
        With .Stats

            For i = 1 To NUMSKILLS

                If Points(i) > 0 Then
                    .SkillPts = .SkillPts - Points(i)
                    .UserSkills(i) = .UserSkills(i) + Points(i)
                    
                    'Client should prevent this, but just in case...
                    If .UserSkills(i) > 100 Then
                        .SkillPts = .SkillPts + .UserSkills(i) - 100
                        .UserSkills(i) = 100

                    End If
                    
                    Call CheckEluSkill(UserIndex, i, True)

                End If

            Next i

        End With

    End With

    '<EhFooter>
    Exit Sub

HandleModifySkills_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleModifySkills " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Train" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleTrain_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim SpawnedNpc As Integer

        Dim PetIndex   As Byte
        
        PetIndex = Reader.ReadInt()
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        If Npclist(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
            If PetIndex > 0 And PetIndex < Npclist(.flags.TargetNPC).NroCriaturas + 1 Then
                'Create the creature
                SpawnedNpc = SpawnNpc(Npclist(.flags.TargetNPC).Criaturas(PetIndex).NpcIndex, Npclist(.flags.TargetNPC).Pos, True, False)
                
                If SpawnedNpc > 0 Then
                    Npclist(SpawnedNpc).MaestroNpc = .flags.TargetNPC
                    Npclist(.flags.TargetNPC).Mascotas = Npclist(.flags.TargetNPC).Mascotas + 1

                End If

            End If

        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite))

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleTrain_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleTrain " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleCommerceBuy_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Slot          As Byte

        Dim Amount        As Integer
            
        Dim SelectedPrice As Byte
            
        Slot = Reader.ReadInt()
        Amount = Reader.ReadInt()
        SelectedPrice = Reader.ReadInt8()
              
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
            
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite))

            Exit Sub

        End If
        
        'Only if in commerce mode....
        If Not .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "No estás comerciando.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        'User compra el item
        Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNPC, Slot, Amount, SelectedPrice)

    End With

    '<EhFooter>
    Exit Sub

HandleCommerceBuy_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCommerceBuy " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleBankExtractItem_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim Slot     As Byte

        Dim Amount   As Integer
        
        Dim TypeBank As E_BANK
        
        Slot = Reader.ReadInt()
        Amount = Reader.ReadInt()
        TypeBank = Reader.ReadInt()
        
        If Slot <= 0 Then Exit Sub
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿Es el banquero?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then

            Exit Sub

        End If
        
        If .flags.SlotEvent > 0 Then
            If Events(.flags.SlotEvent).ChangeClass > 0 Then
                Call WriteConsoleMsg(UserIndex, "En este tipo de eventos no es posible retirar/depositar objetos.", FontTypeNames.FONTTYPE_INFORED)

                Exit Sub

            End If

        End If
        
        Select Case TypeBank

            Case E_BANK.e_User
                Call UserRetiraItem(UserIndex, Slot, Amount)

            Case E_BANK.e_Account
                Call UserRetiraItem_Account(UserIndex, Slot, Amount)

        End Select
        
    End With

    '<EhFooter>
    Exit Sub

HandleBankExtractItem_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleBankExtractItem " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "CommerceSell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleCommerceSell_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Slot          As Byte

        Dim Amount        As Integer
            
        Dim SelectedPrice As Byte
            
        Slot = Reader.ReadInt()
        Amount = Reader.ReadInt()
        SelectedPrice = Reader.ReadInt8
              
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite))

            Exit Sub

        End If
        
        'User compra el item del slot
        Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNPC, Slot, Amount, SelectedPrice)

    End With

    '<EhFooter>
    Exit Sub

HandleCommerceSell_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCommerceSell " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "BankDeposit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleBankDeposit_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Slot     As Byte

        Dim Amount   As Integer
        
        Dim TypeBank As E_BANK
        
        Slot = Reader.ReadInt()
        Amount = Reader.ReadInt()
        TypeBank = Reader.ReadInt()
        
        If Slot <= 0 Or Amount <= 0 Then Exit Sub
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then

            Exit Sub

        End If
        
        If .flags.SlotEvent > 0 Then
            If Events(.flags.SlotEvent).ChangeClass > 0 Then
                Call WriteConsoleMsg(UserIndex, "En este tipo de eventos no es posible retirar/depositar objetos.", FontTypeNames.FONTTYPE_INFORED)

                Exit Sub

            End If

        End If

        Select Case TypeBank

            Case E_BANK.e_User
                Call UserDepositaItem(UserIndex, Slot, Amount)

            Case E_BANK.e_Account
                Call UserDepositaItem_Account(UserIndex, Slot, Amount)

        End Select
        
    End With

    '<EhFooter>
    Exit Sub

HandleBankDeposit_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleBankDeposit " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "MoveSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleMoveSpell_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
        
    Dim dir     As Integer
        
    Dim SlotOld As Byte

    Dim SlotNew As Byte
    
    SlotOld = Reader.ReadInt
    SlotNew = Reader.ReadInt
        
    Call ChangeSlotSpell(UserIndex, SlotOld, SlotNew)

    '<EhFooter>
    Exit Sub

HandleMoveSpell_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleMoveSpell " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "MoveBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveBank(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleMoveBank_Err

    '</EhHeader>

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 06/14/09
    '
    '***************************************************
        
    Dim dir      As Integer

    Dim Slot     As Byte

    Dim TempItem As Obj
        
    If Reader.ReadBool() Then
        dir = 1
    Else
        dir = -1

    End If
        
    Slot = Reader.ReadInt()

    With UserList(UserIndex)
        TempItem.ObjIndex = .BancoInvent.Object(Slot).ObjIndex
        TempItem.Amount = .BancoInvent.Object(Slot).Amount
        
        If dir = 1 Then 'Mover arriba
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot - 1)
            .BancoInvent.Object(Slot - 1).ObjIndex = TempItem.ObjIndex
            .BancoInvent.Object(Slot - 1).Amount = TempItem.Amount
        Else 'mover abajo
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot + 1)
            .BancoInvent.Object(Slot + 1).ObjIndex = TempItem.ObjIndex
            .BancoInvent.Object(Slot + 1).Amount = TempItem.Amount

        End If

    End With
    
    Call UpdateBanUserInv(True, UserIndex, 0)
    Call UpdateVentanaBanco(UserIndex)

    '<EhFooter>
    Exit Sub

HandleMoveBank_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleMoveBank " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleUserCommerceOffer_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 24/11/2009
    '24/11/2009: ZaMa - Nuevo sistema de comercio
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Amount    As Long

        Dim Slot      As Byte

        Dim tUser     As Integer

        Dim OfferSlot As Byte

        Dim ObjIndex  As Integer
        
        Slot = Reader.ReadInt()
        Amount = Reader.ReadInt()
        OfferSlot = Reader.ReadInt()
        
        'Get the other player
        tUser = .ComUsu.DestUsu
        
        ' If he's already confirmed his offer, but now tries to change it, then he's cheating
        If UserList(UserIndex).ComUsu.Confirmo = True Then
            
            ' Finish the trade
            Call FinComerciarUsu(UserIndex)
        
            If tUser <= 0 Or tUser > MaxUsers Then
                Call FinComerciarUsu(tUser)
                Call Protocol.FlushBuffer(tUser)

            End If
        
            Exit Sub

        End If
        
        'If slot is invalid and it's not gold or it's not 0 (Substracting), then ignore it.
        If ((Slot < 0 Or Slot > UserList(UserIndex).CurrentInventorySlots) And Slot <> FLAGORO And Slot <> FLAGELDHIR) Then Exit Sub
        
        'If OfferSlot is invalid, then ignore it.
        If OfferSlot < 1 Or OfferSlot > MAX_OFFER_SLOTS + 2 Then Exit Sub
        
        ' Can be negative if substracted from the offer, but never 0.
        If Amount = 0 Then Exit Sub
        
        'Has he got enough??
        If Slot = FLAGORO Then

            ' Can't offer more than he has
            If Amount > .Stats.Gld - .ComUsu.GoldAmount Then
                Call WriteCommerceChat(UserIndex, "No tienes esa cantidad de oro para agregar a la oferta.", FontTypeNames.FONTTYPE_TALK)

                Exit Sub

            End If
            
            If Amount < 0 Then
                If Abs(Amount) > .ComUsu.GoldAmount Then
                    Amount = .ComUsu.GoldAmount * (-1)

                End If

            End If

        ElseIf Slot = FLAGELDHIR Then

            ' Can't offer more than he has
            If Amount > .Stats.Eldhir - .ComUsu.EldhirAmount Then
                Call WriteCommerceChat(UserIndex, "No tienes esa cantidad de Eldhir para agregar a la oferta.", FontTypeNames.FONTTYPE_TALK)

                Exit Sub

            End If
            
            If Amount < 0 Then
                If Abs(Amount) > .ComUsu.EldhirAmount Then
                    Amount = .ComUsu.EldhirAmount * (-1)

                End If

            End If

        Else

            'If modifing a filled offerSlot, we already got the objIndex, then we don't need to know it
            If Slot <> 0 Then ObjIndex = .Invent.Object(Slot).ObjIndex

            ' Can't offer more than he has
            If Not HasEnoughItems(UserIndex, ObjIndex, TotalOfferItems(ObjIndex, UserIndex) + Amount) Then
                
                Call WriteCommerceChat(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)

                Exit Sub

            End If
            
            If Amount < 0 Then
                If Abs(Amount) > .ComUsu.cant(OfferSlot) Then
                    Amount = .ComUsu.cant(OfferSlot) * (-1)

                End If

            End If
        
            If ItemNewbie(ObjIndex) Then
                Call WriteCancelOfferItem(UserIndex, OfferSlot)

                Exit Sub

            End If
            
            'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
            If .flags.Navegando = 1 Then
                If .Invent.BarcoSlot = Slot Then
                    Call WriteCommerceChat(UserIndex, "No puedes vender tu barco mientras lo estés usando.", FontTypeNames.FONTTYPE_TALK)

                    Exit Sub

                End If

            End If
            
            If .Invent.MochilaEqpSlot > 0 Then
                If .Invent.MochilaEqpSlot = Slot Then
                    Call WriteCommerceChat(UserIndex, "No puedes vender tu mochila mientras la estés usando.", FontTypeNames.FONTTYPE_TALK)

                    Exit Sub

                End If

            End If
            
            If ObjData(ObjIndex).OBJType = otGemaTelep Then
                Call WriteCommerceChat(UserIndex, "No puedes vender los scrolls de viajes.", FontTypeNames.FONTTYPE_TALK)

                Exit Sub

            End If
            
            If Not EsGmPriv(UserIndex) Then
                If ObjData(ObjIndex).NoNada = 1 Then
                    Call WriteCommerceChat(UserIndex, "No puedes realizar ninguna acción con este objeto. ¡Podría ser de uso personal!", FontTypeNames.FONTTYPE_TALK)
    
                    Exit Sub
    
                End If

            End If

        End If
        
        Call AgregarOferta(UserIndex, OfferSlot, ObjIndex, Amount, Slot = FLAGORO, Slot = FLAGELDHIR)
        Call EnviarOferta(tUser, OfferSlot)

    End With

    '<EhFooter>
    Exit Sub

HandleUserCommerceOffer_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleUserCommerceOffer " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Online" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnline(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleOnline_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i     As Long

    Dim Count As Long
    
    With UserList(UserIndex)
        
        Dim ArmadasON   As Long

        Dim CaosON      As Long

        Dim lstName     As String
            
        Dim lstCaos     As String

        Dim lstArmada   As String
            
        Dim ViewFaction As Boolean
            
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoyalCouncil Or PlayerType.ChaosCouncil)) <> 0 Then
            ViewFaction = True

        End If
            
        For i = 1 To LastUser

            If Len(UserList(i).Account.Email) > 0 Then

                'If UserList(i).flags.Privilegios And (PlayerType.User ) Then
                If UserList(i).Faction.Status = r_Caos Then
                    CaosON = CaosON + 1

                    If ViewFaction Then
                        lstCaos = lstCaos & UserList(i).Name & ", "

                    End If

                ElseIf UserList(i).Faction.Status = r_Armada Then
                    ArmadasON = ArmadasON + 1

                    If ViewFaction Then
                        lstArmada = lstArmada & UserList(i).Name & ", "

                    End If

                End If
                
                lstName = lstName & UserList(i).Name & ", "
                Count = Count + 1
                    
                'End If
            End If

        Next i
        
        If Count > 0 Then
            lstName = Left$(lstName, Len(lstName) - 2)

        End If
        
        If ViewFaction Then
            If ArmadasON > 0 Then
                lstArmada = Left$(lstArmada, Len(lstArmada) - 2)

            End If
    
            If CaosON > 0 Then
                lstCaos = Left$(lstCaos, Len(lstCaos) - 2)

            End If

        End If
            
        Count = Count + UsersBot
        Call WriteConsoleMsg(UserIndex, "Número de usuarios online: " & CStr(Count) & ". El Record de usuarios conectados simultaneamente fue de " & RECORDusuarios, FontTypeNames.FONTTYPE_INFO)

        If EsGmPriv(UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "Nombres de los usuarios: " & lstName, FontTypeNames.FONTTYPE_INFO)

        End If

        Call WriteConsoleMsg(UserIndex, "Usuarios de la facción <Legión Oscura>: " & CStr(CaosON) & ". " & lstCaos, FontTypeNames.FONTTYPE_INFORED)
        Call WriteConsoleMsg(UserIndex, "Usuarios de la facción <Armada Real>: " & CStr(ArmadasON) & ". " & lstArmada, FontTypeNames.FONTTYPE_INFOGREEN)

    End With

    '<EhFooter>
    Exit Sub

HandleOnline_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleOnline " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Quit_AddNew(ByVal UserIndex As Integer, ByVal IsAccount As Boolean)

    '<EhHeader>
    On Error GoTo Quit_AddNew_Err

    '</EhHeader>

    Dim tUser As Integer
        
    With UserList(UserIndex)

        'exit secure commerce
        If .ComUsu.DestUsu > 0 Then
            tUser = .ComUsu.DestUsu
            
            If UserList(tUser).flags.UserLogged Then
                If UserList(tUser).ComUsu.DestUsu = UserIndex Then
                    Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(tUser)

                End If

            End If
            
            Call WriteConsoleMsg(UserIndex, "Comercio cancelado.", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(UserIndex)

        End If

        .flags.DeslogeandoCuenta = IsAccount
        Call Cerrar_Usuario(UserIndex)
    
    End With

    '<EhFooter>
    Exit Sub

Quit_AddNew_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.Quit_AddNew " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Quit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/2008 (NicoNZ)
    'If user is invisible, it automatically becomes
    'visible before doing the countdown to exit
    '04/15/2008 - No se reseteaban lso contadores de invi ni de ocultar. (NicoNZ)
    '***************************************************
    Dim tUser     As Integer

    Dim IsAccount As Boolean
        
    IsAccount = Reader.ReadBool
        
    Dim isNotVisible As Boolean
    
    With UserList(UserIndex)
        
        If .flags.Paralizado = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)

            Exit Sub

        End If
        
        Quit_AddNew UserIndex, IsAccount
    
    End With

End Sub

''
' Handles the "Meditate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMeditate(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleMeditate_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/08 (NicoNZ)
    'Arreglé un bug que mandaba un index de la meditacion diferente
    'al que decia el server.
    '***************************************************
    With UserList(UserIndex)
        
        'Si ya tiene el mana completo, no lo dejamos meditar.
        If .Stats.MinMan = .Stats.MaxMan Then Exit Sub

        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes meditar cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        'Can he meditate?
        If .Stats.MaxMan = 0 Then
            Call WriteConsoleMsg(UserIndex, "Sólo las clases mágicas conocen el arte de la meditación.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
            
        If .flags.TeleportInvoker > 0 Then
            Exit Sub

        End If
            
        .flags.Meditando = Not .flags.Meditando

        If .flags.Meditando Then
            .Char.loops = INFINITE_LOOPS
            .Counters.TimerMeditar = 0
            .Counters.TiempoInicioMeditar = 0

            If .MeditationSelected = 0 Then
                .Char.FX = UserFxMeditation(UserIndex)
            Else
                .Char.FX = Meditation(.MeditationSelected)

            End If
            
        Else
            
            .Char.FX = 0

        End If

        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, .Char.FX, .Pos.X, .Pos.Y))

    End With

    '<EhFooter>
    Exit Sub

HandleMeditate_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleMeditate " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function UserFxMeditation(ByVal UserIndex As Integer) As Integer

    '<EhHeader>
    On Error GoTo UserFxMeditation_Err

    '</EhHeader>
    
    With UserList(UserIndex)

        If .Stats.Elv < 15 Then
            UserFxMeditation = FXIDs.FXMEDITARCHICO

        ElseIf .Stats.Elv < 30 Then
            UserFxMeditation = FXIDs.FXMEDITARMEDIANO

        ElseIf .Stats.Elv < 45 Then
            UserFxMeditation = FXIDs.FXMEDITARGRANDE ' Celeste Mediana
                
        ElseIf .Stats.Elv < STAT_MAXELV Then
            UserFxMeditation = FXIDs.FXMEDITARXGRANDE

        Else
            UserFxMeditation = FXIDs.FXMEDITARXXXGRANDE

        End If

    End With
    
    '<EhFooter>
    Exit Function

UserFxMeditation_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.UserFxMeditation " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Handles the "Resucitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleResucitate_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        'Validate NPC and make sure player is dead
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor And (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) Or .flags.Muerto = 0 Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede resucitarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        Call RevivirUsuario(UserIndex)
        .Stats.MinHp = .Stats.MaxHp
        Call WriteUpdateHP(UserIndex)
        Call WriteConsoleMsg(UserIndex, "¡¡Has sido resucitado!!", FontTypeNames.FONTTYPE_INFO)

    End With

    '<EhFooter>
    Exit Sub

HandleResucitate_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleResucitate " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Consultation" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleConsultation(ByVal UserIndex As String)

    '<EhHeader>
    On Error GoTo HandleConsultation_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 01/05/2010
    'Habilita/Deshabilita el modo consulta.
    '01/05/2010: ZaMa - Agrego validaciones.
    '16/09/2010: ZaMa - No se hace visible en los clientes si estaba navegando (porque ya lo estaba).
    '***************************************************
    
    Dim UserConsulta As Integer
    
    With UserList(UserIndex)
        
        ' Comando exclusivo para gms
        If Not EsGm(UserIndex) Then Exit Sub
        
        UserConsulta = .flags.TargetUser
        
        'Se asegura que el target es un usuario
        If UserConsulta = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un usuario, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        ' No podes ponerte a vos mismo en modo consulta.
        If UserConsulta = UserIndex Then Exit Sub
        
        ' No podes estra en consulta con otro gm
        If EsGm(UserConsulta) Then
            Call WriteConsoleMsg(UserIndex, "No puedes iniciar el modo consulta con otro administrador.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        Dim UserName As String

        UserName = UserList(UserConsulta).Name
        
        ' Si ya estaba en consulta, termina la consulta
        If UserList(UserConsulta).flags.EnConsulta Then
            Call WriteConsoleMsg(UserIndex, "Has terminado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has terminado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Termino consulta con " & UserName)
            UserList(UserConsulta).flags.EnConsulta = False
        
            ' Sino la inicia
        Else
            Call WriteConsoleMsg(UserIndex, "Has iniciado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has iniciado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Inicio consulta con " & UserName)
            
            With UserList(UserConsulta)
                .flags.EnConsulta = True
                
                ' Pierde invi u ocu
                If .flags.Invisible = 1 Or .flags.Oculto = 1 Then
                    .flags.Oculto = 0
                    .flags.Invisible = 0
                    .Counters.TiempoOculto = 0
                    .Counters.Invisibilidad = 0
                    
                    If UserList(UserConsulta).flags.Navegando = 0 Then
                        Call UsUaRiOs.SetInvisible(UserConsulta, UserList(UserConsulta).Char.charindex, False)

                    End If

                End If

            End With

        End If
        
        Call UsUaRiOs.SetConsulatMode(UserConsulta)

    End With

    '<EhFooter>
    Exit Sub

HandleConsultation_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleConsultation " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Heal" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHeal(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleHeal_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor And Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie) Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        .Stats.MinHp = .Stats.MaxHp
        
        Call WriteUpdateHP(UserIndex)
        
        Call WriteConsoleMsg(UserIndex, "¡¡Has sido curado!!", FontTypeNames.FONTTYPE_INFO)

    End With

    '<EhFooter>
    Exit Sub

HandleHeal_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleHeal " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "RequestStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestStats(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRequestStats_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call SendUserStatsTxt(UserIndex, UserIndex)
    '<EhFooter>
    Exit Sub

HandleRequestStats_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRequestStats " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Help" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHelp(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleHelp_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call SendHelp(UserIndex)
    '<EhFooter>
    Exit Sub

HandleHelp_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleHelp " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "CommerceStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleCommerceStart_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i As Integer

    With UserList(UserIndex)
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        'Is it already in commerce mode??
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya estás comerciando.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If .flags.SlotEvent > 0 Then Exit Sub
        If .flags.SlotFast > 0 Then Exit Sub
        If .flags.SlotReto > 0 Then Exit Sub
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then

            'Does the NPC want to trade??
            If Npclist(.flags.TargetNPC).Comercia = 0 Then
                If LenB(Npclist(.flags.TargetNPC).Desc) <> 0 Then
                    Call WriteChatOverHead(UserIndex, "No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)

                End If
                
                Exit Sub

            End If
            
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 5 Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            'Start commerce....
            Call IniciarComercioNPC(UserIndex)
            '[Alejo]
        ElseIf .flags.TargetUser > 0 Then
        
            'User commerce...
            'Can he commerce??
            If .flags.Privilegios And PlayerType.SemiDios Then
                Call WriteConsoleMsg(UserIndex, "No puedes vender ítems.", FontTypeNames.FONTTYPE_WARNING)

                Exit Sub

            End If
            
            'Is the other one dead??
            If UserList(.flags.TargetUser).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡¡No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            'Is it me??
            If .flags.TargetUser = UserIndex Then
                Call WriteConsoleMsg(UserIndex, "¡¡No puedes comerciar con vos mismo!!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            'Check distance
            If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 5 Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            'Is he already trading?? is it with me or someone else??
            If UserList(.flags.TargetUser).flags.Comerciando = True And UserList(.flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            If .Stats.Elv < 4 Then
                Call WriteConsoleMsg(UserIndex, "¡Entrena hasta Nivel 4 para usar este comando!", FontTypeNames.FONTTYPE_INFORED)
    
                Exit Sub
    
            End If
            
            ' 133
            If MapInfo(.Pos.Map).Pk Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar en ZONA INSEGURA.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            If Not Interval_Commerce(UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "¡¡Debes esperar algunos segundos para enviar solicitud!!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            'Initialize some variables...
            .ComUsu.DestUsu = .flags.TargetUser
            .ComUsu.DestNick = UserList(.flags.TargetUser).Name

            For i = 1 To MAX_OFFER_SLOTS
                .ComUsu.cant(i) = 0
                .ComUsu.Objeto(i) = 0
            Next i

            .ComUsu.GoldAmount = 0
            .ComUsu.EldhirAmount = 0
            .ComUsu.Acepto = False
            .ComUsu.Confirmo = False
            
            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(UserIndex, .flags.TargetUser)
        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleCommerceStart_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCommerceStart " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "BankStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankStart(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleBankStart_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim TypeBank As E_BANK
    
    TypeBank = Reader.ReadInt
    
    With UserList(UserIndex)
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya estás comerciando.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 5 Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            'If it's the banker....
            If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Banquero Then

                Select Case TypeBank

                    Case E_BANK.e_User, E_BANK.e_Account
                        Call IniciarDeposito(UserIndex, TypeBank)
                        
                    Case Else

                End Select
                
            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleBankStart_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleBankStart " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ShareNpc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShareNpc(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleShareNpc_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/04/2010
    'Shares owned npcs with other user
    '***************************************************
    
    Dim targetUserIndex  As Integer

    Dim SharingUserIndex As Integer
    
    With UserList(UserIndex)
        
        ' Didn't target any user
        targetUserIndex = .flags.TargetUser

        If targetUserIndex = 0 Then Exit Sub
        
        ' Can't share with admins
        If EsGm(targetUserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No puedes compartir npcs con administradores!!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        ' Pk or Caos?
        If Escriminal(UserIndex) Then

            ' Caos can only share with other caos
            If esCaos(UserIndex) Then
                If Not esCaos(targetUserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Solo puedes compartir npcs con miembros de tu misma facción!!", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                
                ' Pks don't need to share with anyone
            Else

                Exit Sub

            End If
        
            ' Ciuda or Army?
        Else

            ' Can't share
            If Escriminal(targetUserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No puedes compartir npcs con criminales!!", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If

        End If
        
        ' Already sharing with target
        SharingUserIndex = .flags.ShareNpcWith

        If SharingUserIndex = targetUserIndex Then Exit Sub
        
        ' Aviso al usuario anterior que dejo de compartir
        If SharingUserIndex <> 0 Then
            Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)

        End If
        
        .flags.ShareNpcWith = targetUserIndex
        
        Call WriteConsoleMsg(targetUserIndex, .Name & " ahora comparte sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Ahora compartes tus npcs con " & UserList(targetUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleShareNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleShareNpc " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "StopSharingNpc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleStopSharingNpc(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleStopSharingNpc_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/04/2010
    'Stop Sharing owned npcs with other user
    '***************************************************
    
    Dim SharingUserIndex As Integer
    
    With UserList(UserIndex)
        
        SharingUserIndex = .flags.ShareNpcWith
        
        If SharingUserIndex <> 0 Then
            
            ' Aviso al que compartia y al que le compartia.
            Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SharingUserIndex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
            
            .flags.ShareNpcWith = 0

        End If
        
    End With

    '<EhFooter>
    Exit Sub

HandleStopSharingNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleStopSharingNpc " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "PartyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyMessage(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandlePartyMessage_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim chat As String
        
        chat = Reader.ReadString8()
        
        If .GroupIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, "No conformas ninguna party", FontTypeNames.FONTTYPE_INFO)
            
        Else

            If Interval_Message(UserIndex) Then
                If LenB(chat) <> 0 Then
                    SendMessageGroup .GroupIndex, .Name, chat

                End If

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandlePartyMessage_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandlePartyMessage " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "CouncilMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilMessage(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleCouncilMessage_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim chat As String
        
        chat = Reader.ReadString8()
        
        Dim ValidChat As Boolean

        ValidChat = True
        
        If .flags.SlotEvent > 0 Then
            If Events(.flags.SlotEvent).Modality = eModalityEvent.DeathMatch Then ValidChat = False

        End If
        
        If LenB(chat) <> 0 And ValidChat Then
            
            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                Call SendData(SendTarget.ToConsejoYCaos, UserIndex, PrepareMessageConsoleMsg("(Privado Consejo) " & .Name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJO))
            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                Call SendData(SendTarget.ToConsejoYCaos, UserIndex, PrepareMessageConsoleMsg("(Privado Concilio) " & .Name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleCouncilMessage_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCouncilMessage " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeDescription_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '********

    With UserList(UserIndex)
        
        Dim description As String
        
        description = Reader.ReadString8()
        
        If .Account.Premium > 1 Then
            If .flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes cambiar la descripción estando muerto.", FontTypeNames.FONTTYPE_INFO)
            Else

                If Not AsciiValidos(description) Then
                    Call WriteConsoleMsg(UserIndex, "La descripción tiene caracteres inválidos.", FontTypeNames.FONTTYPE_INFO)
                Else
                    .Desc = Trim$(description)
                    Call WriteConsoleMsg(UserIndex, "La descripción ha cambiado.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Solo las cuentas TIER 2 o superior pueden cambiar la descripción de sus personajes. Consulta las promociones en /SHOP", FontTypeNames.FONTTYPE_INFO)
        
        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleChangeDescription_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeDescription " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Punishments" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePunishments(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandlePunishments_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 25/08/2009
    '25/08/2009: ZaMa - Now only admins can see other admins' punishment list
    '***************************************************

    With UserList(UserIndex)
        
        Dim Name  As String

        Dim Count As Integer
        
        Name = Reader.ReadString8()
        
        If LenB(Name) <> 0 Then
            If (InStrB(Name, "\") <> 0) Then
                Name = Replace(Name, "\", "")

            End If

            If (InStrB(Name, "/") <> 0) Then
                Name = Replace(Name, "/", "")

            End If

            If (InStrB(Name, ":") <> 0) Then
                Name = Replace(Name, ":", "")

            End If

            If (InStrB(Name, "|") <> 0) Then
                Name = Replace(Name, "|", "")

            End If
            
            If UCase$(Name) = UCase$(.Name) Then
                If FileExist(CharPath & Name & ".chr", vbNormal) Then
                    Count = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))

                    If Count = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
                    Else

                        While Count > 0

                            Call WriteConsoleMsg(UserIndex, Count & " - " & GetVar(CharPath & Name & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
                            Count = Count - 1

                        Wend

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "Personaje """ & Name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
            
                If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
                    If (EsAdmin(Name) Or EsDios(Name) Or EsSemiDios(Name)) And (UserList(UserIndex).flags.Privilegios And PlayerType.User) Then
                        Call WriteConsoleMsg(UserIndex, "No puedes ver las penas de los administradores.", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If FileExist(CharPath & Name & ".chr", vbNormal) Then
                            Count = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))

                            If Count = 0 Then
                                Call WriteConsoleMsg(UserIndex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
                            Else

                                While Count > 0

                                    Call WriteConsoleMsg(UserIndex, Count & " - " & GetVar(CharPath & Name & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
                                    Count = Count - 1

                                Wend

                            End If

                        Else
                            Call WriteConsoleMsg(UserIndex, "Personaje """ & Name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If

                End If

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandlePunishments_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandlePunishments " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Gamble" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGamble(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleGamble_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '10/07/2010: ZaMa - Now normal npcs don't answer if asked to gamble.
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Amount  As Integer

        Dim TypeNpc As eNPCType
        
        Amount = Reader.ReadInt()
        
        ' Dead?
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
        
            'Validate target NPC
        ElseIf .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
        
            ' Validate Distance
        ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        
            ' Validate NpcType
        ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
            
            Dim TargetNpcType As eNPCType

            TargetNpcType = Npclist(.flags.TargetNPC).NPCtype
            
            ' Normal npcs don't speak
            If TargetNpcType <> eNPCType.Comun And TargetNpcType <> eNPCType.DRAGON And TargetNpcType <> eNPCType.Pretoriano Then
                Call WriteChatOverHead(UserIndex, "No tengo ningún interés en apostar.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)

            End If
            
            ' Validate amount
        ElseIf Amount < 1 Then
            Call WriteChatOverHead(UserIndex, "El mínimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
        
            ' Validate amount
        ElseIf Amount > 50000 Then
            Call WriteChatOverHead(UserIndex, "El máximo de apuesta es 50000 Monedas de Oro.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
        
            ' Validate user gold
        ElseIf .Stats.Gld < Amount Then
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
        
        Else

            If RandomNumber(1, 100) <= 47 Then
                .Stats.Gld = .Stats.Gld + Amount
                Call WriteChatOverHead(UserIndex, "¡Felicidades! Has ganado " & CStr(Amount) & " Monedas de Oro.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
                
                Apuestas.Perdidas = Apuestas.Perdidas + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                .Stats.Gld = .Stats.Gld - Amount
                Call WriteChatOverHead(UserIndex, "Lo siento, has perdido " & CStr(Amount) & " Monedas de Oro.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
                
                Apuestas.Ganancias = Apuestas.Ganancias + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))

            End If
            
            Apuestas.Jugadas = Apuestas.Jugadas + 1
            
            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
            Call WriteUpdateGold(UserIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleGamble_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGamble " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "BankGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankGold(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleBankGold_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Amount  As Long

        Dim TypeGLD As Byte

        Dim Extract As Boolean
        
        Amount = Reader.ReadInt()
        TypeGLD = Reader.ReadInt()
        Extract = Reader.ReadBool()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If .flags.SlotEvent > 0 Then
            If Events(.flags.SlotEvent).ChangeClass > 0 Then
                Call WriteConsoleMsg(UserIndex, "En este tipo de eventos no es posible retirar/depositar objetos.", FontTypeNames.FONTTYPE_INFORED)

                Exit Sub

            End If

        End If
        
        Select Case TypeGLD

            Case 0 ' Monedas de Oro
                    
                If Extract Then
                          
                    If (Amount > 0 And Amount <= .Account.Gld) Then

                        If Amount + .Stats.Gld < MAXORO Then
                            .Account.Gld = .Account.Gld - Amount
                            .Stats.Gld = .Stats.Gld + Amount
                                
                            Call WriteChatOverHead(UserIndex, "Tenés " & .Account.Gld & " Monedas de Oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)

                        End If

                    Else
                        .Stats.Gld = .Stats.Gld + .Account.Gld
                        .Account.Gld = 0

                    End If

                    If .Stats.Gld > MAXORO Then
                        .Stats.Gld = MAXORO

                    End If

                Else

                    If Amount > 0 And Amount <= .Stats.Gld Then
                        .Account.Gld = .Account.Gld + Amount
                        .Stats.Gld = .Stats.Gld - Amount
                        Call WriteChatOverHead(UserIndex, "Tenés " & .Account.Gld & " Monedas de Oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
                                
                    Else
                        .Account.Gld = .Account.Gld + .Stats.Gld
                        .Stats.Gld = 0
                        Call WriteChatOverHead(UserIndex, "Tenés " & .Account.Gld & " Monedas de Oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
            
                    End If
                        
                    If .Account.Gld > MAXORO Then
                        .Account.Gld = MAXORO

                    End If

                End If
                
                Call WriteUpdateGold(UserIndex)
            
            Case 1 ' Monedas de Eldhir

                If Extract Then
                    If Amount > 0 And Amount <= .Account.Eldhir Then
                        .Account.Eldhir = .Account.Eldhir - Amount
                        .Stats.Eldhir = .Stats.Eldhir + Amount
                        Call WriteChatOverHead(UserIndex, "Tenés " & .Account.Eldhir & " Monedas de Eldhir en tu cuenta.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
                    Else
                        .Stats.Eldhir = .Stats.Eldhir + .Account.Eldhir
                        .Account.Eldhir = 0

                    End If

                Else

                    If Amount > 0 And Amount <= .Stats.Eldhir Then
                        .Account.Eldhir = .Account.Eldhir + Amount
                        .Stats.Eldhir = .Stats.Eldhir - Amount
                        Call WriteChatOverHead(UserIndex, "Tenés " & .Account.Eldhir & " Monedas de Eldhir en tu cuenta.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)

                    Else
                        .Account.Eldhir = .Account.Eldhir + .Stats.Eldhir
                        .Stats.Eldhir = 0
                        Call WriteChatOverHead(UserIndex, "Tenés " & .Account.Eldhir & " Monedas de Eldhir en tu cuenta.", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
            
                    End If
                
                End If
                
                Call WriteUpdateDsp(UserIndex)
            
        End Select
        
        Call WriteUpdateBankGold(UserIndex)
        
    End With

    '<EhFooter>
    Exit Sub

HandleBankGold_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleBankGold " & "at line " & Erl

    '</EhFooter>
End Sub

''
' Handles the "Denounce" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleDenounce_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 14/11/2010
    '14/11/2010: ZaMa - Now denounces can be desactivated.
    '***************************************************

    With UserList(UserIndex)
        
        Dim Text As String

        Dim msg  As String
        
        Text = Reader.ReadString8()
        
        Dim ValidChat As Boolean

        ValidChat = True
        
        If UCase$(Left$(Text, 11)) <> "[SEGURIDAD]" Then
            If .flags.SlotEvent > 0 Then
                If Events(.flags.SlotEvent).Modality = eModalityEvent.DeathMatch Then ValidChat = False

            End If

        End If
        
        If Len(Text) < 10 Then
            Call WriteConsoleMsg(UserIndex, "Por favor, utiliza este comando para describir tu error de forma concreta. No solicites GMS, ni pongas cosas sin explicarlas de forma prolija. Queremos ayudarte rápido, ayudanos vos a nosotros", FontTypeNames.FONTTYPE_INFO)
            ValidChat = False

        End If
        
        If .flags.Silenciado = 0 And ValidChat And (.Counters.TimeDenounce = 0) Then
            
            If UCase$(Left$(Text, 11)) = "[SEGURIDAD]" Then
                '   .flags.ToleranceCheat = .flags.ToleranceCheat + 1

                ' If .flags.ToleranceCheat >= 5 Then
                ' .flags.ToleranceCheat = 0
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[ANTI-CHEAT] " & .Name & ": " & Text, FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
                Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, .Name & " IP: " & .Account.Sec.IP_Address & " Email: " & .Account.Email & " : " & Text)
                'End If
                
            ElseIf UCase$(Left$(Text, 15)) = "[FOTODENUNCIAS]" Then
                SendData SendTarget.ToGM, 0, PrepareMessageConsoleMsg(Text & ". Hecha por: " & .Name, FontTypeNames.FONTTYPE_INFO)
                .Counters.TimeDenounce = 20
            Else
                msg = LCase$(.Name) & " DENUNCIA: " & Text

                Call Denuncias.Push(msg, False)
        
                Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(msg, FontTypeNames.FONTTYPE_GUILDMSG), True)
                Call WriteConsoleMsg(UserIndex, "Denuncia enviada. Si quieres comunicarte mediante whatsapp y recibir una respuesta rápida ingresa a WWW.ARGENTUMGAME.COM", FontTypeNames.FONTTYPE_INFO)
                .Counters.TimeDenounce = 5

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleDenounce_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleDenounce " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "GMMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMMessage(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleGMMessage_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '***************************************************

    With UserList(UserIndex)
        
        Dim Message As String

        Dim Priv    As Boolean
        
        Message = Reader.ReadString8()
        Priv = Reader.ReadBool()
        
        If Not EsGm(UserIndex) Then Exit Sub
        
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Mensaje a Gms:" & Message)
        
        If LenB(Message) <> 0 Then
            Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(.Name & "> " & Message, FontTypeNames.FONTTYPE_ADMIN))

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleGMMessage_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGMMessage " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ShowName" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleShowName_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.Admin) Then
            .ShowName = Not .ShowName 'Show / Hide the name
            
            Call RefreshCharStatus(UserIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleShowName_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleShowName " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineChaosLegion(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleOnlineChaosLegion_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 28/05/2010
    '28/05/2010: ZaMa - Ahora solo dioses pueden ver otros dioses online.
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim i    As Long

        Dim List As String

        Dim Priv As PlayerType

        Priv = PlayerType.User Or PlayerType.SemiDios
        
        ' Solo dioses pueden ver otros dioses online
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            Priv = Priv Or PlayerType.Dios Or PlayerType.Admin

        End If
     
        For i = 1 To LastUser

            If UserList(i).ConnIDValida Then
                If UserList(i).Faction.Status = r_Caos Then
                    If UserList(i).flags.Privilegios And Priv Then
                        List = List & UserList(i).Name & ", "

                    End If

                End If

            End If

        Next i

    End With

    If Len(List) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Caos conectados: " & Left$(List, Len(List) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay Caos conectados.", FontTypeNames.FONTTYPE_INFO)

    End If

    '<EhFooter>
    Exit Sub

HandleOnlineChaosLegion_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleOnlineChaosLegion " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ServerTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleServerTime_Err

    '</EhHeader>
    
    With UserList(UserIndex)

        If Not EsGmPriv(UserIndex) Then Exit Sub
    
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Hora.")

    End With
    
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & Time & " " & Date, FontTypeNames.FONTTYPE_INFO))
    '<EhFooter>
    Exit Sub

HandleServerTime_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleServerTime " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Where" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleWhere_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 18/11/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '18/11/2010: ZaMa - Obtengo los privs del charfile antes de mostrar la posicion de un usuario offline.
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName   As String

        Dim tUser      As Integer

        Dim miPos      As String
        
        Dim Guild      As Boolean
        
        Dim GuildIndex As Integer
        
        UserName = Reader.ReadString8()
        Guild = Reader.ReadBool()
        
        If Not EsGmPriv(UserIndex) Then Exit Sub

        If Guild Then
            GuildIndex = Guilds_SearchIndex(UCase$(UserName))
            
            If GuildIndex = 0 Then
                Call WriteConsoleMsg(UserIndex, "¡Clan inexistente!", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub

            End If
            
            Call Guilds_PrepareOnline(UserIndex, GuildIndex)
            
        Else
            tUser = NameIndex(UserName)
    
            If tUser <= 0 Then
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    
                    miPos = GetVar(CharPath & UserName & ".chr", "INIT", "POSITION")
                    Call WriteConsoleMsg(UserIndex, "Ubicación  " & UserName & " (Offline): " & ReadField(1, miPos, 45) & ", " & ReadField(2, miPos, 45) & ", " & ReadField(3, miPos, 45) & ".", FontTypeNames.FONTTYPE_INFO)

                End If
    
            Else
                Call WriteConsoleMsg(UserIndex, "Ubicación  " & UserName & ": " & UserList(tUser).Pos.Map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/Donde " & UserName)
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleWhere_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleWhere " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "CreaturesInMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleCreaturesInMap_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 30/07/06
    'Pablo (ToxicWaste): modificaciones generales para simplificar la visualización.
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Map As Integer

        Dim i, j As Long

        Dim NPCcount1, NPCcount2 As Integer

        Dim NPCcant1() As Integer

        Dim NPCcant2() As Integer

        Dim List1()    As String

        Dim List2()    As String
        
        Map = Reader.ReadInt()
        
        If Not EsGmPriv(UserIndex) Then Exit Sub
        
        If MapaValido(Map) Then

            For i = 1 To LastNPC

                'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                If Npclist(i).Pos.Map = Map Then

                    '¿esta vivo?
                    If Npclist(i).flags.NPCActive And Npclist(i).Hostile = 1 And Npclist(i).flags.AIAlineacion = 2 Then
                        If NPCcount1 = 0 Then
                            ReDim List1(0) As String
                            ReDim NPCcant1(0) As Integer
                            NPCcount1 = 1
                            List1(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant1(0) = 1
                        Else

                            For j = 0 To NPCcount1 - 1

                                If Left$(List1(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
                                    List1(j) = List1(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant1(j) = NPCcant1(j) + 1

                                    Exit For

                                End If

                            Next j

                            If j = NPCcount1 Then
                                ReDim Preserve List1(0 To NPCcount1) As String
                                ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
                                NPCcount1 = NPCcount1 + 1
                                List1(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant1(j) = 1

                            End If

                        End If

                    Else

                        If NPCcount2 = 0 Then
                            ReDim List2(0) As String
                            ReDim NPCcant2(0) As Integer
                            NPCcount2 = 1
                            List2(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant2(0) = 1
                        Else

                            For j = 0 To NPCcount2 - 1

                                If Left$(List2(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
                                    List2(j) = List2(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant2(j) = NPCcant2(j) + 1

                                    Exit For

                                End If

                            Next j

                            If j = NPCcount2 Then
                                ReDim Preserve List2(0 To NPCcount2) As String
                                ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
                                NPCcount2 = NPCcount2 + 1
                                List2(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant2(j) = 1

                            End If

                        End If

                    End If

                End If

            Next i
            
            Call WriteConsoleMsg(UserIndex, "Npcs Hostiles en mapa: ", FontTypeNames.FONTTYPE_WARNING)

            If NPCcount1 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay NPCS Hostiles.", FontTypeNames.FONTTYPE_INFO)
            Else

                For j = 0 To NPCcount1 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
                Next j

            End If

            Call WriteConsoleMsg(UserIndex, "Otros Npcs en mapa: ", FontTypeNames.FONTTYPE_WARNING)

            If NPCcount2 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay más NPCS.", FontTypeNames.FONTTYPE_INFO)
            Else

                For j = 0 To NPCcount2 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
                Next j

            End If

            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Numero enemigos en mapa " & Map)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleCreaturesInMap_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCreaturesInMap " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "WarpChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleWarpChar_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/03/2009
    '26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim Map      As Integer

        Dim X        As Integer

        Dim Y        As Integer

        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        Map = Reader.ReadInt()
        X = Reader.ReadInt()
        Y = Reader.ReadInt()
        
        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")

        End If
        
        If Not EsGm(UserIndex) Then Exit Sub
        If Not MapaValido(Map) Then Exit Sub
              
        If UCase$(UserName) = "YO" Then
            tUser = UserIndex
                  
            ' @ Si no son DIOS, no pueden ir a ZONA INSEGURA.
            If Not EsGmDios(UserIndex) And MapInfo(Map).Pk = True Then Exit Sub
        Else

            ' @ Si no son DIOS NO PUEDEN TEPEAR USUARIOS
            If Not EsGmDios(UserIndex) Then Exit Sub
            tUser = NameIndex(UserName)

        End If
        
        If tUser <= 0 Then
            If (EsDios(UserName) Or EsAdmin(UserName) And Not EsAdmin(.Name)) Then
                Call WriteConsoleMsg(UserIndex, "No puedes transportar dioses o admins.", FontTypeNames.FONTTYPE_INFO)
            Else

                If InMapBounds(Map, X, Y) Then
                    If PersonajeExiste(UserName) Then
                        Call WriteVar(CharPath & UCase$(UserName) & ".chr", "INIT", "Position", Map & "-" & X & "-" & Y)
                        Call WriteConsoleMsg(UserIndex, "Usuario offline. Se ha modificado su posición.", FontTypeNames.FONTTYPE_INFO)
        
                    Else
                        Call WriteConsoleMsg(UserIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If
                    
        Else
                
            If Not EsGm(tUser) Then
                If Not CanUserTelep(Map, tUser) Then Exit Sub

            End If
                
            If InMapBounds(Map, X, Y) Then
                If MapData(Map, X, Y).TileExit.Map = 0 Then
                    If UserList(tUser).PosAnt.Map <> Map Then
                        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Transportó a " & UserList(tUser).Name & " hacia " & "Mapa" & Map & " X:" & X & " Y:" & Y)

                    End If
                            
                    UserList(tUser).PosAnt.Map = UserList(tUser).Pos.Map
                    UserList(tUser).PosAnt.X = UserList(tUser).Pos.X
                    UserList(tUser).PosAnt.Y = UserList(tUser).Pos.Y
                                    
                    Call FindLegalPos(tUser, Map, X, Y)

                    If Map <> 0 And X <> 0 And Y <> 0 Then
                        Call WarpUserChar(tUser, Map, X, Y, True, True)

                    End If

                End If

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleWarpChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleWarpChar " & "at line " & Erl

    '</EhFooter>
End Sub

''
' Handles the "Silence" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSilence(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleSilence_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        
        If Not EsGmPriv(UserIndex) Then Exit Sub
        
        tUser = NameIndex(UserName)
        
        If tUser <= 0 Then
            Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
        Else

            If UserList(tUser).flags.Silenciado = 0 Then
                UserList(tUser).flags.Silenciado = 1
                Call WriteConsoleMsg(UserIndex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
                Call WriteShowMessageBox(tUser, "Estimado usuario, ud. ha sido silenciado por los administradores. Sus denuncias  y mensajes serán ignoradas por el servidor de aquí en más. Utilice /GM para contactar un administrador.")
                Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/silenciar " & UserList(tUser).Name)
                
                'Flush the other user's buffer
                Call FlushBuffer(tUser)
            Else
                UserList(tUser).flags.Silenciado = 0
                Call WriteConsoleMsg(UserIndex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
                Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/DESsilenciar " & UserList(tUser).Name)

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleSilence_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleSilence " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "GoToChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleGoToChar_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/03/2009
    '26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim X        As Integer

        Dim Y        As Integer
        
        Dim Rank     As PlayerType
        
        UserName = Reader.ReadString8()
        tUser = NameIndex(UserName)
        
        If Not EsGmDios(UserIndex) Then Exit Sub ' Comando único para Gm's
        If Not EsGmPriv(UserIndex) And EsAdmin(UserName) Then Exit Sub
              
        If tUser <= 0 Then
                   
            Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
        Else

            If Not EsGmDios(UserIndex) And MapInfo(UserList(tUser).Pos.Map).Pk Then Exit Sub
                 
            X = UserList(tUser).Pos.X
            Y = UserList(tUser).Pos.Y + 1
            Call FindLegalPos(UserIndex, UserList(tUser).Pos.Map, X, Y)
            Call WarpUserChar(UserIndex, UserList(tUser).Pos.Map, X, Y, True)
                    
            If .flags.AdminInvisible = 0 Then
                Call WriteConsoleMsg(tUser, .Name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                Call FlushBuffer(tUser)

            End If
                    
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.Map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y & " (" & MapInfo(UserList(tUser).Pos.Map).Name & ")")

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleGoToChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGoToChar " & "at line " & Erl

    '</EhFooter>
End Sub

''
' Handles the "Invisible" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleInvisible_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)

        If Not EsGm(UserIndex) Then Exit Sub
        
        Call DoAdminInvisible(UserIndex)
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/INVISIBLE")

    End With

    '<EhFooter>
    Exit Sub

HandleInvisible_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleInvisible " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleGMPanel_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If Not EsGmDios(UserIndex) Then Exit Sub
        
        Call WriteShowGMPanelForm(UserIndex)

    End With

    '<EhFooter>
    Exit Sub

HandleGMPanel_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGMPanel " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRequestUserList_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/09/07
    'Last modified by: Lucas Tavolaro Ortiz (Tavo)
    'I haven`t found a solution to split, so i make an array of names
    '***************************************************
    Dim i       As Long

    Dim names() As String

    Dim Count   As Long
    
    With UserList(UserIndex)
        
        If Not EsGmDios(UserIndex) Then Exit Sub
        
        ReDim names(1 To LastUser) As String
        Count = 1
        
        For i = 1 To LastUser

            If (LenB(UserList(i).Name) <> 0) Then
                If UserList(i).flags.Privilegios And PlayerType.User Then
                    names(Count) = UserList(i).Name
                    Count = Count + 1

                End If

            End If

        Next i
        
        If Count > 1 Then Call WriteUserNameList(UserIndex, names(), Count - 1)

    End With

    '<EhFooter>
    Exit Sub

HandleRequestUserList_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRequestUserList " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Jail" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleJail(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleJail_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim Reason   As String

        Dim jailTime As Byte

        Dim Count    As Byte

        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        Reason = Reader.ReadString8()
        jailTime = Reader.ReadInt()
        
        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")

        End If
        
        If Not EsGmDios(UserIndex) Then Exit Sub
        
        '/carcel nick@motivo@<tiempo>
        If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
        Else
            tUser = NameIndex(UserName)
                
            If tUser <= 0 Then
                If (EsDios(UserName) Or EsAdmin(UserName)) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                Else

                    If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                        Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(Reason) & " " & Date & " " & Time)
                        Call WriteVar(CharPath & UserName & ".chr", "COUNTERS", "Pena", jailTime)
                        Call WriteVar(CharPath & UserName & ".chr", "INIT", "Position", CStr(Prision.Map & "-" & Prision.X & "-" & Prision.Y))

                    End If
                        
                    Call WriteConsoleMsg(UserIndex, "El usuario ha sido enviado a la carcel estando OFFLINE.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                ElseIf jailTime > 60 Then
                    Call WriteConsoleMsg(UserIndex, "No puedés encarcelar por más de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
                Else

                    If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")

                    End If

                    If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")

                    End If
                        
                    If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                        Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(Reason) & " " & Date & " " & Time)

                    End If
                        
                    Call Encarcelar(tUser, jailTime, .Name)
                    Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, " encarceló a " & UserName)

                End If

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleJail_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleJail " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "KillNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleKillNPC_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/22/08 (NicoNZ)
    '
    '***************************************************
    With UserList(UserIndex)
        
        If Not EsGm(UserIndex) Then Exit Sub
        
        Dim tNpc   As Integer

        Dim auxNPC As Npc
        
        tNpc = .flags.TargetNPC
        
        If tNpc > 0 Then
            If isNPCResucitador(tNpc) Then
                Call DeleteAreaResuTheNpc(tNpc)

            End If
        
            Call WriteConsoleMsg(UserIndex, "RMatas (con posible respawn) a: " & Npclist(tNpc).Name, FontTypeNames.FONTTYPE_INFO)
            
            auxNPC = Npclist(tNpc)
            Call QuitarNPC(tNpc)
            Call RespawnNpc(auxNPC)
            
            .flags.TargetNPC = 0
        Else
            Call WriteConsoleMsg(UserIndex, "Antes debes hacer click sobre el NPC.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleKillNPC_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleKillNPC " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "WarnUser" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleWarnUser_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/26/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim Reason   As String

        Dim Privs    As PlayerType

        Dim Count    As Byte

        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        Reason = Reader.ReadString8()
        
        If EsGmDios(UserIndex) Then
            If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Else
                Privs = UserDarPrivilegioLevel(UserName)
                
                If Not Privs And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "No puedes advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
                Else

                    If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")

                    End If

                    If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")

                    End If
                    
                    If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                        Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": ADVERTENCIA por: " & LCase$(Reason) & " " & Date & " " & Time)
                        
                        tUser = NameIndex(UserName)
                        
                        If tUser > 0 Then
                            Call Encarcelar(tUser, 5)
                        Else
                            Call WriteVar(CharPath & UserName & ".chr", "COUNTERS", "Pena", "5")
                            Call WriteVar(CharPath & UserName & ".chr", "INIT", "Position", Prision.Map & "-" & Prision.X & "-" & Prision.Y)

                        End If
                        
                        Call WriteConsoleMsg(UserIndex, "Has advertido a " & UCase$(UserName) & ".", FontTypeNames.FONTTYPE_INFO)
                        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, " advirtio a " & UserName)

                    End If

                End If

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleWarnUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleWarnUser " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "RequestCharInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInfo(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRequestCharInfo_Err

    '</EhHeader>

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid).. alto bug zapallo..
    '***************************************************

    With UserList(UserIndex)
                
        Dim TargetName  As String

        Dim TargetIndex As Integer
        
        TargetName = Replace$(Reader.ReadString8(), "+", " ")
        TargetIndex = NameIndex(TargetName)
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then

            'is the player offline?
            If TargetIndex <= 0 Then

                'don't allow to retrieve administrator's info
                If Not (EsDios(TargetName) Or EsAdmin(TargetName)) Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, buscando en charfile.", FontTypeNames.FONTTYPE_INFO)
                          
                    If EsGmPriv(UserIndex) Then
                        Call SendUserStatsTxtOFF(UserIndex, TargetName)

                    End If

                End If

            Else

                Call SendUserStatsTxt(UserIndex, TargetIndex)

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleRequestCharInfo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRequestCharInfo " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "RequestCharInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInventory(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRequestCharInventory_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName         As String

        Dim tUser            As Integer
        
        Dim UserIsAdmin      As Boolean

        Dim OtherUserIsAdmin As Boolean
        
        UserName = Reader.ReadString8()
        
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/INV " & UserName)
            
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo del charfile...", FontTypeNames.FONTTYPE_TALK)
                    
                    Call SendUserInvTxtFromChar(UserIndex, UserName)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el inventario de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserInvTxt(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el inventario de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleRequestCharInventory_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRequestCharInventory " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "RequestCharBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharBank(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRequestCharBank_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName         As String

        Dim tUser            As Integer
        
        Dim UserIsAdmin      As Boolean

        Dim OtherUserIsAdmin As Boolean
        
        Dim TypeBank         As E_BANK
        
        UserName = Reader.ReadString8()
        TypeBank = Reader.ReadInt()
        
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin)) <> 0
        
        If UserIsAdmin Then
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/BOV " & UserName)
            
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                    Select Case TypeBank

                        Case E_BANK.e_User
                            Call SendUserBovedaTxtFromChar(UserIndex, UserName)

                        Case E_BANK.e_Account
                            Call SendUserBovedaTxtFromChar_Account(UserIndex, UserName)

                    End Select
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver la bóveda de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If UserIsAdmin Or Not OtherUserIsAdmin Then

                    Select Case TypeBank

                        Case E_BANK.e_User
                            Call SendUserBovedaTxt(UserIndex, tUser)

                        Case E_BANK.e_Account
                            Call SendUserBovedaTxt_Account(UserIndex, tUser)

                    End Select
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver la bóveda de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleRequestCharBank_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRequestCharBank " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ReviveChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReviveChar(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleReviveChar_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 11/03/2010
    '11/03/2010: ZaMa - Al revivir con el comando, si esta navegando le da cuerpo e barca.
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim LoopC    As Byte
        
        UserName = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
            Else
                tUser = UserIndex

            End If
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else

                With UserList(tUser)

                    If MapInfo(.Pos.Map).Pk Then Exit Sub
                         
                    'If dead, show him alive (naked).
                    If .flags.Muerto = 1 Then
                        .flags.Muerto = 0
                        
                        If .flags.Navegando = 1 Then
                            Call ToggleBoatBody(tUser)
                        Else
                            Call DarCuerpoDesnudo(tUser)

                        End If
                        
                        If .flags.Traveling = 1 Then
                            Call EndTravel(tUser, True)

                        End If
                        
                        Call ChangeUserChar(tUser, .Char.Body, .OrigChar.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
                        
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)

                    End If
                    
                    .Stats.MinHp = .Stats.MaxHp
                    
                    If .flags.Traveling = 1 Then
                        Call EndTravel(tUser, True)

                    End If
                    
                End With
                
                Call WriteUpdateHP(tUser)
                
                Call FlushBuffer(tUser)
                
                Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Resucito a " & UserName)

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleReviveChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleReviveChar " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "OnlineGM" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineGM(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleOnlineGM_Err

    '</EhHeader>

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 12/28/06
    '
    '***************************************************
    Dim i    As Long

    Dim List As String

    Dim Priv As PlayerType
    
    With UserList(UserIndex)
        
        If Not EsGm(UserIndex) Then Exit Sub

        For i = 1 To LastUser

            If UserList(i).flags.UserLogged Then
                If EsGm(i) And Not EsGmPriv(i) Then
                    List = List & UserList(i).Name & ", "

                End If

            End If

        Next i
        
        If LenB(List) <> 0 Then
            List = Left$(List, Len(List) - 2)
            Call WriteConsoleMsg(UserIndex, List & ".", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleOnlineGM_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleOnlineGM " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "OnlineMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineMap(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleOnlineMap_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 23/03/2009
    '23/03/2009: ZaMa - Ahora no requiere estar en el mapa, sino que por defecto se toma en el que esta, pero se puede especificar otro
    '***************************************************
    With UserList(UserIndex)
        
        Dim Map As Integer

        Map = Reader.ReadInt
        
        If Not EsGmPriv(UserIndex) Then Exit Sub
        
        Dim LoopC As Long

        Dim List  As String

        Dim Priv  As PlayerType
        
        For LoopC = 1 To LastUser

            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).Pos.Map = Map Then
                List = List & UserList(LoopC).Name & ", "

            End If

        Next LoopC
        
        If Len(List) > 2 Then List = Left$(List, Len(List) - 2)
        
        Call WriteConsoleMsg(UserIndex, "Usuarios en el mapa: " & List, FontTypeNames.FONTTYPE_INFO)
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/ONLINEMAP " & Map)

    End With

    '<EhFooter>
    Exit Sub

HandleOnlineMap_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleOnlineMap " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Forgive" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForgive(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleForgive_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName    As String

        Dim tUser       As Integer

        Dim ResetArmada As Boolean
        
        UserName = Reader.ReadString8()
        ResetArmada = Reader.ReadBool()
            
        If ResetArmada Then
            If Not (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.ChaosCouncil)) <> 0 Then Exit Sub
        Else

            If Not (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoyalCouncil)) <> 0 Then Exit Sub

        End If
            
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.ChaosCouncil Or PlayerType.ChaosCouncil)) <> 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If UserList(tUser).Faction.Status <> r_None Then
                    Call WriteConsoleMsg(UserIndex, "El personaje ya pertenece a alguna facción.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Has perdonado al personaje " & UserList(tUser).Name & ". " & IIf((ResetArmada = True), "Se reiniciaron Frags de Ciudadanos: EX VALOR: " & UserList(tUser).Faction.FragsCiu, vbNullString), FontTypeNames.FONTTYPE_INFOGREEN)
                    
                    If ResetArmada Then UserList(tUser).Faction.FragsCiu = 0

                    Call Faction_RemoveUser(tUser)
                    
                    Call LogPerdones("El GM " & .Name & " ha perdonado al personaje " & UserList(tUser).Name & ".")

                End If
                
            Else
                Call WriteConsoleMsg(UserIndex, "El personaje esta offline", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleForgive_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleForgive " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Kick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKick(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleKick_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim Rank     As Integer

        Dim IsAdmin  As Boolean
        
        Rank = PlayerType.Admin Or PlayerType.Dios
        
        UserName = Reader.ReadString8()
        IsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And PlayerType.SemiDios) Or IsAdmin Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                If Not (EsDios(UserName) Or EsAdmin(UserName)) Or IsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes echar a alguien con jerarquía mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If (UserList(tUser).flags.Privilegios And Rank) > (.flags.Privilegios And Rank) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes echar a alguien con jerarquía mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " echó a " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
                    Call WriteDisconnect(tUser)
                    Call FlushBuffer(tUser)
                        
                    Call CloseSocket(tUser)
                    Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Echó a " & UserName)

                End If

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleKick_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleKick " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Execute" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExecute(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleExecute_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "¿¿Estás loco?? ¿¿Cómo vas a piñatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)
                Else

                    If UserList(tUser).flags.Desafiando = 0 And UserList(tUser).flags.SlotReto = 0 And UserList(tUser).flags.SlotFast = 0 And UserList(tUser).flags.SlotEvent = 0 Then
                        
                        Call UserDie(tUser)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " ha ejecutado a " & UserName & ".", FontTypeNames.FONTTYPE_EJECUCION))
                        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, " ejecuto a " & UserName)
                    
                    Else
                        Call WriteConsoleMsg(UserIndex, "El usuario no puede ser ejecutado en este momento.", FontTypeNames.FONTTYPE_INFO)
                    
                    End If

                End If

            Else

                If Not (EsDios(UserName) Or EsAdmin(UserName)) Then
                    Call WriteConsoleMsg(UserIndex, "No está online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "¿¿Estás loco?? ¿¿Cómo vas a piñatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleExecute_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleExecute " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "BanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanChar(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleBanChar_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
        
    Dim UserName As String

    Dim Reason   As String
        
    Dim Tipo     As Byte
        
    Dim DataDay  As String
        
    UserName = Reader.ReadString8()
    Reason = Reader.ReadString8()
    Tipo = Reader.ReadInt()
    DataDay = Reader.ReadString8()
        
    If Not EsGmDios(UserIndex) Then Exit Sub
        
    Select Case Tipo

        Case 0 ' Baneo de personajes
            Call BanCharacter(UserIndex, UserName, Reason, DataDay)
            
        Case 1 ' Baneo de cuenta
            Call BanCharacter_Account(UserIndex, UserName, Reason, DataDay)

    End Select
    
    '<EhFooter>
    Exit Sub

HandleBanChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleBanChar " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "UnbanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanChar(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleUnbanChar_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName  As String

        Dim cantPenas As Byte
        
        UserName = Reader.ReadString8()
        
        If EsGmDios(UserIndex) Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")

            End If

            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")

            End If
            
            If Not FileExist(CharPath & UserName & ".chr", vbNormal) Then
                Call WriteConsoleMsg(UserIndex, "Charfile inexistente (no use +).", FontTypeNames.FONTTYPE_INFO)
            Else

                If (val(GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban")) = 1) Then
                    Call UnBan(UserName)
                
                    'penas
                    cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": UNBAN. " & Date & " " & Time)
                
                    Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/UNBAN a " & UserName)
                    Call WriteConsoleMsg(UserIndex, UserName & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & " no está baneado. Imposible unbanear.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleUnbanChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleUnbanChar " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "NPCFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNPCFollow(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleNPCFollow_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If Not EsGm(UserIndex) Then Exit Sub
        
        If .flags.TargetNPC > 0 Then
            Call DoFollow(.flags.TargetNPC, .Name)
            Npclist(.flags.TargetNPC).flags.Inmovilizado = 0
            Npclist(.flags.TargetNPC).flags.Paralizado = 0
            Npclist(.flags.TargetNPC).Contadores.Paralisis = 0

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleNPCFollow_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleNPCFollow " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "SummonChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSummonChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 26/03/2009
    '26/03/2009: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim X        As Integer

        Dim Y        As Integer
        
        Dim IsEvent  As Byte
        
        UserName = Reader.ReadString8()
        IsEvent = Reader.ReadBool()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                If (EsDios(UserName) Or EsAdmin(UserName)) And Not EsAdmin(.Name) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "El jugador no está online.", FontTypeNames.FONTTYPE_INFO)

                End If
                
            Else

                If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (UserList(tUser).flags.Privilegios And (PlayerType.User)) <> 0 Or (UserList(tUser).flags.Privilegios And (PlayerType.SemiDios)) <> 0 Then
                  
                    If Not IsEvent Then

                        ' Usuario participando en otro eventos
                        If Not CanUserTelep(.Pos.Map, tUser) Then
                            WriteConsoleMsg UserIndex, "El personaje no está disponible para ser sumoneado.", FontTypeNames.FONTTYPE_INFO
                            Exit Sub
        
                        End If

                        If MapInfo(.Pos.Map).Pk Then
                            WriteConsoleMsg UserIndex, "El personaje no está disponible para ser sumoneado.", FontTypeNames.FONTTYPE_INFO
                            Exit Sub
        
                        End If
        
                        UserList(tUser).PosAnt.Map = UserList(tUser).Pos.Map
                        UserList(tUser).PosAnt.X = UserList(tUser).Pos.X
                        UserList(tUser).PosAnt.Y = UserList(tUser).Pos.Y
                                
                        Call WriteConsoleMsg(tUser, .Name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
                        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/SUM " & UserName & " Map:" & .Pos.Map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                    Else

                        If Not UserList(tUser).flags.SlotEvent > 0 Then Exit Sub    ' @ Si no está en evento no puede usar el comando, capaz tenga que refresh.

                    End If
                          
                    X = .Pos.X
                    Y = .Pos.Y + 1
                    Call FindLegalPos(tUser, .Pos.Map, X, Y)
                    Call WarpUserChar(tUser, .Pos.Map, X, Y, True, True)

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
    End With
    
End Sub

''
' Handles the "SpawnListRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnListRequest(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleSpawnListRequest_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User) Then Exit Sub
        
        Call EnviarSpawnList(UserIndex)

    End With

    '<EhFooter>
    Exit Sub

HandleSpawnListRequest_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleSpawnListRequest " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "SpawnCreature" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnCreature(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleSpawnCreature_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Npc      As Integer

        Dim NpcIndex As Integer
            
        Npc = Reader.ReadInt()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then

            If MapInfo(.Pos.Map).Pk Then Exit Sub
                  
            If Npc > 0 And Npc <= UBound(Declaraciones.SpawnList()) Then
                    
                NpcIndex = SpawnNpc(Declaraciones.SpawnList(Npc).NpcIndex, .Pos, True, False)

                Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Sumoneo " & Declaraciones.SpawnList(Npc).NpcName)

            End If

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleSpawnCreature_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleSpawnCreature " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ResetNPCInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResetNPCInventory(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleResetNPCInventory_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If Not EsGm(UserIndex) Then Exit Sub
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call ResetNpcInv(.flags.TargetNPC)
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/RESETINV " & Npclist(.flags.TargetNPC).Name)

    End With

    '<EhFooter>
    Exit Sub

HandleResetNPCInventory_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleResetNPCInventory " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "CleanWorld" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCleanWorld(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleCleanWorld_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)

        If Not EsGm(UserIndex) Then Exit Sub
        
        Call LimpiarMundo

        'CountDownLimpieza = 5
    End With

    '<EhFooter>
    Exit Sub

HandleCleanWorld_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCleanWorld " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ServerMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerMessage(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleServerMessage_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 28/05/2010
    '28/05/2010: ZaMa - Ahora no dice el nombre del gm que lo dice.
    '***************************************************

    With UserList(UserIndex)
        
        Dim Message As String

        Message = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(Message) <> 0 Then
                Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Mensaje Broadcast:" & Message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & "» " & Message, FontTypeNames.FONTTYPE_RMSG, eMessageType.Admin))

                ''''''''''''''''SOLO PARA EL TESTEO'''''''
                ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
                'frmMain.txtChat.Text = frmMain.txtChat.Text & vbNewLine & UserList(UserIndex).name & " > " & message
            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleServerMessage_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleServerMessage " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "MapMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMapMessage(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleMapMessage_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    '***************************************************

    With UserList(UserIndex)
        
        Dim Message As String

        Message = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(Message) <> 0 Then
                
                Dim mapa As Integer

                mapa = .Pos.Map
                
                Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Mensaje a mapa " & mapa & ":" & Message)
                Call SendData(SendTarget.toMap, mapa, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_TALK))

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleMapMessage_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleMapMessage " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "NickToIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNickToIP(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleNickToIP_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    'Pablo (ToxicWaste): Agrego para que el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim Priv     As PlayerType

        Dim IsAdmin  As Boolean
        
        UserName = Reader.ReadString8()
        
        If EsGmPriv(UserIndex) Then
            tUser = NameIndex(UserName)
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "NICK2IP Solicito la IP de " & UserName)
            
            IsAdmin = (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0

            If IsAdmin Then
                Priv = PlayerType.User Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            Else
                Priv = PlayerType.User

            End If
            
            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And Priv Then
                    Call WriteConsoleMsg(UserIndex, "El ip de " & UserName & " es " & UserList(tUser).IpAddress, FontTypeNames.FONTTYPE_INFO)

                    Dim IP    As String

                    Dim lista As String

                    Dim LoopC As Long

                    IP = UserList(tUser).IpAddress

                    For LoopC = 1 To LastUser

                        If UserList(LoopC).IpAddress = IP Then
                            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                                If UserList(LoopC).flags.Privilegios And Priv Then
                                    lista = lista & UserList(LoopC).Name & ", "

                                End If

                            End If

                        End If

                    Next LoopC

                    If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                    Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & IP & " son: " & lista, FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If Not (EsDios(UserName) Or EsAdmin(UserName)) Or IsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "No hay ningún personaje con ese nick.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleNickToIP_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleNickToIP " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "IPToNick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleIPToNick(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleIPToNick_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim IP    As String

        Dim LoopC As Long

        Dim lista As String

        Dim Priv  As PlayerType
        
        IP = Reader.ReadInt() & "."
        IP = IP & Reader.ReadInt() & "."
        IP = IP & Reader.ReadInt() & "."
        IP = IP & Reader.ReadInt()
        
        If Not EsGmPriv(UserIndex) Then Exit Sub
        
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "IP2NICK Solicito los Nicks de IP " & IP)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            Priv = PlayerType.User Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
        Else
            Priv = PlayerType.User

        End If

        For LoopC = 1 To LastUser

            If UserList(LoopC).IpAddress = IP Then
                If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).flags.Privilegios And Priv Then
                        lista = lista & UserList(LoopC).Name & ", "

                    End If

                End If

            End If

        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & IP & " son: " & lista, FontTypeNames.FONTTYPE_INFO)

    End With

    '<EhFooter>
    Exit Sub

HandleIPToNick_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleIPToNick " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "TeleportCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportCreate(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleTeleportCreate_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 22/03/2010
    '15/11/2009: ZaMa - Ahora se crea un teleport con un radio especificado.
    '22/03/2010: ZaMa - Harcodeo los teleps y radios en el dat, para evitar mapas bugueados.
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim mapa  As Integer

        Dim X     As Byte

        Dim Y     As Byte

        Dim Radio As Byte
        
        mapa = Reader.ReadInt()
        X = Reader.ReadInt()
        Y = Reader.ReadInt()
        Radio = Reader.ReadInt()
        
        Radio = MinimoInt(Radio, 6)
        
        If Not EsGm(UserIndex) Then Exit Sub
  
        If Not MapaValido(mapa) Or Not InMapBounds(mapa, X, Y) Then Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).TileExit.Map > 0 Then Exit Sub

        If Not EsGmPriv(UserIndex) Then
            If MapInfo(.Pos.Map).Pk Then Exit Sub
    
            ' Crea con destino inseguro y es semi dios
            If Not EsGmDios(UserIndex) Then
                If MapInfo(mapa).Pk Then Exit Sub

            End If

        End If
            
        If MapData(mapa, X, Y).ObjInfo.ObjIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Hay un objeto en el piso en ese lugar.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If MapData(mapa, X, Y).TileExit.Map > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        Dim ET As Obj

        ET.Amount = 1
        ' Es el numero en el dat. El indice es el comienzo + el radio, todo harcodeado :(.
        ET.ObjIndex = TELEP_OBJ_INDEX 'TELEP_OBJ_INDEX + Radio
        
        With MapData(.Pos.Map, .Pos.X, .Pos.Y - 1)
            .TileExit.Map = mapa
            .TileExit.X = X
            .TileExit.Y = Y

        End With
        
        Call MakeObj(ET, .Pos.Map, .Pos.X, .Pos.Y - 1)
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/CT " & mapa & "," & X & "," & Y & "," & Radio)

    End With

    '<EhFooter>
    Exit Sub

HandleTeleportCreate_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleTeleportCreate " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "TeleportDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportDestroy(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleTeleportDestroy_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)

        Dim mapa As Integer

        Dim X    As Byte

        Dim Y    As Byte
        
        '/dt
            
        If Not EsGm(UserIndex) Then Exit Sub
            
        mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY
        
        If Not InMapBounds(mapa, X, Y) Then Exit Sub
        
        With MapData(mapa, X, Y)

            If .ObjInfo.ObjIndex = 0 Then Exit Sub
            
            If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And .TileExit.Map > 0 Then
                Call Logs_User(UserList(UserIndex).Name, eLog.eGm, eLogDescUser.eNone, "/DT: " & mapa & "," & X & "," & Y)
                
                Call EraseObj(.ObjInfo.Amount, mapa, X, Y)
                
                If MapData(.TileExit.Map, .TileExit.X, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
                    Call EraseObj(1, .TileExit.Map, .TileExit.X, .TileExit.Y)

                End If
                
                .TileExit.Map = 0
                .TileExit.X = 0
                .TileExit.Y = 0

            End If

        End With

    End With

    '<EhFooter>
    Exit Sub

HandleTeleportDestroy_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleTeleportDestroy " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "EnableDenounces" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnableDenounces(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleEnableDenounces_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    'Enables/Disables
    '***************************************************

    With UserList(UserIndex)
        
        If Not EsGmPriv(UserIndex) Then Exit Sub
        
        Dim Activado As Boolean

        Dim msg      As String
        
        Activado = Not .flags.SendDenounces
        .flags.SendDenounces = Activado
        
        msg = "Denuncias por consola " & IIf(Activado, "activadas", "desactivadas") & "."
        
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, msg)
        
        Call WriteConsoleMsg(UserIndex, msg, FontTypeNames.FONTTYPE_INFO)

    End With

    '<EhFooter>
    Exit Sub

HandleEnableDenounces_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleEnableDenounces " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ShowDenouncesList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowDenouncesList(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleShowDenouncesList_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowDenounces(UserIndex)

    End With

    '<EhFooter>
    Exit Sub

HandleShowDenouncesList_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleShowDenouncesList " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ForceMIDIToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HanldeForceMIDIToMap(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HanldeForceMIDIToMap_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        Dim midiID As Byte

        Dim mapa   As Integer
        
        midiID = Reader.ReadInt
        mapa = Reader.ReadInt
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then

            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, 50, 50) Then
                mapa = .Pos.Map

            End If
        
            If midiID = 0 Then
                'Ponemos el default del mapa
                'Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayMusic(MapInfo(.Pos.Map).Music))
            Else

                'Ponemos el pedido por el GM
                'Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayMusic(midiID))
            End If

        End If

    End With

    '<EhFooter>
    Exit Sub

HanldeForceMIDIToMap_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HanldeForceMIDIToMap " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ForceWAVEToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEToMap(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleForceWAVEToMap_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim waveID As Integer

        Dim mapa   As Integer

        Dim X      As Byte

        Dim Y      As Byte
        
        waveID = Reader.ReadInt()
        mapa = Reader.ReadInt()
        X = Reader.ReadInt()
        Y = Reader.ReadInt()
        
        'Solo dioses, admins y RMS
        If EsGmDios(UserIndex) Then

            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, X, Y) Then
                mapa = .Pos.Map
                X = .Pos.X
                Y = .Pos.Y

            End If
            
            'Ponemos el pedido por el GM
            'Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayEffect(waveID, X, Y))
        End If

    End With

    '<EhFooter>
    Exit Sub

HandleForceWAVEToMap_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleForceWAVEToMap " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "RoyalArmyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyMessage(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRoyalArmyMessage_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim Message As String

        Message = Reader.ReadString8()
        
        'Solo dioses, admins, semis y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoyalCouncil) Then
            Call SendData(SendTarget.ToCiudadanos, 0, PrepareMessageConsoleMsg("[Consejo de Banderbill] " & .Name & "> " & Message, FontTypeNames.FONTTYPE_CONSEJOVesA))
        
        Else

            If .Faction.Status = r_Armada Then
                Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("[Armada Real] " & .Name & "> " & Message, FontTypeNames.FONTTYPE_INFOGREEN))

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleRoyalArmyMessage_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRoyalArmyMessage " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ChaosLegionMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionMessage(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChaosLegionMessage_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim Message As String

        Message = Reader.ReadString8()
        
        'Solo dioses, admins, concilios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.ChaosCouncil) Then
            Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("[Concilio de las Sombras] " & .Name & "> " & Message, FontTypeNames.FONTTYPE_EJECUCION))
        Else
            
            If .Faction.Status = r_Caos Then
                Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("[Legión Oscura] " & .Name & "> " & Message, FontTypeNames.FONTTYPE_INFORED))

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleChaosLegionMessage_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChaosLegionMessage " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "TalkAsNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalkAsNPC(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleTalkAsNPC_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim Message As String

        Message = Reader.ReadString8()
        
        ' Solo dioses, admins y RMS
        If EsGmPriv(UserIndex) Then

            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(Message, Npclist(.flags.TargetNPC).Char.charindex, vbWhite))
            Else
                Call WriteConsoleMsg(UserIndex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleTalkAsNPC_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleTalkAsNPC " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "DestroyAllItemsInArea" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyAllItemsInArea(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleDestroyAllItemsInArea_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If Not EsGm(UserIndex) Then Exit Sub

        Dim X       As Long

        Dim Y       As Long

        Dim bIsExit As Boolean
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1

                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex > 0 Then
                        bIsExit = MapData(.Pos.Map, X, Y).TileExit.Map > 0

                        If ItemNoEsDeMapa(.Pos.Map, X, Y, bIsExit) Then
                            Call EraseObj(MAX_INVENTORY_OBJS, .Pos.Map, X, Y)

                        End If

                    End If

                End If

            Next X
        Next Y
        
        Call Logs_User(UserList(UserIndex).Name, eLog.eGm, eLogDescUser.eNone, "/MASSDEST")

    End With

    '<EhFooter>
    Exit Sub

HandleDestroyAllItemsInArea_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleDestroyAllItemsInArea " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "AcceptRoyalCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptRoyalCouncilMember(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleAcceptRoyalCouncilMember_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim LoopC    As Byte
        
        UserName = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))

                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                    If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil
                    
                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)

                End With

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleAcceptRoyalCouncilMember_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleAcceptRoyalCouncilMember " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ChaosCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptChaosCouncilMember(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleAcceptChaosCouncilMember_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim LoopC    As Byte
        
        UserName = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))
                
                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                    If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil

                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)

                End With

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleAcceptChaosCouncilMember_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleAcceptChaosCouncilMember " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ItemsInTheFloor" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleItemsInTheFloor(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleItemsInTheFloor_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.SemiDios) Then Exit Sub
        
        Dim tobj  As Integer

        Dim lista As String

        Dim X     As Long

        Dim Y     As Long
        
        For X = 5 To 95
            For Y = 5 To 95
                tobj = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex

                If tobj > 0 Then
                    If ObjData(tobj).OBJType <> eOBJType.otArboles Then
                        Call WriteConsoleMsg(UserIndex, "(" & X & "," & Y & ") " & ObjData(tobj).Name, FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            Next Y
        Next X

    End With

    '<EhFooter>
    Exit Sub

HandleItemsInTheFloor_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleItemsInTheFloor " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "CouncilKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilKick(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleCouncilKick_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, echando de los consejos.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECE", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECECAOS", 0)
                Else
                    Call WriteConsoleMsg(UserIndex, "No se encuentra el charfile " & CharPath & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Banderbill.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))

                    End If
                    
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))

                    End If

                End With

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleCouncilKick_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCouncilKick " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "SetTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetTrigger(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleSetTrigger_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim tTrigger As Byte

        Dim tLog     As String

        Dim ObjIndex As Integer
        
        tTrigger = Reader.ReadInt()
        
        If Not EsGmDios(UserIndex) Then Exit Sub
        
        If tTrigger >= 0 Then
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.zonaOscura Then
                If tTrigger <> eTrigger.zonaOscura Then
                    If Not (.flags.AdminInvisible = 1) Then Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False))
                    
                    ObjIndex = MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex
                    
                    If ObjIndex > 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageObjectCreate(ObjIndex, ObjData(ObjIndex).GrhIndex, .Pos.X, .Pos.Y, vbNullString, 0, ObjData(ObjIndex).Sound))

                    End If

                End If

            Else

                If tTrigger = eTrigger.zonaOscura Then
                    If Not (.flags.AdminInvisible = 1) Then Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True))
                    
                    ObjIndex = MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex
                    
                    If ObjIndex > 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageObjectDelete(.Pos.X, .Pos.Y))

                    End If

                End If

            End If
            
            MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = tTrigger
            tLog = "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & "," & .Pos.Y
            
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, tLog)
            Call WriteConsoleMsg(UserIndex, tLog, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleSetTrigger_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleSetTrigger " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "AskTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAskTrigger(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleAskTrigger_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 04/13/07
    '
    '***************************************************
    Dim tTrigger As Byte
    
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.SemiDios) Then Exit Sub
        
        tTrigger = MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger
        
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Miro el trigger en " & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)
        
        Call WriteConsoleMsg(UserIndex, "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_INFO)

    End With

    '<EhFooter>
    Exit Sub

HandleAskTrigger_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleAskTrigger " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "BannedIPList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleBannedIPList_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If Not EsGmPriv(UserIndex) Then Exit Sub
        
        Dim lista As String

        Dim LoopC As Long
        
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/BANIPLIST")
        
        For LoopC = 1 To BanIps.Count
            lista = lista & BanIps.Item(LoopC) & ", "
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        
        Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)

    End With

    '<EhFooter>
    Exit Sub

HandleBannedIPList_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleBannedIPList " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleBannedIPReload_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If Not EsGmPriv(UserIndex) Then Exit Sub
        
        Call BanIpGuardar
        Call BanIpCargar

    End With

    '<EhFooter>
    Exit Sub

HandleBannedIPReload_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleBannedIPReload " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "BanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleBanIP_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 07/02/09
    'Agregado un CopyBuffer porque se producia un bucle
    'inifito al intentar banear una ip ya baneada. (NicoNZ)
    '07/02/09 Pato - Ahora no es posible saber si un gm está o no online.
    '***************************************************

    With UserList(UserIndex)
        
        Dim bannedIP As String

        Dim tUser    As Integer

        Dim Reason   As String

        Dim i        As Long
        
        ' Is it by ip??
        If Reader.ReadBool() Then
            bannedIP = Reader.ReadInt() & "."
            bannedIP = bannedIP & Reader.ReadInt() & "."
            bannedIP = bannedIP & Reader.ReadInt() & "."
            bannedIP = bannedIP & Reader.ReadInt()
        Else
            tUser = NameIndex(Reader.ReadString8())
            
            If tUser > 0 Then bannedIP = UserList(tUser).IpAddress

        End If
        
        Reason = Reader.ReadString8()
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If LenB(bannedIP) > 0 Then
                Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/BanIP " & bannedIP & " por " & Reason)
                
                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call BanIpAgrega(bannedIP)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " baneó la IP " & bannedIP & " por " & Reason, FontTypeNames.FONTTYPE_FIGHT))
                    
                    'Find every player with that ip and ban him!
                    For i = 1 To LastUser

                        If UserList(i).ConnIDValida Then
                            If UserList(i).IpAddress = bannedIP Then
                                Call BanCharacter(UserIndex, UserList(i).Name, "IP POR " & Reason)

                            End If

                        End If

                    Next i

                End If

            ElseIf tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje no está online.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleBanIP_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleBanIP " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "UnbanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanIP(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleUnbanIP_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim bannedIP As String
        
        bannedIP = Reader.ReadInt() & "."
        bannedIP = bannedIP & Reader.ReadInt() & "."
        bannedIP = bannedIP & Reader.ReadInt() & "."
        bannedIP = bannedIP & Reader.ReadInt()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.SemiDios) Then Exit Sub
        
        If BanIpQuita(bannedIP) Then
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleUnbanIP_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleUnbanIP " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "CreateItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateItem(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleCreateItem_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    With UserList(UserIndex)

        Dim tobj As Integer

        Dim tStr As String

        tobj = Reader.ReadInt()

        If Not EsGmPriv(UserIndex) Then Exit Sub
            
        Dim mapa As Integer

        Dim X    As Byte

        Dim Y    As Byte
        
        mapa = .Pos.Map
        X = .Pos.X
        Y = .Pos.Y
            
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/CI: " & tobj & " en mapa " & mapa & " (" & X & "," & Y & ")")
        
        If MapData(mapa, X, Y - 1).ObjInfo.ObjIndex > 0 Then Exit Sub
        
        If MapData(mapa, X, Y - 1).TileExit.Map > 0 Then Exit Sub
        
        If tobj < 1 Or tobj > NumObjDatas Then Exit Sub
        
        'Is the object not null?
        If LenB(ObjData(tobj).Name) = 0 Then Exit Sub
                
        If Not EsGmPriv(UserIndex) Then
                
            'Silla
            'Trono
            'Sillon
            'Silla
                    
            If tobj <> 882 And tobj <> 162 And tobj <> 168 And tobj <> 826 Then
                        
                Exit Sub
                        
            End If

        End If
                
        Dim Objeto As Obj
            
        'NoCrear = 1
            
        Call WriteConsoleMsg(UserIndex, "¡¡ATENCIÓN: FUERON CREADOS ***25*** ÍTEMS, TIRE Y /DEST LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
        
        Objeto.Amount = 25
        Objeto.ObjIndex = tobj
        Call MakeObj(Objeto, mapa, X, Y - 1)

        Call Logs_User(.Name, eGm, eNone, "/CI: [" & tobj & "]" & ObjData(tobj).Name & " en mapa " & mapa & " (" & X & "," & Y & ")")
        
    End With

    '<EhFooter>
    Exit Sub

HandleCreateItem_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCreateItem " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "DestroyItems" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyItems(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleDestroyItems_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If Not EsGm(UserIndex) Then Exit Sub
        
        Dim mapa As Integer

        Dim X    As Byte

        Dim Y    As Byte
        
        mapa = .Pos.Map
        X = .Pos.X
        Y = .Pos.Y
        
        Dim ObjIndex As Integer

        ObjIndex = MapData(mapa, X, Y).ObjInfo.ObjIndex
        
        If ObjIndex = 0 Then Exit Sub
        
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/DEST " & ObjIndex & " en mapa " & mapa & " (" & X & "," & Y & "). Cantidad: " & MapData(mapa, X, Y).ObjInfo.Amount)
        
        If ObjData(ObjIndex).OBJType = eOBJType.otTeleport And MapData(mapa, X, Y).TileExit.Map > 0 Then
            
            Call WriteConsoleMsg(UserIndex, "No puede destruir teleports así. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        Call EraseObj(10000, mapa, X, Y)

    End With

    '<EhFooter>
    Exit Sub

HandleDestroyItems_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleDestroyItems " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ChaosLegionKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionKick(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChaosLegionKick_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.ChaosCouncil)) <> 0 Or .flags.PrivEspecial Then
            
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")

            End If

            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")

            End If

            tUser = NameIndex(UserName)
            
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "ECHO DEL CAOS A: " & UserName)
    
            If tUser > 0 Then
                If .Faction.Status > 0 Then
                    Call mFacciones.Faction_RemoveUser(tUser)
                    Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas del caos.", FontTypeNames.FONTTYPE_FIGHT)
                    Call FlushBuffer(tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "Solicita la expulsión a los superiores. El personaje no pertenece a ninguna facción.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If FileExist(CharPath & UserName & ".chr") Then
                
                    Dim Status As Byte

                    Status = val(GetVar(CharPath & UserName & ".chr", "FACTION", "STATUS"))
                    
                    If Status > 0 Then
                        Call WriteVar(CharPath & UserName & ".chr", "FACTION", "STATUS", "0")
                        Call WriteVar(CharPath & UserName & ".chr", "FACTION", "EXFACTION", CStr(Status))
                        Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Solicita la expulsión a los superiores. El personaje no pertenece a ninguna facción.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleChaosLegionKick_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChaosLegionKick " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "RoyalArmyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyKick(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRoyalArmyKick_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoyalCouncil)) <> 0 Or .flags.PrivEspecial Then
            
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")

            End If

            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")

            End If

            tUser = NameIndex(UserName)
            
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "ECHÓ DE LA REAL A: " & UserName)
            
            If tUser > 0 Then
            
                If .Faction.Status > 0 Then
                    Call mFacciones.Faction_RemoveUser(tUser)
                    Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas reales.", FontTypeNames.FONTTYPE_FIGHT)
                    Call FlushBuffer(tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "Solicita la expulsión a los superiores. El personaje no pertenece a ninguna facción.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
            
                If FileExist(CharPath & UserName & ".chr") Then
                
                    Dim Status As Byte

                    Status = val(GetVar(CharPath & UserName & ".chr", "FACTION", "STATUS"))
                    
                    If Status > 0 Then
                        Call WriteVar(CharPath & UserName & ".chr", "FACTION", "STATUS", 0)
                        Call WriteVar(CharPath & UserName & ".chr", "FACTION", "EXFACTION", CStr(Status))
                        Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Solicita la expulsión a los superiores. El personaje no pertenece a ninguna facción.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleRoyalArmyKick_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRoyalArmyKick " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ForceMIDIAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceMIDIAll(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleForceMIDIAll_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    With UserList(UserIndex)

        Dim midiID As Byte

        midiID = Reader.ReadInt()
        
        If Not EsGmPriv(UserIndex) Then Exit Sub
        
        'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " broadcast música: " & midiID, FontTypeNames.FONTTYPE_SERVER))
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMusic(midiID))

    End With

    '<EhFooter>
    Exit Sub

HandleForceMIDIAll_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleForceMIDIAll " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ForceWAVEAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEAll(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleForceWAVEAll_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    With UserList(UserIndex)

        Dim waveID As Byte

        waveID = Reader.ReadInt()
        
        If Not EsGmPriv(UserIndex) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayEffect(waveID, NO_3D_SOUND, NO_3D_SOUND))

    End With

    '<EhFooter>
    Exit Sub

HandleForceWAVEAll_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleForceWAVEAll " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "TileBlockedToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTileBlockedToggle(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleTileBlockedToggle_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)

        If Not EsGm(UserIndex) Then Exit Sub
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/BLOQ")
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 1
        Else
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0

        End If
        
        Call Bloquear(True, .Pos.Map, .Pos.X, .Pos.Y, MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked)

    End With

    '<EhFooter>
    Exit Sub

HandleTileBlockedToggle_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleTileBlockedToggle " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "KillNPCNoRespawn" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPCNoRespawn(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleKillNPCNoRespawn_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If Not EsGm(UserIndex) Then Exit Sub
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Pretoriano Then Exit Sub
        
        If isNPCResucitador(.flags.TargetNPC) Then
            Call DeleteAreaResuTheNpc(.flags.TargetNPC)

        End If
        
        Call QuitarNPC(.flags.TargetNPC)
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/MATA " & Npclist(.flags.TargetNPC).Name)

    End With

    '<EhFooter>
    Exit Sub

HandleKillNPCNoRespawn_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleKillNPCNoRespawn " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleKillAllNearbyNPCs_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If Not EsGmDios(UserIndex) Then Exit Sub
        
        Dim X As Long

        Dim Y As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1

                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(.Pos.Map, X, Y).NpcIndex)

                End If

            Next X
        Next Y

        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/MASSKILL")

    End With

    '<EhFooter>
    Exit Sub

HandleKillAllNearbyNPCs_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleKillAllNearbyNPCs " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "LastIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLastIP(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleLastIP_Err

    '</EhHeader>

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName   As String

        Dim lista      As String

        Dim LoopC      As Byte

        Dim Priv       As Integer

        Dim validCheck As Boolean
        
        Priv = PlayerType.Admin Or PlayerType.Dios
        UserName = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin)) <> 0 Then

            'Handle special chars
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")

            End If

            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "/", "")

            End If

            If (InStrB(UserName, "+") <> 0) Then
                UserName = Replace(UserName, "+", " ")

            End If
            
            'Only Gods and Admins can see the ips of adminsitrative characters. All others can be seen by every adminsitrative char.
            If NameIndex(UserName) > 0 Then
                validCheck = (UserList(NameIndex(UserName)).flags.Privilegios And Priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            Else
                validCheck = (UserDarPrivilegioLevel(UserName) And Priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0

            End If
            
            If validCheck Then
                Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/LASTIP " & UserName)
                
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    lista = "Las ultimas IPs con las que " & UserName & " se conectó son:"

                    For LoopC = 1 To 5
                        lista = lista & vbCrLf & LoopC & " - " & GetVar(CharPath & UserName & ".chr", "INIT", "LastIP" & LoopC)
                    Next LoopC

                    Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Charfile """ & UserName & """ inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                Call WriteConsoleMsg(UserIndex, UserName & " es de mayor jerarquía que vos.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleLastIP_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleLastIP " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "ChatColor" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleChatColor(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChatColor_Err

    '</EhHeader>

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Change the user`s chat color
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Color As Long
        
        Color = RGB(Reader.ReadInt(), Reader.ReadInt(), Reader.ReadInt())
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            .flags.ChatColor = Color

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleChatColor_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChatColor " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Ignored" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIgnored(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleIgnored_Err

    '</EhHeader>

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Ignore the user
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            .flags.AdminPerseguible = Not .flags.AdminPerseguible

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleIgnored_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleIgnored " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handle the "SaveChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleSaveChars_Err

    '</EhHeader>

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Save the characters
    '***************************************************
    With UserList(UserIndex)
        
        If Not EsGmDios(UserIndex) Then Exit Sub
        
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha guardado todos los chars.")
        
        Call DistributeExpAndGldGroups
        Call GuardarUsuarios(False)

    End With

    '<EhFooter>
    Exit Sub

HandleSaveChars_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleSaveChars " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handle the "ChangeMapInfoBackup" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeMapInfoBackup_Err

    '</EhHeader>

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Change the backup`s info of the map
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim doTheBackUp As Boolean
        
        doTheBackUp = Reader.ReadBool()
        
        If (.flags.Privilegios And (PlayerType.Admin)) = 0 Then Exit Sub
        
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre el BackUp.")
        
        'Change the boolean to byte in a fast way
        If doTheBackUp Then
            MapInfo(.Pos.Map).BackUp = 1
        Else
            MapInfo(.Pos.Map).BackUp = 0

        End If
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "backup", MapInfo(.Pos.Map).BackUp)
        
        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Backup: " & MapInfo(.Pos.Map).BackUp, FontTypeNames.FONTTYPE_INFO)

    End With

    '<EhFooter>
    Exit Sub

HandleChangeMapInfoBackup_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeMapInfoBackup " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handle the "ChangeMapInfoPK" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoPK(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeMapInfoPK_Err

    '</EhHeader>

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Change the pk`s info of the  map
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim isMapPk As Boolean
        
        isMapPk = Reader.ReadBool()
        
        If Not EsGmDios(UserIndex) Then Exit Sub
        
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si es PK el mapa.")
        
        MapInfo(.Pos.Map).Pk = isMapPk
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "Pk", IIf(isMapPk, "1", "0"))

        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " PK: " & MapInfo(.Pos.Map).Pk, FontTypeNames.FONTTYPE_INFO)

    End With

    '<EhFooter>
    Exit Sub

HandleChangeMapInfoPK_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeMapInfoPK " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleChangeMapInfoLvl(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeMapInfoLvl_Err

    '</EhHeader>

    '***************************************************
    'Author:
    'Last Modification:
    'Restringido de Nivel -> Options: Todo nivel disponible.
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Elv As Byte
        
        Elv = Reader.ReadInt()

        If EsGmDios(UserIndex) Then
            If (Elv > STAT_MAXELV + 1) Then
                Call WriteConsoleMsg(UserIndex, "El nivel máximo que puedes elegir es el máximo del juego +1", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si es restringido por nivel el mapa.")
                
            MapInfo(UserList(UserIndex).Pos.Map).LvlMin = Elv
                
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "LvlMin", Elv)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " RestringidoNivel: " & MapInfo(.Pos.Map).LvlMin, FontTypeNames.FONTTYPE_INFO)

        End If
        
    End With

    '<EhFooter>
    Exit Sub

HandleChangeMapInfoLvl_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeMapInfoLvl " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleChangeMapInfoLimpieza(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeMapInfoLimpieza_Err

    '</EhHeader>

    '***************************************************
    'Author:
    'Last Modification:
    'Restringido de Limpieza -> Options: Si/No
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Value As Byte
        
        Value = Reader.ReadInt()

        If EsGmPriv(UserIndex) Then
            
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si afecta la limpieza en el mapa o no")
                
            MapInfo(UserList(UserIndex).Pos.Map).Limpieza = Value
                
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Limpieza", Value)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Limpieza: " & IIf((MapInfo(.Pos.Map).Limpieza = 1), "SI", "NO"), FontTypeNames.FONTTYPE_INFO)

        End If
        
    End With

    '<EhFooter>
    Exit Sub

HandleChangeMapInfoLimpieza_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeMapInfoLimpieza " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleChangeMapInfoItems(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeMapInfoItems_Err

    '</EhHeader>

    '***************************************************
    'Author:
    'Last Modification:
    'Restringido de Items -> Options: Si/No
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Value As Byte
        
        Value = Reader.ReadInt()

        If EsGmDios(UserIndex) Then
            
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si caen items o no")
                
            MapInfo(UserList(UserIndex).Pos.Map).CaenItems = Value
                
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "CaenItems", Value)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Caen Items: " & IIf((MapInfo(.Pos.Map).CaenItems = 1), "SI", "NO"), FontTypeNames.FONTTYPE_INFO)

        End If
        
    End With

    '<EhFooter>
    Exit Sub

HandleChangeMapInfoItems_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeMapInfoItems " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleChangeMapInfoExp(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeMapInfoExp_Err

    '</EhHeader>

    Dim Exp As Single
    
    Exp = Reader.ReadReal32
    
    With UserList(UserIndex)

        If EsGmPriv(UserIndex) Then
            If Exp = 255 Then
                Call CheckHappyHour
                'frmMain.chkHappy.Value = IIf(HappyHour = True, 1, 0)
            ElseIf Exp = 254 Then
                Call CheckPartyTime
                'frmMain.chkParty.Value = IIf(PartyTime = True, 1, 0)
            Else
                
                MapInfo(.Pos.Map).Exp = Exp
                
                If Exp > 0 Then
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Experiencia aumentada en " & MapInfo(.Pos.Map).Name & " (" & .Pos.Map & ")" & " x" & CStr(Exp), FontTypeNames.FONTTYPE_USERPREMIUM))
                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("La experiencia de " & MapInfo(.Pos.Map).Name & " (" & .Pos.Map & ")" & " ha vuelto a la normalidad.", FontTypeNames.FONTTYPE_USERPREMIUM))

                End If

            End If
                
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Cambio de EXP x" & Exp & IIf(Exp <> 255, " en el mapa " & MapInfo(.Pos.Map).Name & "(" & .Pos.Map & ")", vbNullString))

        End If
        
    End With

    '<EhFooter>
    Exit Sub

HandleChangeMapInfoExp_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeMapInfoExp " & "at line " & Erl

    '</EhFooter>
End Sub

Private Sub HandleChangeMapInfoAttack(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeMapInfoAttack_Err

    '</EhHeader>

    Dim Activado As Byte
    
    Activado = Reader.ReadInt8
    
    With UserList(UserIndex)
    
        If EsGmDios(UserIndex) Then
            MapInfo(.Pos.Map).FreeAttack = IIf((Activado = 0), False, True)
            
            If Activado > 0 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Libre ataque en " & MapInfo(.Pos.Map).Name & " (" & .Pos.Map & ")", FontTypeNames.FONTTYPE_USERPREMIUM))
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ataque limitado entre facciones en " & MapInfo(.Pos.Map).Name & " (" & .Pos.Map & ")", FontTypeNames.FONTTYPE_USERPREMIUM))

            End If

        End If
        
    End With

    '<EhFooter>
    Exit Sub

HandleChangeMapInfoAttack_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeMapInfoAttack " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handle the "ChangeMapInfoRestricted" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeMapInfoRestricted_Err

    '</EhHeader>

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS", "FACCION".
    '***************************************************

    Dim tStr As String
    
    With UserList(UserIndex)
        
        tStr = Reader.ReadString8()
        
        If EsGmPriv(UserIndex) Then
            If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Then
                Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si es restringido el mapa.")
                
                MapInfo(UserList(UserIndex).Pos.Map).Restringir = RestrictStringToByte(tStr)
                
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Restringir", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Restringido: " & RestrictByteToString(MapInfo(.Pos.Map).Restringir), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleChangeMapInfoRestricted_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeMapInfoRestricted " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handle the "ChangeMapInfoNoMagic" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoMagic(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeMapInfoNoMagic_Err

    '</EhHeader>

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'MagiaSinEfecto -> Options: "1" , "0".
    '***************************************************
    
    Dim nomagic As Boolean
    
    With UserList(UserIndex)
        
        nomagic = Reader.ReadBool
        
        If EsGmDios(UserIndex) Then
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si está permitido usar la magia el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).MagiaSinEfecto = nomagic
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "MagiaSinEfecto", nomagic)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " MagiaSinEfecto: " & MapInfo(.Pos.Map).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleChangeMapInfoNoMagic_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeMapInfoNoMagic " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handle the "ChangeMapInfoNoInvi" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvi(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeMapInfoNoInvi_Err

    '</EhHeader>

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'InviSinEfecto -> Options: "1", "0"
    '***************************************************
    
    Dim noinvi As Boolean
    
    With UserList(UserIndex)
        
        noinvi = Reader.ReadBool()
        
        If EsGmDios(UserIndex) Then
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si está permitido usar la invisibilidad en el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).InviSinEfecto = noinvi
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "InviSinEfecto", noinvi)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " InviSinEfecto: " & MapInfo(.Pos.Map).InviSinEfecto, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleChangeMapInfoNoInvi_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeMapInfoNoInvi " & "at line " & Erl
        
    '</EhFooter>
End Sub
            
''
' Handle the "ChangeMapInfoNoResu" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoResu(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeMapInfoNoResu_Err

    '</EhHeader>

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'ResuSinEfecto -> Options: "1", "0"
    '***************************************************
    
    Dim noresu As Boolean
    
    With UserList(UserIndex)
        
        noresu = Reader.ReadBool()
        
        If EsGmDios(UserIndex) <> 0 Then
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si está permitido usar el resucitar en el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).ResuSinEfecto = noresu
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "ResuSinEfecto", noresu)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " ResuSinEfecto: " & MapInfo(.Pos.Map).ResuSinEfecto, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleChangeMapInfoNoResu_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeMapInfoNoResu " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handle the "ChangeMapInfoLand" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoLand(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeMapInfoLand_Err

    '</EhHeader>

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
    '***************************************************

    Dim tStr As String
    
    With UserList(UserIndex)
        
        tStr = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información del terreno del mapa.")
                
                MapInfo(UserList(UserIndex).Pos.Map).Terreno = TerrainStringToByte(tStr)
                
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Terreno", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Terreno: " & TerrainByteToString(MapInfo(.Pos.Map).Terreno), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el único útil es 'NIEVE' ya que al ingresarlo, la gente muere de frío en el mapa.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleChangeMapInfoLand_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeMapInfoLand " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handle the "ChangeMapInfoZone" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoZone(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeMapInfoZone_Err

    '</EhHeader>

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
    '***************************************************

    Dim tStr As String
    
    With UserList(UserIndex)
        
        tStr = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información de la zona del mapa.")
                MapInfo(UserList(UserIndex).Pos.Map).Zona = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Zona", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Zona: " & MapInfo(.Pos.Map).Zona, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el único útil es 'DUNGEON' ya que al ingresarlo, NO se sentirá el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleChangeMapInfoZone_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeMapInfoZone " & "at line " & Erl
        
    '</EhFooter>
End Sub
            
''
' Handle the "ChangeMapInfoStealNp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoStealNpc(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeMapInfoStealNpc_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 25/07/2010
    'RoboNpcsPermitido -> Options: "1", "0"
    '***************************************************
    
    Dim RoboNpc As Byte
    
    With UserList(UserIndex)
        
        RoboNpc = val(IIf(Reader.ReadBool(), 1, 0))
        
        If EsGmDios(UserIndex) Then
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si está permitido robar npcs en el mapa.")
            
            MapInfo(UserList(UserIndex).Pos.Map).RoboNpcsPermitido = RoboNpc
            
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "RoboNpcsPermitido", RoboNpc)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " RoboNpcsPermitido: " & MapInfo(.Pos.Map).RoboNpcsPermitido, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleChangeMapInfoStealNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeMapInfoStealNpc " & "at line " & Erl
        
    '</EhFooter>
End Sub
            
''
' Handle the "ChangeMapInfoNoOcultar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoOcultar(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeMapInfoNoOcultar_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 18/09/2010
    'OcultarSinEfecto -> Options: "1", "0"
    '***************************************************
    
    Dim NoOcultar As Byte

    Dim mapa      As Integer
    
    With UserList(UserIndex)
        
        NoOcultar = val(IIf(Reader.ReadBool(), 1, 0))
        
        If EsGmDios(UserIndex) Then
            
            mapa = .Pos.Map
            
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si está permitido ocultarse en el mapa " & mapa & ".")
            
            MapInfo(mapa).OcultarSinEfecto = NoOcultar
            
            Call WriteVar(App.Path & MapPath & "mapa" & mapa & ".dat", "Mapa" & mapa, "OcultarSinEfecto", NoOcultar)
            Call WriteConsoleMsg(UserIndex, "Mapa " & mapa & " OcultarSinEfecto: " & NoOcultar, FontTypeNames.FONTTYPE_INFO)

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleChangeMapInfoNoOcultar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeMapInfoNoOcultar " & "at line " & Erl
        
    '</EhFooter>
End Sub
           
''
' Handle the "ChangeMapInfoNoInvocar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvocar(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeMapInfoNoInvocar_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 18/09/2010
    'InvocarSinEfecto -> Options: "1", "0"
    '***************************************************
    
    Dim NoInvocar As Byte

    Dim mapa      As Integer
    
    With UserList(UserIndex)
        
        NoInvocar = val(IIf(Reader.ReadBool(), 1, 0))
        
        If EsGmDios(UserIndex) Then
            
            mapa = .Pos.Map
            
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha cambiado la información sobre si está permitido invocar en el mapa " & mapa & ".")
            
            MapInfo(mapa).InvocarSinEfecto = NoInvocar
            
            Call WriteVar(App.Path & MapPath & "mapa" & mapa & ".dat", "Mapa" & mapa, "InvocarSinEfecto", NoInvocar)
            Call WriteConsoleMsg(UserIndex, "Mapa " & mapa & " InvocarSinEfecto: " & NoInvocar, FontTypeNames.FONTTYPE_INFO)

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleChangeMapInfoNoInvocar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeMapInfoNoInvocar " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handle the "SaveMap" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleSaveMap_Err

    '</EhHeader>

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Saves the map
    '***************************************************
    With UserList(UserIndex)
        
        If Not EsGmDios(UserIndex) Then Exit Sub
        
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha guardado el mapa " & CStr(.Pos.Map))
        
        Call GrabarMapa(.Pos.Map, Maps_FilePath & "WORLDBACKUP\Mapa" & .Pos.Map)
        
        Call WriteConsoleMsg(UserIndex, "Mapa Guardado.", FontTypeNames.FONTTYPE_INFO)

    End With

    '<EhFooter>
    Exit Sub

HandleSaveMap_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleSaveMap " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handle the "DoBackUp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleDoBackUp_Err

    '</EhHeader>

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Show dobackup messages
    '***************************************************
    With UserList(UserIndex)
        
        If Not EsGmPriv(UserIndex) Then Exit Sub
        
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, .Name & " ha hecho un backup.")
        
        Call ES.DoBackUp 'Sino lo confunde con la id del paquete

    End With

    '<EhFooter>
    Exit Sub

HandleDoBackUp_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleDoBackUp " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handle the "HandleCreateNPC" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleCreateNPC_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/09/2010
    '26/09/2010: ZaMa - Ya no se pueden crear npcs pretorianos.
    '***************************************************
    With UserList(UserIndex)
        
        Dim NpcIndex As Integer
        
        NpcIndex = Reader.ReadInt()
        
        If Not EsGmDios(UserIndex) Then Exit Sub
                
        If GetVar(Npcs_FilePath, "NPC" & NpcIndex, "NAME") = vbNullString Then Exit Sub
               
        If val(GetVar(Npcs_FilePath, "NPC" & NpcIndex, "NPCTYPE")) = eNPCType.Pretoriano Or val(GetVar(Npcs_FilePath, "NPC" & NpcIndex, "NPCTYPE")) = eNPCType.eCommerceChar Then
            Call WriteConsoleMsg(UserIndex, "No puedes sumonear esta criatura. Revisa el numero de la misma. Gracias atentamente lautaro.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub

        End If
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)
        
        If NpcIndex <> 0 Then
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Sumoneó a " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleCreateNPC_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCreateNPC " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handle the "CreateNPCWithRespawn" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPCWithRespawn(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleCreateNPCWithRespawn_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/09/2010
    '26/09/2010: ZaMa - Ya no se pueden crear npcs pretorianos.
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim NpcIndex As Integer
        
        NpcIndex = Reader.ReadInt()
        
        If NpcIndex > NumNpcs Then Exit Sub
        
        If Not EsGmPriv(UserIndex) Then Exit Sub
         
        If val(GetVar(Npcs_FilePath, "NPC" & NpcIndex, "NPCTYPE")) = eNPCType.Pretoriano Then
            Call WriteConsoleMsg(UserIndex, "No puedes sumonear miembros que funcionan como guardines/pretorianos.", FontTypeNames.FONTTYPE_WARNING)

            Exit Sub

        End If
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True)
        
        If NpcIndex <> 0 Then
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Sumoneó con respawn " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleCreateNPCWithRespawn_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCreateNPCWithRespawn " & "at line " & Erl

    '</EhFooter>
End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleServerOpenToUsersToggle_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If Not EsGmPriv(UserIndex) Then Exit Sub
        
        If ServerSoloGMs > 0 Then
            Call WriteConsoleMsg(UserIndex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 0
            frmServidor.chkServerHabilitado.Value = vbUnchecked
        Else

            Dim A As Long
                
            For A = 1 To LastUser

                If Not EsGm(A) Then
                    Call Protocol.Kick(UserIndex, "Servidor restringido para administradores")

                End If

            Next A
                
            Call WriteConsoleMsg(UserIndex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 1
            frmServidor.chkServerHabilitado.Value = vbChecked

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleServerOpenToUsersToggle_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleServerOpenToUsersToggle " & "at line " & Erl

    '</EhFooter>
End Sub

''
' Handle the "TurnOffServer" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnOffServer(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleTurnOffServer_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    'Turns off the server
    '***************************************************
    Dim handle As Integer
    
    With UserList(UserIndex)
        
        If Not EsGmPriv(UserIndex) Then Exit Sub
        
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/APAGAR")
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡¡¡" & .Name & " VA A APAGAR EL SERVIDOR!!!", FontTypeNames.FONTTYPE_FIGHT))
        
        'Log
        handle = FreeFile
        Open LogPath & "Main.log" For Append Shared As #handle
        
        Print #handle, Date & " " & Time & " server apagado por " & .Name & ". "
        
        Close #handle
        
        Unload frmMain

    End With

    '<EhFooter>
    Exit Sub

HandleTurnOffServer_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleTurnOffServer " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handle the "TurnCriminal" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnCriminal(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleTurnCriminal_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/CONDEN " & UserName)
            
            tUser = NameIndex(UserName)

            If tUser > 0 Then Call VolverCriminal(tUser)

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleTurnCriminal_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleTurnCriminal " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handle the "ResetFactions" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleResetFactions(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleResetFactions_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 06/09/09
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim Char     As String

        Dim Temp     As Integer
        
        UserName = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/RAJAR " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If UserList(tUser).Faction.Status = 0 Then
                    Call WriteConsoleMsg(UserIndex, "El personaje " & UserName & " no pertenece a ninguna facción.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call mFacciones.Faction_RemoveUser(tUser)

                End If

            Else
                Char = CharPath & UserName & ".chr"
                
                If FileExist(Char, vbNormal) Then
                    Temp = val(GetVar(Char, "FACTION", "STATUS"))
                    
                    If Temp = 0 Then
                        Call WriteConsoleMsg(UserIndex, "El personaje " & UserName & " no pertenece a ninguna facción.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteVar(Char, "FACTION", "STATUS", "0")
                        Call WriteVar(Char, "FACTION", "ExFaction", CStr(Temp))
                        Call WriteVar(Char, "FACTION", "StartDate", vbNullString)
                        Call WriteVar(Char, "FACTION", "StartElv", "0")
                        Call WriteVar(Char, "FACTION", "StartFrags", "0")
                                
                        Dim A As Long
                                
                        For A = 1 To MAX_INVENTORY_SLOTS

                            If .Invent.Object(A).ObjIndex > 0 Then
                                If ObjData(.Invent.Object(A).ObjIndex).Real = 1 Or ObjData(.Invent.Object(A).ObjIndex).Caos = 1 Then
                                    Call QuitarObjetos(.Invent.Object(A).ObjIndex, .Invent.Object(A).Amount, UserIndex)

                                End If

                            End If

                        Next A

                    End If
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "El personaje " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleResetFactions_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleResetFactions " & "at line " & Erl

    '</EhFooter>
End Sub

''
' Handle the "SystemMessage" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleSystemMessage_Err

    '</EhHeader>

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/29/06
    'Send a message to all the users
    '***************************************************

    With UserList(UserIndex)
        
        Dim Message As String

        Message = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "Mensaje de sistema:" & Message)
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(Message))

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleSystemMessage_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleSystemMessage " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handle the "Ping" message
'
' @param userIndex The index of the user sending the message

Public Sub HandlePing(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandlePing_Err

    '</EhHeader>

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Show ping messages
    '***************************************************

    Call WritePong(UserIndex, Reader.ReadReal64())

    '<EhFooter>
    Exit Sub

HandlePing_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandlePing " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handle the "CreatePretorianClan" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreatePretorianClan(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleCreatePretorianClan_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/10/2010
    '***************************************************

    Dim Map   As Integer

    Dim X     As Byte

    Dim Y     As Byte

    Dim Index As Long
    
    With UserList(UserIndex)
        
        Map = Reader.ReadInt()
        X = Reader.ReadInt()
        Y = Reader.ReadInt()
        
        If Not EsGmPriv(UserIndex) Then Exit Sub
        
        ' Valid pos?
        If Not InMapBounds(Map, X, Y) Then
            Call WriteConsoleMsg(UserIndex, "Posición inválida.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
            
        ' Is already active any clan?
        If Not ClanPretoriano(7).Active Then
            
            If Not ClanPretoriano(Index).SpawnClan(Map, X, Y, Index) Then
                Call WriteConsoleMsg(UserIndex, "La posición no es apropiada para crear el clan", FontTypeNames.FONTTYPE_INFO)

            End If
        
        Else
            Call WriteConsoleMsg(UserIndex, "El clan pretoriano se encuentra activo en el mapa " & ClanPretoriano(Index).ClanMap & ". Utilice /EliminarPretorianos MAPA y reintente.", FontTypeNames.FONTTYPE_INFO)

        End If
    
    End With

    Exit Sub

    Call LogError("Error en HandleCreatePretorianClan. Error: " & Err.number & " - " & Err.description)
    '<EhFooter>
    Exit Sub

HandleCreatePretorianClan_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCreatePretorianClan " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handle the "CreatePretorianClan" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDeletePretorianClan(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleDeletePretorianClan_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/10/2010
    '***************************************************
    
    Dim Map   As Integer

    Dim Index As Long
    
    With UserList(UserIndex)
        
        Map = Reader.ReadInt()
        
        ' User Admin?
        If Not EsGmPriv(UserIndex) Then Exit Sub
        
        ' Valid map?
        If Map < 1 Or Map > NumMaps Then
            Call WriteConsoleMsg(UserIndex, "Mapa inválido.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        ' Search for the clan to be deleted
        If ClanPretoriano(7).ClanMap = Map Then
            ClanPretoriano(7).DeleteClan

        End If
    
    End With

    Exit Sub

    Call LogError("Error en HandleDeletePretorianClan. Error: " & Err.number & " - " & Err.description)
    '<EhFooter>
    Exit Sub

HandleDeletePretorianClan_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleDeletePretorianClan " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "Logged" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteLoggedMessage_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Logged" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.logged)
    
    With UserList(UserIndex)
    
        Call Writer.WriteInt8(.Clase)
        Call Writer.WriteInt8(.Raza)
        Call Writer.WriteInt8(.Genero)
        Call Writer.WriteInt8(.Account.CharsAmount)
        Call Writer.WriteInt32(.Account.Gld)

    End With
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteLoggedMessage_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteLoggedMessage " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteRemoveAllDialogs_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.RemoveDialogs)

    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteRemoveAllDialogs_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteRemoveAllDialogs " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal charindex As Integer)

    '<EhHeader>
    On Error GoTo WriteRemoveCharDialog_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
    '***************************************************

    Call SendData(ToOne, UserIndex, PrepareMessageRemoveCharDialog(charindex))
    '<EhFooter>
    Exit Sub

WriteRemoveCharDialog_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteRemoveCharDialog " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "NavigateToggle" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteNavigateToggle_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NavigateToggle" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.NavigateToggle)

    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteNavigateToggle_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteNavigateToggle " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal UserIndex As Integer, _
                           Optional ByVal Account As Boolean = False)

    '<EhHeader>
    On Error GoTo WriteDisconnect_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Disconnect" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.Disconnect)
    Call Writer.WriteBool(Account)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteDisconnect_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteDisconnect " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "UserOfferConfirm" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserOfferConfirm(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUserOfferConfirm_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    'Writes the "UserOfferConfirm" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.UserOfferConfirm)

    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteUserOfferConfirm_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUserOfferConfirm " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteCommerceEnd_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceEnd" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.CommerceEnd)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteCommerceEnd_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteCommerceEnd " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteBankEnd_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankEnd" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.BankEnd)

    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteBankEnd_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteBankEnd " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal UserIndex As Integer, _
                             ByVal NpcName As String, _
                             ByVal Quest As Byte, _
                             ByRef QuestList() As Byte)

    '<EhHeader>
    On Error GoTo WriteCommerceInit_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceInit" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.CommerceInit)
    Call Writer.WriteString8(NpcName)
    Call Writer.WriteInt8(Quest)
    
    If Quest > 0 Then
        Call Writer.WriteSafeArrayInt8(QuestList)

    End If
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteCommerceInit_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteCommerceInit " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal UserIndex As Integer, ByVal TypeBank As Byte)

    '<EhHeader>
    On Error GoTo WriteBankInit_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankInit" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.BankInit)
    Call Writer.WriteInt(UserList(UserIndex).Account.Gld)
    Call Writer.WriteInt(UserList(UserIndex).Account.Eldhir)
    Call Writer.WriteInt(TypeBank)
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteBankInit_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteBankInit " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUserCommerceInit_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.UserCommerceInit)
    Call Writer.WriteString8(UserList(UserIndex).ComUsu.DestNick)

    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteUserCommerceInit_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUserCommerceInit " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUserCommerceEnd_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.UserCommerceEnd)

    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteUserCommerceEnd_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUserCommerceEnd " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUpdateSta_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.UpdateSta)
    Call Writer.WriteInt(UserList(UserIndex).Stats.MinSta)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteUpdateSta_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateSta " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUpdateMana_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.UpdateMana)
    Call Writer.WriteInt(UserList(UserIndex).Stats.MinMan)
    Call Writer.WriteInt16(1)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteUpdateMana_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateMana " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUpdateHP_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.UpdateHP)
    Call Writer.WriteInt(UserList(UserIndex).Stats.MinHp)
    Call Writer.WriteInt16(1)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteUpdateHP_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateHP " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUpdateGold_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateGold" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.UpdateGold)
    Call Writer.WriteInt(UserList(UserIndex).Stats.Gld)
        
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteUpdateGold_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateGold " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "UpdateDsp" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateDsp(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUpdateDsp_Err

    '</EhHeader>

    '***************************************************
    'Author: WAICON
    'Last Modification: 06/05/2019
    'Writes the "UpdateDsp" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.UpdateDsp)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Eldhir)
    Call Writer.WriteInt32(UserList(UserIndex).Account.Eldhir)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteUpdateDsp_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateDsp " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "UpdateBankGold" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateBankGold(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUpdateBankGold_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    'Writes the "UpdateBankGold" message to the given user's outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ServerPacketID.UpdateBankGold)
    Call Writer.WriteInt(UserList(UserIndex).Account.Gld)
    Call Writer.WriteInt(UserList(UserIndex).Account.Eldhir)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteUpdateBankGold_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateBankGold " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUpdateExp_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateExp" message to the given user's outgoing data buffer
    '**************************************************

    Call Writer.WriteInt(ServerPacketID.UpdateExp)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Exp)
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteUpdateExp_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateExp " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenghtAndDexterity(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUpdateStrenghtAndDexterity_Err

    '</EhHeader>

    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.UpdateStrenghtAndDexterity)
    Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
    Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteUpdateStrenghtAndDexterity_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateStrenghtAndDexterity " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateDexterity(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUpdateDexterity_Err

    '</EhHeader>

    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.UpdateDexterity)
    Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteUpdateDexterity_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateDexterity " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenght(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUpdateStrenght_Err

    '</EhHeader>

    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.UpdateStrenght)
    Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteUpdateStrenght_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateStrenght " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal UserIndex As Integer, ByVal Map As Integer)

    '<EhHeader>
    On Error GoTo WriteChangeMap_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMap" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.ChangeMap)
    Call Writer.WriteInt(Map)
        
    If Map <> 0 Then
        Call Writer.WriteString8(MapInfo(Map).Name)
    Else
        Call Writer.WriteString8(vbNullString)

    End If

    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteChangeMap_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteChangeMap " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WritePosUpdate_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PosUpdate" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.PosUpdate)
    Call Writer.WriteInt(UserList(UserIndex).Pos.X)
    Call Writer.WriteInt(UserList(UserIndex).Pos.Y)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WritePosUpdate_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WritePosUpdate " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatOverHead(ByVal UserIndex As Integer, _
                             ByVal chat As String, _
                             ByVal charindex As Integer, _
                             ByVal Color As Long)

    '<EhHeader>
    On Error GoTo WriteChatOverHead_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatOverHead" message to the given user's outgoing data buffer
    '***************************************************

    Call SendData(ToOne, UserIndex, PrepareMessageChatOverHead(chat, charindex, Color))

    '<EhFooter>
    Exit Sub

WriteChatOverHead_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteChatOverHead " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteChatPersonalizado(ByVal UserIndex As Integer, _
                                  ByVal chat As String, _
                                  ByVal charindex As Integer, _
                                  ByVal Tipo As Byte)

    '<EhHeader>
    On Error GoTo WriteChatPersonalizado_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Dalmasso (CHOTS)
    'Last Modification: 11/06/2011
    'Writes the "ChatPersonalizado" message to the given user's outgoing data buffer
    '***************************************************

    Call SendData(ToOne, UserIndex, PrepareMessageChatPersonalizado(chat, charindex, Tipo))

    '<EhFooter>
    Exit Sub

WriteChatPersonalizado_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteChatPersonalizado " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, _
                           ByVal chat As String, _
                           ByVal FontIndex As FontTypeNames, _
                           Optional ByVal MessageType As eMessageType = Info)

    '<EhHeader>
    On Error GoTo WriteConsoleMsg_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
    '***************************************************

    Call SendData(ToOne, UserIndex, PrepareMessageConsoleMsg(chat, FontIndex, MessageType))
    
    '<EhFooter>
    Exit Sub

WriteConsoleMsg_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteConsoleMsg " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteCommerceChat(ByVal UserIndex As Integer, _
                             ByVal chat As String, _
                             ByVal FontIndex As FontTypeNames)

    '<EhHeader>
    On Error GoTo WriteCommerceChat_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 05/17/06
    'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
    '***************************************************
    
    Call SendData(ToOne, UserIndex, PrepareCommerceConsoleMsg(chat, FontIndex))
    
    '<EhFooter>
    Exit Sub

WriteCommerceChat_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteCommerceChat " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal Message As String)

    '<EhHeader>
    On Error GoTo WriteShowMessageBox_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.ShowMessageBox)
    Call Writer.WriteString8(Message)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteShowMessageBox_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteShowMessageBox " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUserIndexInServer_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.UserIndexInServer)
    Call Writer.WriteInt(UserIndex)
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteUserIndexInServer_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUserIndexInServer " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUserCharIndexInServer_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.UserCharIndexInServer)
    Call Writer.WriteInt(UserList(UserIndex).Char.charindex)
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteUserCharIndexInServer_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUserCharIndexInServer " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "CharacterCreate" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, _
                                ByVal Body As Integer, _
                                ByVal BodyAttack As Integer, _
                                ByVal Head As Integer, _
                                ByVal Heading As eHeading, _
                                ByVal charindex As Integer, _
                                ByVal X As Byte, _
                                ByVal Y As Byte, _
                                ByVal Weapon As Integer, _
                                ByVal Shield As Integer, _
                                ByVal FX As Integer, _
                                ByVal FXLoops As Integer, _
                                ByVal helmet As Integer, _
                                ByVal Name As String, _
                                ByVal NickColor As Byte, _
                                ByVal Privileges As Byte, _
                                ByRef AuraIndex() As Byte, _
                                ByVal speeding As Single, _
                                ByVal Idle As Boolean, _
                                Optional ByVal NpcIndex As Integer = 0)

    '<EhHeader>
    On Error GoTo WriteCharacterCreate_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterCreate" message to the given user's outgoing data buffer
    '***************************************************
    
    Call SendData(ToOne, UserIndex, PrepareMessageCharacterCreate(Body, BodyAttack, Head, Heading, charindex, X, Y, Weapon, Shield, FX, FXLoops, helmet, Name, NickColor, Privileges, AuraIndex, NpcIndex, Idle, False, speeding))

    '<EhFooter>
    Exit Sub

WriteCharacterCreate_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteCharacterCreate " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal UserIndex As Integer, ByVal charindex As Integer)

    '<EhHeader>
    On Error GoTo WriteCharacterRemove_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterRemove" message to the given user's outgoing data buffer
    '***************************************************
    
    Call SendData(ToOne, UserIndex, PrepareMessageCharacterRemove(charindex))

    '<EhFooter>
    Exit Sub

WriteCharacterRemove_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteCharacterRemove " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal UserIndex As Integer, _
                              ByVal charindex As Integer, _
                              ByVal X As Byte, _
                              ByVal Y As Byte)

    '<EhHeader>
    On Error GoTo WriteCharacterMove_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterMove" message to the given user's outgoing data buffer
    '***************************************************

    Call SendData(ToOne, UserIndex, PrepareMessageCharacterMove(charindex, X, Y))

    '<EhFooter>
    Exit Sub

WriteCharacterMove_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteCharacterMove " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteForceCharMove(ByVal UserIndex, ByVal Direccion As eHeading)

    '<EhHeader>
    On Error GoTo WriteForceCharMove_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/03/2009
    'Writes the "ForceCharMove" message to the given user's outgoing data buffer
    '***************************************************

    Call SendData(ToOne, UserIndex, PrepareMessageForceCharMove(Direccion))

    '<EhFooter>
    Exit Sub

WriteForceCharMove_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteForceCharMove " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "CharacterChange" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterChange(ByVal UserIndex As Integer, _
                                ByVal Body As Integer, _
                                ByVal Head As Integer, _
                                ByVal Heading As eHeading, _
                                ByVal charindex As Integer, _
                                ByVal Weapon As Integer, _
                                ByVal Shield As Integer, _
                                ByVal FX As Integer, _
                                ByVal FXLoops As Integer, _
                                ByVal helmet As Integer, _
                                ByRef AuraIndex() As Byte, _
                                ByVal Idle As Boolean, _
                                ByVal Navegacion As Boolean)

    '<EhHeader>
    On Error GoTo WriteCharacterChange_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterChange" message to the given user's outgoing data buffer
    '***************************************************

    Call SendData(ToOne, UserIndex, PrepareMessageCharacterChange(Body, 0, Head, Heading, charindex, Weapon, Shield, FX, FXLoops, helmet, AuraIndex, UserList(UserIndex).flags.ModoStream, Idle, Navegacion))

    '<EhFooter>
    Exit Sub

WriteCharacterChange_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteCharacterChange " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectCreate(ByVal UserIndex As Integer, _
                             ByVal ObjIndex As Integer, _
                             ByVal GrhIndex As Long, _
                             ByVal X As Byte, _
                             ByVal Y As Byte, _
                             ByVal Sound As Integer)

    '<EhHeader>
    On Error GoTo WriteObjectCreate_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ObjectCreate" message to the given user's outgoing data buffer
    '***************************************************

    Call SendData(ToOne, UserIndex, PrepareMessageObjectCreate(ObjIndex, GrhIndex, X, Y, vbNullString, 0, Sound))

    '<EhFooter>
    Exit Sub

WriteObjectCreate_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteObjectCreate " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

    '<EhHeader>
    On Error GoTo WriteObjectDelete_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ObjectDelete" message to the given user's outgoing data buffer
    '***************************************************

    Call SendData(ToOne, UserIndex, PrepareMessageObjectDelete(X, Y))

    '<EhFooter>
    Exit Sub

WriteObjectDelete_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteObjectDelete " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal UserIndex As Integer, _
                              ByVal X As Byte, _
                              ByVal Y As Byte, _
                              ByVal Blocked As Boolean)

    '<EhHeader>
    On Error GoTo WriteBlockPosition_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlockPosition" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.BlockPosition)
    Call Writer.WriteInt(X)
    Call Writer.WriteInt(Y)
    Call Writer.WriteBool(Blocked)

    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteBlockPosition_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteBlockPosition " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "PlayMusic" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMusic(ByVal UserIndex As Integer, ByVal Music As Integer)

    '<EhHeader>
    On Error GoTo WritePlayMusic_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PlayMusic" message to the given user's outgoing data buffer
    '***************************************************

    Call SendData(ToOne, UserIndex, PrepareMessagePlayMusic(Music))

    '<EhFooter>
    Exit Sub

WritePlayMusic_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WritePlayMusic " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "PlayEffect" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayEffect(ByVal UserIndex As Integer, _
                           ByVal Wave As Integer, _
                           ByVal X As Byte, _
                           ByVal Y As Byte)

    '<EhHeader>
    On Error GoTo WritePlayEffect_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/08/07
    'Last Modified by: Rapsodius
    'Added X and Y positions for 3D Sounds
    '***************************************************

    Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(Wave, X, Y))

    '<EhFooter>
    Exit Sub

WritePlayEffect_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WritePlayEffect " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WritePauseToggle_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PauseToggle" message to the given user's outgoing data buffer
    '***************************************************

    Call SendData(ToOne, UserIndex, PrepareMessagePauseToggle())

    '<EhFooter>
    Exit Sub

WritePauseToggle_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WritePauseToggle " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "CreateFX" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal UserIndex As Integer, _
                         ByVal charindex As Integer, _
                         ByVal FX As Integer, _
                         ByVal FXLoops As Integer)

    '<EhHeader>
    On Error GoTo WriteCreateFX_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateFX" message to the given user's outgoing data buffer
    '***************************************************

    Call SendData(ToOne, UserIndex, PrepareMessageCreateFX(charindex, FX, FXLoops))

    '<EhFooter>
    Exit Sub

WriteCreateFX_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteCreateFX " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUpdateUserStats_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.UpdateUserStats)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxHp)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinHp)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxMan)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinMan)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxSta)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinSta)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Gld)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Eldhir)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.Elv)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Elu)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Exp)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Points)
    
    Dim estatus As Byte
    
    If UserList(UserIndex).flags.Bronce = 1 Then
        estatus = 1

    End If
    
    If UserList(UserIndex).flags.Plata = 1 Then
        estatus = 2

    End If
    
    If UserList(UserIndex).flags.Oro = 1 Then
        estatus = 3

    End If
    
    Call Writer.WriteInt8(estatus)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteUpdateUserStats_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateUserStats " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ChangeInventorySlot" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

    '<EhHeader>
    On Error GoTo WriteChangeInventorySlot_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 25/05/2011 (Amraphen)
    'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
    '3/12/09: Budi - Ahora se envia MaxDef y MinDef en lugar de Def
    '25/05/2011: Amraphen - Ahora se envía la defensa según se tiene equipado armadura de segunda jerarquía o no.
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.ChangeInventorySlot)
    Call Writer.WriteInt(Slot)
        
    Dim ObjIndex As Integer

    Dim obData   As ObjData
        
    ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    Call Writer.WriteInt(ObjIndex)
        
    If ObjIndex > 0 Then
        obData = ObjData(ObjIndex)
        
        'Si tiene armadura de segunda jerarquía obtiene un porcentaje de defensa adicional.
        If obData.Caos = 1 Or obData.Real = 1 Then
            If UserList(UserIndex).Faction.Status > 0 Then
                obData.MinDef = obData.MinDef + InfoFaction(UserList(UserIndex).Faction.Status).Range(UserList(UserIndex).Faction.Range).MinDef
                obData.MaxDef = obData.MaxDef + InfoFaction(UserList(UserIndex).Faction.Status).Range(UserList(UserIndex).Faction.Range).MaxDef

            End If

        End If
                
    End If
        
    Call Writer.WriteString8(obData.Name)
    Call Writer.WriteInt(UserList(UserIndex).Invent.Object(Slot).Amount)
    Call Writer.WriteBool(UserList(UserIndex).Invent.Object(Slot).Equipped)
    Call Writer.WriteInt(obData.GrhIndex)
    Call Writer.WriteInt(obData.OBJType)
    Call Writer.WriteInt(obData.MaxHit)
    Call Writer.WriteInt(obData.MinHit)
    Call Writer.WriteInt(obData.MaxDef)
    Call Writer.WriteInt(obData.MinDef)
    Call Writer.WriteReal32(SalePrice(ObjIndex))
    Call Writer.WriteReal32(SalePriceDiamanteAzul(ObjIndex))
    Call Writer.WriteBool(CanUse_Inventory(UserIndex, ObjIndex))
    
    Call Writer.WriteInt(obData.MinHitMag)
    Call Writer.WriteInt(obData.MaxHitMag)
    Call Writer.WriteInt(obData.DefensaMagicaMin)
    Call Writer.WriteInt(obData.DefensaMagicaMax)
    
    Call Writer.WriteInt8(obData.Bronce)
    Call Writer.WriteInt8(obData.Plata)
    Call Writer.WriteInt8(obData.Oro)
    Call Writer.WriteInt8(obData.Premium)
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteChangeInventorySlot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteChangeInventorySlot " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteAddSlots(ByVal UserIndex As Integer, ByVal Mochila As eMochilas)

    '<EhHeader>
    On Error GoTo WriteAddSlots_Err

    '</EhHeader>

    '***************************************************
    'Author: Budi
    'Last Modification: 01/12/09
    'Writes the "AddSlots" message to the given user's outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ServerPacketID.AddSlots)
    Call Writer.WriteInt(Mochila)
    
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteAddSlots_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteAddSlots " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

    '<EhHeader>
    On Error GoTo WriteChangeBankSlot_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/03/09
    'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
    '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.ChangeBankSlot)
    Call Writer.WriteInt(Slot)
        
    Dim ObjIndex As Integer

    Dim obData   As ObjData
        
    ObjIndex = UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex
        
    Call Writer.WriteInt(ObjIndex)
        
    If ObjIndex > 0 Then
        obData = ObjData(ObjIndex)

    End If
        
    Call Writer.WriteString8(obData.Name)
    Call Writer.WriteInt(UserList(UserIndex).BancoInvent.Object(Slot).Amount)
    Call Writer.WriteInt(obData.GrhIndex)
    Call Writer.WriteInt(obData.OBJType)
    Call Writer.WriteInt(obData.MaxHit)
    Call Writer.WriteInt(obData.MinHit)
    Call Writer.WriteInt(obData.MaxDef)
    Call Writer.WriteInt(obData.MinDef)
    Call Writer.WriteInt(obData.Valor)
    Call Writer.WriteInt(obData.ValorEldhir)
    Call Writer.WriteBool(CanUse_Inventory(UserIndex, ObjIndex))
        
    Call Writer.WriteInt(obData.MinHitMag)
    Call Writer.WriteInt(obData.MaxHitMag)
    Call Writer.WriteInt(obData.DefensaMagicaMin)
    Call Writer.WriteInt(obData.DefensaMagicaMax)

    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteChangeBankSlot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteChangeBankSlot " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ChangeBankSlot_Account" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot_Account(ByVal UserIndex As Integer, ByVal Slot As Byte)

    '<EhHeader>
    On Error GoTo WriteChangeBankSlot_Account_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/03/09
    'Writes the "ChangeBankSlot_Account" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.ChangeBankSlot_Account)
    Call Writer.WriteInt(Slot)
        
    Dim ObjIndex As Integer

    Dim obData   As ObjData
        
    ObjIndex = UserList(UserIndex).Account.BancoInvent.Object(Slot).ObjIndex
        
    Call Writer.WriteInt(ObjIndex)
        
    If ObjIndex > 0 Then
        obData = ObjData(ObjIndex)

    End If
        
    Call Writer.WriteString8(obData.Name)
    Call Writer.WriteInt(UserList(UserIndex).Account.BancoInvent.Object(Slot).Amount)
    Call Writer.WriteInt(obData.GrhIndex)
    Call Writer.WriteInt(obData.OBJType)
    Call Writer.WriteInt(obData.MaxHit)
    Call Writer.WriteInt(obData.MinHit)
    Call Writer.WriteInt(obData.MaxDef)
    Call Writer.WriteInt(obData.MinDef)
    Call Writer.WriteInt(obData.Valor)
    Call Writer.WriteInt(obData.ValorEldhir)
    Call Writer.WriteBool(CanUse_Inventory(UserIndex, ObjIndex))
        
    Call Writer.WriteInt(obData.MinHitMag)
    Call Writer.WriteInt(obData.MaxHitMag)
    Call Writer.WriteInt(obData.DefensaMagicaMin)
    Call Writer.WriteInt(obData.DefensaMagicaMax)

    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteChangeBankSlot_Account_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteChangeBankSlot_Account " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal Slot As Integer)

    '<EhHeader>
    On Error GoTo WriteChangeSpellSlot_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.ChangeSpellSlot)
    Call Writer.WriteInt(Slot)
    Call Writer.WriteInt(UserList(UserIndex).Stats.UserHechizos(Slot))
        
    If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
        Call Writer.WriteString8(Hechizos(UserList(UserIndex).Stats.UserHechizos(Slot)).Nombre)
    Else
        Call Writer.WriteString8("(Vacio)")

    End If

    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteChangeSpellSlot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteChangeSpellSlot " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteAttributes_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Atributes" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.Atributes)
    Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
    Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
    Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
    Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
    Call Writer.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteAttributes_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteAttributes " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "RestOK" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteRestOK_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RestOK" message to the given user's outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ServerPacketID.RestOK)

    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteRestOK_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteRestOK " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal Message As String)

    '<EhHeader>
    On Error GoTo WriteErrorMsg_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ErrorMsg" message to the given user's outgoing data buffer
    '***************************************************

    Call SendData(ToOne, UserIndex, PrepareMessageErrorMsg(Message), , True)

    '<EhFooter>
    Exit Sub

WriteErrorMsg_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteErrorMsg " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "Blind" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteBlind_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Blind" message to the given user's outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ServerPacketID.Blind)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteBlind_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteBlind " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "Dumb" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteDumb_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Dumb" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.Dumb)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteDumb_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteDumb " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data Reader.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal UserIndex As Integer, _
                                       ByVal Slot As Byte, _
                                       ByRef Obj As Obj, _
                                       ByVal Price As Single, _
                                       ByVal Price2 As Single)

    '<EhHeader>
    On Error GoTo WriteChangeNPCInventorySlot_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/03/09
    'Last Modified by: Budi
    'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
    '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
    '***************************************************

    Dim ObjInfo As ObjData
    
    If Obj.ObjIndex >= LBound(ObjData()) And Obj.ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(Obj.ObjIndex)
        
    End If
    
    Call Writer.WriteInt(ServerPacketID.ChangeNPCInventorySlot)
    Call Writer.WriteInt(Slot)
    Call Writer.WriteInt(Obj.ObjIndex)
    Call Writer.WriteString8(ObjInfo.Name)
    Call Writer.WriteInt(Obj.Amount)
    Call Writer.WriteReal32(Price)
    Call Writer.WriteInt(ObjInfo.GrhIndex)
            
    Call Writer.WriteInt(ObjInfo.OBJType)
    Call Writer.WriteInt(ObjInfo.MaxHit)
    Call Writer.WriteInt(ObjInfo.MinHit)
    Call Writer.WriteInt(ObjInfo.MaxDef)
    Call Writer.WriteInt(ObjInfo.MinDef)
    Call Writer.WriteReal32(Price2)
    Call Writer.WriteBool(CanUse_Inventory(UserIndex, Obj.ObjIndex))
            
    Call Writer.WriteInt(ObjInfo.MinHitMag)
    Call Writer.WriteInt(ObjInfo.MaxHitMag)
    Call Writer.WriteInt(ObjInfo.DefensaMagicaMin)
    Call Writer.WriteInt(ObjInfo.DefensaMagicaMax)
        
    Call Writer.WriteInt(NpcInventory_GetAnimation(UserIndex, Obj.ObjIndex))
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteChangeNPCInventorySlot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteChangeNPCInventorySlot " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUpdateHungerAndThirst_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.UpdateHungerAndThirst)
    Call Writer.WriteInt(UserList(UserIndex).Stats.MaxAGU)
    Call Writer.WriteInt(UserList(UserIndex).Stats.MinAGU)
    Call Writer.WriteInt(UserList(UserIndex).Stats.MaxHam)
    Call Writer.WriteInt(UserList(UserIndex).Stats.MinHam)
        
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteUpdateHungerAndThirst_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateHungerAndThirst " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "MiniStats" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal UserIndex As Integer, ByVal tUser As Integer)

    '<EhHeader>
    On Error GoTo WriteMiniStats_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MiniStats" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.MiniStats)
        
    Call Writer.WriteInt16(UserList(tUser).Faction.FragsCiu)
    Call Writer.WriteInt16(UserList(tUser).Faction.FragsCri)
        
    Call Writer.WriteInt8(UserList(tUser).Clase)
    Call Writer.WriteInt8(UserList(tUser).Raza)
    Call Writer.WriteInt32(UserList(tUser).Reputacion.promedio)
    
    Call Writer.WriteInt8(UserList(tUser).Stats.Elv)
    Call Writer.WriteInt32(UserList(tUser).Stats.Exp)
    Call Writer.WriteInt32(UserList(tUser).Stats.Elu)
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteMiniStats_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteMiniStats " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data Reader.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal skillPoints As Integer)

    '<EhHeader>
    On Error GoTo WriteLevelUp_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LevelUp" message to the given user's outgoing data buffer
    '***************************************************
    
    Call Writer.WriteInt(ServerPacketID.LevelUp)
    Call Writer.WriteInt(skillPoints)
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteLevelUp_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteLevelUp " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal UserIndex As Integer, _
                             ByVal charindex As Integer, _
                             ByVal Invisible As Boolean)

    '<EhHeader>
    On Error GoTo WriteSetInvisible_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetInvisible" message to the given user's outgoing data buffer
    '***************************************************
    
    Call SendData(ToOne, UserIndex, PrepareMessageSetInvisible(charindex, Invisible))

    '<EhFooter>
    Exit Sub

WriteSetInvisible_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteSetInvisible " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteBlindNoMore_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlindNoMore" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.BlindNoMore)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteBlindNoMore_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteBlindNoMore " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteDumbNoMore_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DumbNoMore" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.DumbNoMore)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteDumbNoMore_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteDumbNoMore " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteSendSkills_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 11/19/09
    'Writes the "SendSkills" message to the given user's outgoing data buffer
    '11/19/09: Pato - Now send the percentage of progress of the skills.
    '***************************************************

    Dim i As Long
    
    With UserList(UserIndex)
        Call Writer.WriteInt(ServerPacketID.SendSkills)
        
        Call Writer.WriteInt8(UserList(UserIndex).Clase)
        
        For i = 1 To NUMSKILLS
            Call Writer.WriteInt8(UserList(UserIndex).Stats.UserSkills(i))
        Next i

    End With

    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteSendSkills_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteSendSkills " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteParalizeOK_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/12/07
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    'Writes the "ParalizeOK" message to the given user's outgoing data buffer
    'And updates user position
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.ParalizeOK)
    Call WritePosUpdate(UserIndex)
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteParalizeOK_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteParalizeOK " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)

    '<EhHeader>
    On Error GoTo WriteShowUserRequest_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.ShowUserRequest)
    Call Writer.WriteString8(details)
    
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteShowUserRequest_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteShowUserRequest " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "TradeOK" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTradeOK(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteTradeOK_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TradeOK" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.TradeOK)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteTradeOK_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteTradeOK " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "BankOK" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankOK(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteBankOK_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankOK" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.BankOK)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteBankOK_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteBankOK " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, _
                                    ByVal OfferSlot As Byte, _
                                    ByVal ObjIndex As Integer, _
                                    ByVal Amount As Long)

    '<EhHeader>
    On Error GoTo WriteChangeUserTradeSlot_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/03/09
    'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
    '25/11/2009: ZaMa - Now sends the specific offer slot to be modified.
    '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.ChangeUserTradeSlot)
        
    Call Writer.WriteInt(OfferSlot)
    Call Writer.WriteInt(ObjIndex)
    Call Writer.WriteInt(Amount)
        
    If ObjIndex > 0 Then
        Call Writer.WriteInt(ObjData(ObjIndex).GrhIndex)
        Call Writer.WriteInt(ObjData(ObjIndex).OBJType)
        Call Writer.WriteInt(ObjData(ObjIndex).MaxHit)
        Call Writer.WriteInt(ObjData(ObjIndex).MinHit)
        Call Writer.WriteInt(ObjData(ObjIndex).MaxDef)
        Call Writer.WriteInt(ObjData(ObjIndex).MinDef)
        Call Writer.WriteInt(SalePrice(ObjIndex))
        Call Writer.WriteString8(ObjData(ObjIndex).Name)
        Call Writer.WriteInt(SalePriceDiamanteAzul(ObjIndex))
            
        Call Writer.WriteBool(CanUse_Inventory(UserIndex, ObjIndex))
        Call Writer.WriteInt(ObjData(ObjIndex).Bronce)
        Call Writer.WriteInt(ObjData(ObjIndex).Plata)
        Call Writer.WriteInt(ObjData(ObjIndex).Oro)
        Call Writer.WriteInt(ObjData(ObjIndex).Premium)

    End If

    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteChangeUserTradeSlot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteChangeUserTradeSlot " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByRef npcNames() As String)

    '<EhHeader>
    On Error GoTo WriteSpawnList_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpawnList" message to the given user's outgoing data buffer
    '***************************************************

    Dim i   As Long

    Dim Tmp As String
    
    Call Writer.WriteInt(ServerPacketID.SpawnList)
        
    For i = LBound(npcNames()) To UBound(npcNames())
        Tmp = Tmp & npcNames(i) & SEPARATOR
    Next i
        
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
    Call Writer.WriteString8(Tmp)

    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteSpawnList_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteSpawnList " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ShowDenounces" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowDenounces(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteShowDenounces_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    'Writes the "ShowDenounces" message to the given user's outgoing data buffer
    '***************************************************
    
    Dim DenounceIndex As Long

    Dim DenounceList  As String

    Call Writer.WriteInt(ServerPacketID.ShowDenounces)
        
    For DenounceIndex = 1 To Denuncias.Longitud
        DenounceList = DenounceList & Denuncias.VerElemento(DenounceIndex, False) & SEPARATOR
    Next DenounceIndex
        
    If LenB(DenounceList) <> 0 Then DenounceList = Left$(DenounceList, Len(DenounceList) - 1)
        
    Call Writer.WriteString8(DenounceList)
        
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteShowDenounces_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteShowDenounces " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteShowGMPanelForm_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.ShowGMPanelForm)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteShowGMPanelForm_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteShowGMPanelForm " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal UserIndex As Integer, _
                             ByRef userNamesList() As String, _
                             ByVal cant As Integer)

    '<EhHeader>
    On Error GoTo WriteUserNameList_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06 NIGO:
    'Writes the "UserNameList" message to the given user's outgoing data buffer
    '***************************************************

    Dim i   As Long

    Dim Tmp As String
    
    Call Writer.WriteInt(ServerPacketID.UserNameList)
        
    ' Prepare user's names list
    For i = 1 To cant
        Tmp = Tmp & userNamesList(i) & SEPARATOR
    Next i
        
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
    Call Writer.WriteString8(Tmp)
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteUserNameList_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUserNameList " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "Pong" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal UserIndex As Integer, ByVal tPing As Double)

    '<EhHeader>
    On Error GoTo WritePong_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Pong" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.Pong)
    Call Writer.WriteReal64(tPing)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WritePong_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WritePong " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo FlushBuffer_Err

    '</EhHeader>
    
    Server.Flush UserIndex

    '<EhFooter>
    Exit Sub

FlushBuffer_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.FlushBuffer " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal charindex As Integer, _
                                           ByVal Invisible As Boolean, _
                                           Optional ByVal Intermitencia As Boolean = False) As String

    '<EhHeader>
    On Error GoTo PrepareMessageSetInvisible_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "SetInvisible" message and returns it.
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.SetInvisible)
        
    Call Writer.WriteInt(charindex)
    Call Writer.WriteBool(Invisible)
    Call Writer.WriteBool(Intermitencia)
        
    '<EhFooter>
    Exit Function

PrepareMessageSetInvisible_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageSetInvisible " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function PrepareMessageCharacterChangeNick(ByVal charindex As Integer, _
                                                  ByVal NewNick As String) As String

    '<EhHeader>
    On Error GoTo PrepareMessageCharacterChangeNick_Err

    '</EhHeader>

    '***************************************************
    'Author: Budi
    'Last Modification: 07/23/09
    'Prepares the "Change Nick" message and returns it.
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.CharacterChangeNick)
        
    Call Writer.WriteInt(charindex)
    Call Writer.WriteString8(NewNick)

    '<EhFooter>
    Exit Function

PrepareMessageCharacterChangeNick_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageCharacterChangeNick " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageChatOverHead(ByVal chat As String, _
                                           ByVal charindex As Integer, _
                                           ByVal Color As Long) As String

    '<EhHeader>
    On Error GoTo PrepareMessageChatOverHead_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ChatOverHead" message and returns it.
    '***************************************************
    
    Call Writer.WriteInt(ServerPacketID.ChatOverHead)
    Call Writer.WriteString16(chat)
    Call Writer.WriteInt(charindex)
        
    ' Write rgb channels and save one byte from long :D
    Call Writer.WriteInt(Color And &HFF)
    Call Writer.WriteInt((Color And &HFF00&) \ &H100&)
    Call Writer.WriteInt((Color And &HFF0000) \ &H10000)

    '<EhFooter>
    Exit Function

PrepareMessageChatOverHead_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageChatOverHead " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function PrepareMessageChatPersonalizado(ByVal chat As String, _
                                                ByVal charindex As Integer, _
                                                ByVal Tipo As Byte) As String

    '<EhHeader>
    On Error GoTo PrepareMessageChatPersonalizado_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Dalmasso (CHOTS)
    'Last Modification: 11/06/2011
    'Prepares the "ChatPersonalizado" message and returns it.
    '**************************************************
        
    Call Writer.WriteInt(ServerPacketID.ChatPersonalizado)
    Call Writer.WriteString16(chat)
    Call Writer.WriteInt(charindex)
        
    ' Write the type of message
    '1=normal
    '2=clan
    '3=party
    '4=gritar
    '5=palabras magicas
    '6=susurrar
    Call Writer.WriteInt(Tipo)

    '<EhFooter>
    Exit Function

PrepareMessageChatPersonalizado_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageChatPersonalizado " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @param    MessageType type of console message (General, Guild, Party)
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal chat As String, _
                                         ByVal FontIndex As FontTypeNames, _
                                         Optional ByVal MessageType As eMessageType = Info) As String

    '<EhHeader>
    On Error GoTo PrepareMessageConsoleMsg_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/05/11 (D'Artagnan)
    'Prepares the "MessageType" message and returns it.
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.ConsoleMsg)
    Call Writer.WriteString8(chat)
    Call Writer.WriteInt(FontIndex)
    Call Writer.WriteInt(MessageType)
        
    '<EhFooter>
    Exit Function

PrepareMessageConsoleMsg_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageConsoleMsg " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function PrepareCommerceConsoleMsg(ByRef chat As String, _
                                          ByVal FontIndex As FontTypeNames) As String

    '<EhHeader>
    On Error GoTo PrepareCommerceConsoleMsg_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 03/12/2009
    'Prepares the "CommerceConsoleMsg" message and returns it.
    '***************************************************
    
    Call Writer.WriteInt(ServerPacketID.CommerceChat)
    Call Writer.WriteString8(chat)
    Call Writer.WriteInt(FontIndex)

    '<EhFooter>
    Exit Function

PrepareCommerceConsoleMsg_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareCommerceConsoleMsg " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "CreateFX" message and returns it.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCreateFX(ByVal charindex As Integer, _
                                       ByVal FX As Integer, _
                                       ByVal FXLoops As Integer, _
                                       Optional ByVal IsMeditation As Boolean = False) As String

    '<EhHeader>
    On Error GoTo PrepareMessageCreateFX_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CreateFX" message and returns it
    '***************************************************
    
    Call Writer.WriteInt(ServerPacketID.CreateFX)
    Call Writer.WriteInt(charindex)
    Call Writer.WriteInt(FX)
    Call Writer.WriteInt(FXLoops)
    Call Writer.WriteBool(IsMeditation)

    '<EhFooter>
    Exit Function

PrepareMessageCreateFX_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageCreateFX " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "PlayEffect" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayEffect(ByVal Wave As Integer, _
                                         ByVal X As Byte, _
                                         ByVal Y As Byte, _
                                         Optional ByVal Entity As Long = 0, _
                                         Optional ByVal Repeat As Boolean = False, _
                                         Optional ByVal MapOnly As Boolean = False) As String

    '<EhHeader>
    On Error GoTo PrepareMessagePlayEffect_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/08/07
    'Last Modified by: Rapsodius
    'Added X and Y positions for 3D Sounds
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.PlayWave)
    Call Writer.WriteInt(Wave)
    Call Writer.WriteInt(X)
    Call Writer.WriteInt(Y)
    Call Writer.WriteInt(Entity)
    Call Writer.WriteBool(Repeat)
    Call Writer.WriteBool(MapOnly)

    '<EhFooter>
    Exit Function

PrepareMessagePlayEffect_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessagePlayEffect " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function PrepareMessageStopWaveMap(ByVal X As Byte, _
                                          ByVal Y As Byte, _
                                          ByVal Inmediatily As Boolean) As String

    On Error GoTo PrepareMessagePlayEffect_Err
        
    Call Writer.WriteInt(ServerPacketID.StopWaveMap)
    Call Writer.WriteInt(X)
    Call Writer.WriteInt(Y)
    Call Writer.WriteBool(Inmediatily)

    '<EhFooter>
    Exit Function

PrepareMessagePlayEffect_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessagePlayEffect " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageShowMessageBox(ByVal chat As String) As String

    '<EhHeader>
    On Error GoTo PrepareMessageShowMessageBox_Err

    '</EhHeader>

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/08/07
    'Prepares the "ShowMessageBox" message and returns it
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.ShowMessageBox)
    Call Writer.WriteString8(chat)

    '<EhFooter>
    Exit Function

PrepareMessageShowMessageBox_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageShowMessageBox " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "PlayMusic" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMusic(ByVal Music As Integer) As String

    '<EhHeader>
    On Error GoTo PrepareMessagePlayMusic_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "PlayMusic" message and returns it
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.PlayMusic)
    Call Writer.WriteInt(Music)

    '<EhFooter>
    Exit Function

PrepareMessagePlayMusic_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessagePlayMusic " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As String

    '<EhHeader>
    On Error GoTo PrepareMessagePauseToggle_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "PauseToggle" message and returns it
    '***************************************************
    
    Call Writer.WriteInt(ServerPacketID.PauseToggle)

    '<EhFooter>
    Exit Function

PrepareMessagePauseToggle_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessagePauseToggle " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte) As String

    '<EhHeader>
    On Error GoTo PrepareMessageObjectDelete_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ObjectDelete" message and returns it
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.ObjectDelete)
    Call Writer.WriteInt(X)
    Call Writer.WriteInt(Y)

    '<EhFooter>
    Exit Function

PrepareMessageObjectDelete_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageObjectDelete " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal X As Byte, _
                                            ByVal Y As Byte, _
                                            ByVal Blocked As Boolean) As String

    '<EhHeader>
    On Error GoTo PrepareMessageBlockPosition_Err

    '</EhHeader>

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/08/07
    'Prepares the "BlockPosition" message and returns it
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.BlockPosition)
    Call Writer.WriteInt(X)
    Call Writer.WriteInt(Y)
    Call Writer.WriteBool(Blocked)
    
    '<EhFooter>
    Exit Function

PrepareMessageBlockPosition_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageBlockPosition " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectCreate(ByVal ObjIndex As Integer, _
                                           ByVal GrhIndex As Long, _
                                           ByVal X As Byte, _
                                           ByVal Y As Byte, _
                                           ByVal Name As String, _
                                           ByVal Amount As Integer, _
                                           ByVal Sound As Integer) As String

    '<EhHeader>
    On Error GoTo PrepareMessageObjectCreate_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'prepares the "ObjectCreate" message and returns it
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.ObjectCreate)
    Call Writer.WriteInt(X)
    Call Writer.WriteInt(Y)
    Call Writer.WriteInt(GrhIndex)
    Call Writer.WriteInt16(ObjIndex)
    Call Writer.WriteInt(Amount)
    Call Writer.WriteInt16(Sound)

    '<EhFooter>
    Exit Function

PrepareMessageObjectCreate_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageObjectCreate " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal charindex As Integer) As String

    '<EhHeader>
    On Error GoTo PrepareMessageCharacterRemove_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterRemove" message and returns it
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.CharacterRemove)
    Call Writer.WriteInt(charindex)

    '<EhFooter>
    Exit Function

PrepareMessageCharacterRemove_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageCharacterRemove " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal charindex As Integer) As String

    '<EhHeader>
    On Error GoTo PrepareMessageRemoveCharDialog_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.RemoveCharDialog)
    Call Writer.WriteInt(charindex)

    '<EhFooter>
    Exit Function

PrepareMessageRemoveCharDialog_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageRemoveCharDialog " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Writes the "CharacterCreate" message to the given user's outgoing data Reader.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    NickColor Determines if the character is a criminal or not, and if can be atacked by someone
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterCreate(ByVal Body As Integer, _
                                              ByVal BodyAttack As Integer, _
                                              ByVal Head As Integer, _
                                              ByVal Heading As eHeading, _
                                              ByVal charindex As Integer, _
                                              ByVal X As Byte, _
                                              ByVal Y As Byte, _
                                              ByVal Weapon As Integer, _
                                              ByVal Shield As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer, _
                                              ByVal Name As String, _
                                              ByVal NickColor As Byte, _
                                              ByVal Privileges As Byte, _
                                              ByRef AuraIndex() As Byte, _
                                              ByVal NpcIndex As Integer, _
                                              ByVal Idle As Boolean, _
                                              ByVal Navegando As Boolean, _
                                              ByVal speeding As Single) As String

    '<EhHeader>
    On Error GoTo PrepareMessageCharacterCreate_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterCreate" message and returns it
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.CharacterCreate)
        
    Call Writer.WriteInt(charindex)
    Call Writer.WriteInt(Body)
    Call Writer.WriteInt(BodyAttack)
    Call Writer.WriteInt(Head)
    Call Writer.WriteInt(Heading)
    Call Writer.WriteInt(X)
    Call Writer.WriteInt(Y)
    Call Writer.WriteInt(Weapon)
    Call Writer.WriteInt(Shield)
    Call Writer.WriteInt(helmet)
    Call Writer.WriteInt(FX)
    Call Writer.WriteInt(FXLoops)
    Call Writer.WriteString8(Name)
    Call Writer.WriteInt(NickColor)
    Call Writer.WriteInt(Privileges)
          
    Dim A As Long
          
    For A = 1 To MAX_AURAS
        Call Writer.WriteInt(AuraIndex(A))
    Next A

    Call Writer.WriteInt16(NpcIndex)

    Dim flags As Byte

    flags = 0
        
    If Idle Then flags = flags Or &O1 ' 00000001
    If Navegando Then flags = flags Or &O2
    Call Writer.WriteInt8(flags)
    Call Writer.WriteReal32(speeding)
    '<EhFooter>
    Exit Function

PrepareMessageCharacterCreate_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageCharacterCreate " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "CharacterChange" message and returns it.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterChange(ByVal Body As Integer, _
                                              ByVal BodyAttack As Integer, _
                                              ByVal Head As Integer, _
                                              ByVal Heading As eHeading, _
                                              ByVal charindex As Integer, _
                                              ByVal Weapon As Integer, _
                                              ByVal Shield As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer, _
                                              ByRef AuraIndex() As Byte, _
                                              ByVal ModoStreamer As Boolean, _
                                              ByVal Idle As Boolean, _
                                              ByVal Navegando As Boolean) As String

    '<EhHeader>
    On Error GoTo PrepareMessageCharacterChange_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterChange" message and returns it
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.CharacterChange)
        
    Call Writer.WriteInt(charindex)
    Call Writer.WriteInt(Body)
    Call Writer.WriteInt(BodyAttack)
    Call Writer.WriteInt(Head)
    Call Writer.WriteInt(Heading)
    Call Writer.WriteInt(Weapon)
    Call Writer.WriteInt(Shield)
    Call Writer.WriteInt(helmet)
    Call Writer.WriteInt(FX)
    Call Writer.WriteInt(FXLoops)
          
    Dim A As Long

    For A = 1 To MAX_AURAS
        Call Writer.WriteInt(AuraIndex(A))
    Next A
          
    Call Writer.WriteBool(ModoStreamer)

    Dim flags As Byte

    flags = 0

    If Idle Then flags = flags Or &O1
    If Navegando Then flags = flags Or &O2
    Call Writer.WriteInt8(flags)
    '<EhFooter>
    Exit Function

PrepareMessageCharacterChange_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageCharacterChange " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "CharacterChangeHeading" message and returns it.
'

Public Function PrepareMessageCharacterChangeHeading(ByVal charindex As Integer, _
                                                     ByVal Heading As eHeading) As String

    '<EhHeader>
    On Error GoTo PrepareMessageCharacterChangeHeading_Err

    '</EhHeader>

    '***************************************************
    'Author:
    'Last Modification:
    'Prepares the "CharacterChangeHeading" message and returns it
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.CharacterChangeHeading)
        
    Call Writer.WriteInt(charindex)
    Call Writer.WriteInt(Heading)

    '<EhFooter>
    Exit Function

PrepareMessageCharacterChangeHeading_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageCharacterChangeHeading " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal charindex As Integer, _
                                            ByVal X As Byte, _
                                            ByVal Y As Byte) As String

    '<EhHeader>
    On Error GoTo PrepareMessageCharacterMove_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterMove" message and returns it
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.CharacterMove)
    Call Writer.WriteInt(charindex)
    Call Writer.WriteInt(X)
    Call Writer.WriteInt(Y)

    '<EhFooter>
    Exit Function

PrepareMessageCharacterMove_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageCharacterMove " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As String

    '<EhHeader>
    On Error GoTo PrepareMessageForceCharMove_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/03/2009
    'Prepares the "ForceCharMove" message and returns it
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.ForceCharMove)
    Call Writer.WriteInt(Direccion)

    '<EhFooter>
    Exit Function

PrepareMessageForceCharMove_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageForceCharMove " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, _
                                                 ByVal NickColor As Byte, _
                                                 ByRef Tag As String) As String

    '<EhHeader>
    On Error GoTo PrepareMessageUpdateTagAndStatus_Err

    '</EhHeader>

    '***************************************************
    'Author: Alejandro Salvo (Salvito)
    'Last Modification: 04/07/07
    'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
    'Prepares the "UpdateTagAndStatus" message and returns it
    '15/01/2010: ZaMa - Now sends the nick color instead of the status.
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.UpdateTagAndStatus)
        
    Call Writer.WriteInt(UserList(UserIndex).Char.charindex)
    Call Writer.WriteInt(NickColor)
    Call Writer.WriteString8(Tag)
        
    '<EhFooter>
    Exit Function

PrepareMessageUpdateTagAndStatus_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageUpdateTagAndStatus " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal Message As String) As String

    '<EhHeader>
    On Error GoTo PrepareMessageErrorMsg_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ErrorMsg" message and returns it
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.ErrorMsg)
    Call Writer.WriteString8(Message)

    '<EhFooter>
    Exit Function

PrepareMessageErrorMsg_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageErrorMsg " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Writes the "CancelOfferItem" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    Slot      The slot to cancel.

Public Sub WriteCancelOfferItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

    '<EhHeader>
    On Error GoTo WriteCancelOfferItem_Err

    '</EhHeader>

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 05/03/2010
    '
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.CancelOfferItem)
    Call Writer.WriteInt(Slot)
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteCancelOfferItem_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteCancelOfferItem " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "SetDialog" message.
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSetDialog(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleSetDialog_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 18/11/2010
    '20/11/2010: ZaMa - Arreglo privilegios.
    '***************************************************

    With UserList(UserIndex)
        
        Dim NewDialog As String

        NewDialog = Reader.ReadString8
        
        If .flags.TargetNPC > 0 Then

            ' Dsgm/Dsrm/Rm
            If EsGmPriv(UserIndex) Then
                'Replace the NPC's dialog.
                Npclist(.flags.TargetNPC).Desc = NewDialog

            End If

        End If

    End With
    
    '<EhFooter>
    Exit Sub

HandleSetDialog_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleSetDialog " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Impersonate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleImpersonate(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleImpersonate_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    '
    '***************************************************
    With UserList(UserIndex)
        
        ' Dsgm/Dsrm/Rm
        If Not EsGmPriv(UserIndex) Then Exit Sub
        
        Dim NpcIndex As Integer

        NpcIndex = .flags.TargetNPC
        
        If NpcIndex = 0 Then Exit Sub
        
        ' Copy head, body and desc
        Call ImitateNpc(UserIndex, NpcIndex)
        
        ' Teleports user to npc's coords
        Call WarpUserChar(UserIndex, Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, False, True)
        
        ' Log gm
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/IMPERSONAR con " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)
        
        ' Remove npc
        Call QuitarNPC(NpcIndex)
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleImpersonate_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleImpersonate " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "Imitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleImitate(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleImitate_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    '
    '***************************************************
    
    If Not EsGmPriv(UserIndex) Then Exit Sub
    
    With UserList(UserIndex)

        Dim NpcIndex As Integer

        NpcIndex = .flags.TargetNPC
        
        If NpcIndex = 0 Then Exit Sub
        
        ' Copy head, body and desc
        Call ImitateNpc(UserIndex, NpcIndex)
        Call Logs_User(.Name, eLog.eGm, eLogDescUser.eNone, "/MIMETIZAR con " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleImitate_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleImitate " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "RecordAdd" message.
'
' @param UserIndex The index of the user sending the message
           
Public Sub HandleRecordAdd(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRecordAdd_Err

    '</EhHeader>

    '**************************************************************
    'Author: Amraphen
    'Last Modify Date: 29/11/2010
    '
    '**************************************************************

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim Reason   As String
        
        UserName = Reader.ReadString8
        Reason = Reader.ReadString8
    
        If Not (.flags.Privilegios And (PlayerType.User)) Then

            'Verificamos que exista el personaje
            If Not FileExist(CharPath & UCase$(UserName) & ".chr") Then
                Call WriteShowMessageBox(UserIndex, "El personaje no existe")
            Else
                'Agregamos el seguimiento
                Call AddRecord(UserIndex, UserName, Reason)
                
                'Enviamos la nueva lista de personajes
                Call WriteRecordList(UserIndex)

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleRecordAdd_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRecordAdd " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "RecordAddObs" message.
'
' @param UserIndex The index of the user sending the message.

Public Sub HandleRecordAddObs(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRecordAddObs_Err

    '</EhHeader>

    '**************************************************************
    'Author: Amraphen
    'Last Modify Date: 29/11/2010
    '
    '**************************************************************

    With UserList(UserIndex)
        
        Dim RecordIndex As Byte

        Dim Obs         As String
        
        RecordIndex = Reader.ReadInt
        Obs = Reader.ReadString8
        
        If Not (.flags.Privilegios And (PlayerType.User)) Then
            'Agregamos la observación
            Call AddObs(UserIndex, RecordIndex, Obs)
            
            'Actualizamos la información
            Call WriteRecordDetails(UserIndex, RecordIndex)

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleRecordAddObs_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRecordAddObs " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "RecordRemove" message.
'
' @param UserIndex The index of the user sending the message.

Public Sub HandleRecordRemove(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRecordRemove_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    '
    '***************************************************
    Dim RecordIndex As Integer

    With UserList(UserIndex)
    
        RecordIndex = Reader.ReadInt
        
        If .flags.Privilegios And (PlayerType.User) Then Exit Sub
        
        'Sólo dioses pueden remover los seguimientos, los otros reciben una advertencia:
        If (.flags.Privilegios And PlayerType.Dios) Then
            Call RemoveRecord(RecordIndex)
            Call WriteShowMessageBox(UserIndex, "Se ha eliminado el seguimiento.")
            Call WriteRecordList(UserIndex)
        Else
            Call WriteShowMessageBox(UserIndex, "Sólo los dioses pueden eliminar seguimientos.")

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleRecordRemove_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRecordRemove " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "RecordListRequest" message.
'
' @param UserIndex The index of the user sending the message.
            
Public Sub HandleRecordListRequest(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRecordListRequest_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    '
    '***************************************************
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User) Then Exit Sub

        Call WriteRecordList(UserIndex)

    End With

    '<EhFooter>
    Exit Sub

HandleRecordListRequest_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRecordListRequest " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "RecordDetails" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordDetails(ByVal UserIndex As Integer, ByVal RecordIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteRecordDetails_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    'Writes the "RecordDetails" message to the given user's outgoing data buffer
    '***************************************************
    Dim i        As Long

    Dim tIndex   As Integer

    Dim tmpStr   As String

    Dim TempDate As Date

    Call Writer.WriteInt(ServerPacketID.RecordDetails)
        
    'Creador y motivo
    Call Writer.WriteString8(Records(RecordIndex).Creador)
    Call Writer.WriteString8(Records(RecordIndex).Motivo)
        
    tIndex = NameIndex(Records(RecordIndex).Usuario)
        
    'Status del pj (online?)
    Call Writer.WriteBool(tIndex > 0)
        
    'Escribo la IP según el estado del personaje
    If tIndex > 0 Then
        'La IP Actual
        tmpStr = UserList(tIndex).IpAddress
    Else 'String nulo
        tmpStr = vbNullString

    End If

    Call Writer.WriteString8(tmpStr)
        
    'Escribo tiempo online según el estado del personaje
    If tIndex > 0 Then
        'Tiempo logueado.
        TempDate = Now - UserList(tIndex).LogOnTime
        tmpStr = Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate)
    Else
        'Envío string nulo.
        tmpStr = vbNullString

    End If

    Call Writer.WriteString8(tmpStr)

    'Escribo observaciones:
    tmpStr = vbNullString

    If Records(RecordIndex).NumObs Then

        For i = 1 To Records(RecordIndex).NumObs
            tmpStr = tmpStr & Records(RecordIndex).Obs(i).Creador & "> " & Records(RecordIndex).Obs(i).Detalles & vbCrLf
        Next i
            
        tmpStr = Left$(tmpStr, Len(tmpStr) - 1)

    End If

    Call Writer.WriteString8(tmpStr)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteRecordDetails_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteRecordDetails " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "RecordList" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordList(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteRecordList_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    'Writes the "RecordList" message to the given user's outgoing data buffer
    '***************************************************
    Dim i As Long
    
    Call Writer.WriteInt(ServerPacketID.RecordList)
        
    Call Writer.WriteInt(NumRecords)

    For i = 1 To NumRecords
        Call Writer.WriteString8(Records(i).Usuario)
    Next i

    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteRecordList_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteRecordList " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "ShowMenu" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    MenuIndex: The menu index.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMenu(ByVal UserIndex As Integer, ByVal MenuIndex As Byte)

    '<EhHeader>
    On Error GoTo WriteShowMenu_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 10/05/2011
    'Writes the "ShowMenu" message to the given user's outgoing data buffer
    '***************************************************
    Dim i As Long

    Call Writer.WriteInt(ServerPacketID.ShowMenu)
        
    Call Writer.WriteInt(MenuIndex)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteShowMenu_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteShowMenu " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "RecordDetailsRequest" message.
'
' @param UserIndex The index of the user sending the message.
            
Public Sub HandleRecordDetailsRequest(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRecordDetailsRequest_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 07/04/2011
    'Handles the "RecordListRequest" message
    '***************************************************
    Dim RecordIndex As Byte

    With UserList(UserIndex)
        
        RecordIndex = Reader.ReadInt
        
        If .flags.Privilegios And (PlayerType.User) Then Exit Sub
        
        Call WriteRecordDetails(UserIndex, RecordIndex)

    End With

    '<EhFooter>
    Exit Sub

HandleRecordDetailsRequest_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRecordDetailsRequest " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleMoveItem(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleMoveItem_Err

    '</EhHeader>

    '***************************************************
    'Author: Ignacio Mariano Tirabasso (Budi)
    'Last Modification: 01/01/2011
    '
    '***************************************************
    
    With UserList(UserIndex)

        Dim originalSlot As Byte

        Dim newSlot      As Byte
        
        Dim Tipo         As Byte
        
        Dim TypeBank     As Byte
        
        originalSlot = Reader.ReadInt
        newSlot = Reader.ReadInt
        Tipo = Reader.ReadInt
        TypeBank = Reader.ReadInt
        
        If Tipo = eMoveType.Inventory Then
            Call InvUsuario.moveItem(UserIndex, originalSlot, newSlot)
        ElseIf Tipo = eMoveType.Bank Then
            Call InvUsuario.MoveItem_Bank(UserIndex, originalSlot, newSlot, TypeBank)
            
        End If

    End With

    '<EhFooter>
    Exit Sub

HandleMoveItem_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleMoveItem " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function PrepareMessageCharacterAttackMovement(ByVal charindex As Integer) As String

    '<EhHeader>
    On Error GoTo PrepareMessageCharacterAttackMovement_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 24/05/2011
    'Prepares the "CharacterAttackMovement" message and returns it.
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.CharacterAttackMovement)
    Call Writer.WriteInt(charindex)

    '<EhFooter>
    Exit Function

PrepareMessageCharacterAttackMovement_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageCharacterAttackMovement " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function PrepareMessageCharacterAttackNpc(ByVal charindex As Integer, _
                                                 ByVal BodyAttack As Integer) As String

    '<EhHeader>
    On Error GoTo PrepareMessageCharacterAttackNpc_Err

    '</EhHeader>

    '***************************************************
    'Author: Lautarito
    'Last Modification: 09/05/2020
    'Prepares the "CharacterAttackNpc" message and returns it.
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.CharacterAttackNpc)
    Call Writer.WriteInt(charindex)

    '<EhFooter>
    Exit Function

PrepareMessageCharacterAttackNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageCharacterAttackNpc " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Writes the "StrDextRunningOut" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @param    Seconds Seconds left.

Public Sub WriteStrDextRunningOut(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteStrDextRunningOut_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Dalmasso (CHOTS)
    'Last Modification: 08/06/2011
    '
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.StrDextRunningOut)
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteStrDextRunningOut_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteStrDextRunningOut " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Handles the "PMSend" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandlePMSend(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandlePMSend_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 04/08/2011
    'Handles the "PMSend" message.
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName    As String

        Dim Message     As String

        Dim TargetIndex As Integer

        UserName = Reader.ReadString8
        Message = Reader.ReadString8
        
        TargetIndex = NameIndex(UserName)
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            If TargetIndex = 0 Then 'Offline
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    Call AgregarMensajeOFF(UserName, .Name, Message)
                    Call WriteConsoleMsg(UserIndex, "Mensaje enviado.", FontTypeNames.FONTTYPE_GM)
                Else
                    Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else 'Online
                Call AgregarMensaje(TargetIndex, .Name, Message)
                Call WriteConsoleMsg(UserIndex, "Mensaje enviado.", FontTypeNames.FONTTYPE_GM)

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandlePMSend_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandlePMSend " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleSearchObj(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleSearchObj_Err

    '</EhHeader>

    '***************************************************
    'Author: WAICON
    'Last Modification: 06/05/2019
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim Tag     As String

        Dim A       As Long

        Dim cant    As Long

        Dim strTemp As String
        
        Tag = Reader.ReadString8
        
        If .flags.Privilegios And (PlayerType.Admin) Then

            For A = 1 To UBound(ObjData)

                If InStr(1, Tilde(ObjData(A).Name), Tilde(Tag)) Then
                    strTemp = strTemp & A & " " & ObjData(A).Name & vbCrLf
                    
                    cant = cant + 1

                End If

            Next
            
            If cant = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hubo resultados de: '" & Tag & "'", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Hubo " & cant & " resultados de: " & Tag & strTemp, FontTypeNames.FONTTYPE_INFOBOLD)

            End If
        
        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleSearchObj_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleSearchObj " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleUserEditation(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleUserEditation_Err

    '</EhHeader>
    
    With UserList(UserIndex)
        
        Dim Elv As Byte
            
        Select Case .Account.Premium
            
            Case 0
                Exit Sub

            Case 1
                Elv = 30

            Case 2
                Elv = 35

            Case 3
                Elv = 40

        End Select
            
        If .Stats.Elv >= Elv Or .Stats.Elv < 3 Then
            Call WriteConsoleMsg(UserIndex, "Debes ser nivel inferior a " & Elv & " para poder reiniciar tu personaje de acuerdo al Tier elegido.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If
        
        If .GroupIndex > 0 Then
            Call SaveExpAndGldMember(.GroupIndex, UserIndex)
            Call WriteConsoleMsg(UserIndex, "Dirigete a una zona libre de party.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If
            
        If .flags.Navegando > 0 Then
            Call WriteConsoleMsg(UserIndex, "Deja de navegar y podrás reiniciar tu personaje.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If
            
        If MapInfo(.Pos.Map).Pk Then
            Call WriteConsoleMsg(UserIndex, "¡Vete a Zona Segura! Aquí corres peligro...", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If
            
        If .MascotaIndex > 0 Then
            Call QuitarPet(UserIndex, .MascotaIndex)

        End If
            
        ' Tiene objetos, los desequipamos
        Call Reset_DesquiparAll(UserIndex)

        Call InitialUserStats(UserList(UserIndex))
            
        'Call QuitarNewbieObj(UserIndex)
            
        Call LimpiarInventario(UserIndex)
        Call ApplySetInitial_Newbie(UserIndex)
        Call UpdateUserInv(True, UserIndex, 0)
            
        If MapInfo(.Pos.Map).LvlMin > .Stats.Elv Then
            Call WriteConsoleMsg(UserIndex, "Hemos notado que no puedes sobrevivir a la peligrosidad de este mapa. ¡Serás llevado a Ullathorpe! ¡No nos lo agradezcas!", FontTypeNames.FONTTYPE_INFORED)
            Call WarpUserChar(UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, True)

        End If
        
        Call WriteLevelUp(UserIndex, .Stats.SkillPts)
        Call WriteUpdateUserStats(UserIndex)
        Call WriteUpdateStrenghtAndDexterity(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has reiniciado tu personaje. ¡Que tengas un excelente re-comienzo!", FontTypeNames.FONTTYPE_GUILDMSG)

    End With
    
    '<EhFooter>
    Exit Sub

HandleUserEditation_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleUserEditation " & "at line " & Erl

    '</EhFooter>
End Sub

Private Sub HandlePartyClient(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandlePartyClient_Err

    '</EhHeader>

    ' 1) Requiere formulario 'principal'
    ' 4) Abandonar party
    ' 5) Requiere ingresar a party
    
    Dim Paso As Byte
    
    With UserList(UserIndex)
        
        Select Case Reader.ReadInt

            Case 1

                If .GroupIndex <= 0 Then
                    mGroup.CreateGroup (UserIndex)
                Else
                    WriteGroupPrincipal (UserIndex)

                End If

            Case 2 ' Cambia la obtención de Experiencia, para ver si recibe por golpe o acumula...

                If .GroupIndex > 0 Then
                    mGroup.ChangeObtainExp UserIndex
                
                End If
                
            Case 3
                mGroup.AcceptInvitationGroup UserIndex

            Case 4

                If .GroupIndex > 0 Then
                    mGroup.AbandonateGroup UserIndex

                End If
            
            Case 5
                mGroup.SendInvitationGroup UserIndex

        End Select

    End With

    '<EhFooter>
    Exit Sub

HandlePartyClient_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandlePartyClient " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteGroupUpdateExp(ByVal UserIndex As Integer, ByVal GroupIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteGroupUpdateExp_Err

    '</EhHeader>
    Call Writer.WriteInt(ServerPacketID.GroupUpdateExp)

    Dim A As Long
     
    With Groups(GroupIndex)

        For A = 1 To MAX_MEMBERS_GROUP
            Call Writer.WriteInt32(.User(A).Exp)
        Next A

    End With
     
    '<EhFooter>
    Exit Sub

WriteGroupUpdateExp_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteGroupUpdateExp " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteGroupPrincipal(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteGroupPrincipal_Err

    '</EhHeader>

    Dim GroupIndex As Integer

    Dim A          As Long, B As Long
    
    Call Writer.WriteInt(ServerPacketID.GroupPrincipal)
    
    GroupIndex = UserList(UserIndex).GroupIndex
    
    With Groups(GroupIndex)
        Call Writer.WriteBool(.Acumular)
        
        For A = 1 To MAX_MEMBERS_GROUP

            If .User(A).Index > 0 Then
                Call Writer.WriteString8(UserList(.User(A).Index).Name)
            Else
                Call Writer.WriteString8("<Vacio>")

            End If
            
            Call Writer.WriteInt8(.User(A).PorcExp)
            Call Writer.WriteInt32(.User(A).Exp)
        Next A

    End With
              
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteGroupPrincipal_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteGroupPrincipal " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function CheckValidPorc(ByVal UserIndex As Integer, ByRef Exp() As Byte) As Boolean

    On Error GoTo CheckValidPorc_Err
    
    Dim A    As Long

    Dim Porc As Long
    
    With UserList(UserIndex)

        If .Invent.PendientePartyObjIndex = 0 Then
            CheckValidPorc = False
            Exit Function

        End If
            
        Porc = ObjData(.Invent.PendientePartyObjIndex).Porc
            
        For A = LBound(Exp) To UBound(Exp)

            If Exp(A) > Porc Then
                Exit Function

            End If

        Next A

    End With

    CheckValidPorc = True
    '<EhFooter>
    Exit Function

CheckValidPorc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.CheckValidPorc " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub HandleGroupChangePorc(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleGroupChangePorc_Err

    '</EhHeader>
    
    Dim A      As Byte

    Dim Exp(4) As Byte
    
    With UserList(UserIndex)
        
        For A = 0 To 4
            Exp(A) = Reader.ReadInt
        Next A
        
        If .GroupIndex > 0 Then

            If Not CheckValidPorc(UserIndex, Exp) Then
                Call WriteConsoleMsg(UserIndex, "¡No tienes ningún Pendiente de Experiencia o bien no permite cambiar al porcentaje seleccionado!", FontTypeNames.FONTTYPE_ANGEL)
                Exit Sub

            End If

            mGroup.GroupSetPorcentaje UserIndex, .GroupIndex, Exp

        End If
        
    End With

    '<EhFooter>
    Exit Sub

HandleGroupChangePorc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGroupChangePorc " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub WriteUserInEvent(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUserInEvent_Err

    '</EhHeader>

    Call Writer.WriteInt(ServerPacketID.UserInEvent)

    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteUserInEvent_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUserInEvent " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleEntrarDesafio(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleEntrarDesafio_Err

    '</EhHeader>
    
    With UserList(UserIndex)
        
        Select Case Reader.ReadInt
            
            Case 0
            
                Call mDesafios.Desafio_UserAdd(UserIndex)

            Case 1

                'If .flags.Desafiando Then
                'Call Desafio_UserKill(UserIndex)
                'End If
        End Select
    
    End With

    '<EhFooter>
    Exit Sub

HandleEntrarDesafio_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleEntrarDesafio " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Writes the "MontateToggle" message to the given user's outgoing data Reader.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMontateToggle(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteMontateToggle_Err

    '</EhHeader>

    '***************************************************
    'Author: Dragons
    'Last Modification: 30/06/2019
    'Writes the "MontateToggle" message to the given user's outgoing data buffer
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.MontateToggle)

    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteMontateToggle_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteMontateToggle " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleSetPanelClient(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleSetPanelClient_Err

    '</EhHeader>
    
    With UserList(UserIndex)
        
        .flags.MenuCliente = Reader.ReadInt
        .flags.LastSlotClient = Reader.ReadInt

        Dim X As Long, Y As Long
              
        X = Reader.ReadInt
        Y = Reader.ReadInt
              
        If Not (X = 0 And Y = 0) Then
            UpdatePointer UserIndex, .flags.MenuCliente, X, Y, "Solapas Inv-Hec"

        End If
              
        Reader.ReadInt16

    End With

    '<EhFooter>
    Exit Sub

HandleSetPanelClient_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleSetPanelClient " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleSolicitaSeguridad(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleSolicitaSeguridad_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim Tipo     As Byte

        Dim TempName As String, TempHD As String
        
        UserName = Reader.ReadString8
        Tipo = Reader.ReadInt
        
        If CharIs_Admin(UCase$(.Name)) Then
            
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                WriteConsoleMsg UserIndex, "El usuario se ha desconectado.", FontTypeNames.FONTTYPE_INFO
            Else

                If EsGm(tUser) Then
                    WriteConsoleMsg UserIndex, "No puedes ver la información de otros GameMaster", FontTypeNames.FONTTYPE_INFO
                Else

                    Select Case Tipo

                            ' Inicia el Seguimiento
                        Case 0

                            If UserList(tUser).flags.GmSeguidor > 0 Then
                                WriteConsoleMsg UserList(tUser).flags.GmSeguidor, "El GM " & .Name & " ha comenzado a analizar al personaje " & UserList(tUser).Name, FontTypeNames.FONTTYPE_INFORED
                                UserList(tUser).flags.GmSeguidor = UserIndex
                            Else
                                UserList(tUser).flags.GmSeguidor = UserIndex
                                WriteSolicitaCapProc tUser, 0

                            End If
                            
                            Call WriteUpdateListSecurity(UserList(tUser).flags.GmSeguidor, UserList(tUser).Name, vbNullString, 255)
                            
                        Case 1 ' Actualiza la solapa de procesos
                            WriteSolicitaCapProc tUser, 1
                            
                        Case 2 ' Actualiza la solapa de captions
                            WriteSolicitaCapProc tUser, 2
                        
                        Case 3, 4, 5

                            If .flags.Privilegios And (PlayerType.Admin) Then
                                WriteSolicitaCapProc tUser, Tipo

                            End If
                        
                        Case Else

                    End Select
                    
                End If

            End If
        
        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleSolicitaSeguridad_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleSolicitaSeguridad " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleSendListSecurity(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleSendListSecurity_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        Dim List As String

        Dim Tipo As Byte
        
        List = Reader.ReadString8
        Tipo = Reader.ReadInt
        
        If .flags.GmSeguidor > 0 Then
            Call WriteUpdateListSecurity(.flags.GmSeguidor, .Name, List, Tipo)

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleSendListSecurity_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleSendListSecurity " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteUpdateListSecurity(ByVal UserIndex As Integer, _
                                   ByVal CheaterName As String, _
                                   ByVal List As String, _
                                   ByVal Tipo As Byte)

    '<EhHeader>
    On Error GoTo WriteUpdateListSecurity_Err

    '</EhHeader>
    
    Call Writer.WriteInt(ServerPacketID.UpdateListSecurity)
    Call Writer.WriteString8(CheaterName)
    Call Writer.WriteString8(List)
    Call Writer.WriteInt(Tipo)
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteUpdateListSecurity_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateListSecurity " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteSolicitaCapProc(ByVal UserIndex As Integer, _
                                Optional ByVal Tipo As Byte = 0, _
                                Optional ByVal Process As String = vbNullString, _
                                Optional ByVal Captions As String = vbNullString)

    '<EhHeader>
    On Error GoTo WriteSolicitaCapProc_Err

    '</EhHeader>

    Call Writer.WriteInt(ServerPacketID.SolicitaCapProc)
    Call Writer.WriteInt(Tipo)
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteSolicitaCapProc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteSolicitaCapProc " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function PrepareMessageCreateDamage(ByVal X As Byte, _
                                           ByVal Y As Byte, _
                                           ByVal DamageValue As Long, _
                                           ByVal DamageType As eDamageType, _
                                           Optional ByVal Text As String = vbNullString)

    '<EhHeader>
    On Error GoTo PrepareMessageCreateDamage_Err

    '</EhHeader>
 
    Writer.WriteInt ServerPacketID.CreateDamage
    Writer.WriteInt8 X
    Writer.WriteInt8 Y
    Writer.WriteInt32 DamageValue
    Writer.WriteInt8 DamageType

    If DamageType = d_AddMagicWord Then
        Writer.WriteString8 Text

    End If

    '<EhFooter>
    Exit Function

PrepareMessageCreateDamage_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageCreateDamage " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function WriteVesA(ByVal UserIndex As Integer, ByVal Name As String, ByVal Desc As String, ByVal Class As eClass, ByVal Raza As eRaza, ByVal Faction As Byte, ByVal FactionRange As String, ByVal GuildName As String, ByVal GuildRange As Byte, ByVal RangeGm As String, ByVal sPlayerType As Byte, ByVal IsGold As Byte, ByVal IsBronce As Byte, ByVal IsPlata As Byte, ByVal IsPremium As Byte, ByVal IsStreamer As Byte, ByVal IsTransform As Byte, ByVal IsKilled As Byte, ByVal FtOptional As FontTypeNames, ByVal StreamerUrl As String, ByVal Rachas As Integer, ByVal RachasHist As Integer) As String

    '<EhHeader>
    On Error GoTo WriteVesA_Err

    '</EhHeader>

    Call Writer.WriteInt(ServerPacketID.ClickVesA)
        
    Call Writer.WriteString8(Name)
    Call Writer.WriteString8(Desc)
    Call Writer.WriteInt(Class)
    Call Writer.WriteInt(Raza)
    Call Writer.WriteInt(Faction)
    Call Writer.WriteString8(FactionRange)
    Call Writer.WriteString8(GuildName)
    Call Writer.WriteInt(GuildRange)
        
    Call Writer.WriteString8(RangeGm)
    Call Writer.WriteInt(sPlayerType)
        
    Call Writer.WriteInt(IsGold)
    Call Writer.WriteInt(IsBronce)
    Call Writer.WriteInt(IsPlata)
    Call Writer.WriteInt(IsPremium)
    Call Writer.WriteInt(IsStreamer)
    Call Writer.WriteInt(IsTransform)
    Call Writer.WriteInt(IsKilled)
    Call Writer.WriteInt(FtOptional)
    Call Writer.WriteString8(StreamerUrl)
    Call Writer.WriteInt16(Rachas)
    Call Writer.WriteInt16(RachasHist)
    Call SendData(ToOne, UserIndex, vbNullString)
   
    '<EhFooter>
    Exit Function

WriteVesA_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteVesA " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub HandleCheckingGlobal(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleCheckingGlobal_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            
            If GlobalActive Then
                GlobalActive = False
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El chat global ha sido desactivado.", FontTypeNames.FONTTYPE_INFO))
            Else
                GlobalActive = True
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El chat global ha sido activado. Utiliza el comando /GLOBAL para hablar con los demás usuarios del juego.", FontTypeNames.FONTTYPE_GUILD))

            End If

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleCheckingGlobal_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCheckingGlobal " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleChatGlobal(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChatGlobal_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        Dim Message As String

        Message = Reader.ReadString8
        
        If .flags.Streamer = 1 Then
            If GlobalActive = False Then
                Call WriteConsoleMsg(UserIndex, "El Chat Global se encuentra desactivado.", FontTypeNames.FONTTYPE_INFO)
            ElseIf .Counters.TimeGlobal > 0 Then
                Call WriteConsoleMsg(UserIndex, "Debes esperar algunos segundos para volver a enviar un mensaje al global", FontTypeNames.FONTTYPE_INFO)
            ElseIf .Counters.Pena > 0 Then
                Call WriteConsoleMsg(UserIndex, "No puedes enviar mensajes desde la cárcel", FontTypeNames.FONTTYPE_INFO)
            ElseIf .flags.Silenciado > 0 Then
                Call WriteConsoleMsg(UserIndex, "Los administradores te han silenciado. No podrás enviar mensajes al Chat Global", FontTypeNames.FONTTYPE_INFO)
            ElseIf Not AsciiValidos_Chat(Message) Then
                Call WriteConsoleMsg(UserIndex, "Mensaje inválido.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg(.Name & "» " & Message, FontTypeNames.FONTTYPE_GLOBAL))
            
                .Counters.TimeGlobal = 3
            
            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Solo los personajes considerados STREAMERS OFICIALES pueden utilizar este comando.", FontTypeNames.FONTTYPE_INFO)

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleChatGlobal_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChatGlobal " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleCountDown(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleCountDown_Err

    '</EhHeader>
    
    With UserList(UserIndex)
        
        Dim Count    As Byte

        Dim CountMap As Boolean
        
        Count = Reader.ReadInt + 1
        CountMap = Reader.ReadBool
        
        If Count > 240 Then Exit Sub
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            
            If Count = 0 Then
                CountDown_Map = 0
                CountDown_Time = 0

                Exit Sub

            End If
            
            If CountMap Then
                CountDown_Map = .Pos.Map
                CountDown_Time = Count
            Else
                CountDown_Time = Count
                CountDown_Map = 0

            End If
            
        End If
    
    End With

    '<EhFooter>
    Exit Sub

HandleCountDown_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleCountDown " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleGiveBackUser(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleGiveBackUser_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = Reader.ReadString8
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                WriteConsoleMsg UserIndex, "El usuario está offline.", FontTypeNames.FONTTYPE_INFO
            Else

                If UserList(tUser).PosAnt.Map <> 0 Then
                    Call WarpPosAnt(tUser)

                End If

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleGiveBackUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGiveBackUser " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleLearnMeditation(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleLearnMeditation_Err

    '</EhHeader>
    
    Dim Tipo     As Byte

    Dim Selected As Byte
    
    With UserList(UserIndex)
        
        Tipo = Reader.ReadInt
        Selected = Reader.ReadInt
        
        If Selected < 0 Or Selected > MAX_MEDITATION Then Exit Sub
        
        Select Case Tipo

            Case 0 ' Aprender nueva / Reclamar

                If Selected = 0 Then Exit Sub
                
                'Call mMeditations.Meditation_AddNew(UserIndex, Selected)
                
            Case 1 ' Poner en uso
                Call mMeditations.Meditation_Select(UserIndex, Selected)

        End Select
    
    End With

    '<EhFooter>
    Exit Sub

HandleLearnMeditation_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleLearnMeditation " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteCreateDamage(ByVal UserIndex As Integer, _
                             ByVal X As Byte, _
                             ByVal Y As Byte, _
                             ByVal Value As Long, _
                             ByVal DamageType As eDamageType)

    '<EhHeader>
    On Error GoTo WriteCreateDamage_Err

    '</EhHeader>

    Call SendData(ToOne, UserIndex, PrepareMessageCreateDamage(X, Y, Value, DamageType))

    '<EhFooter>
    Exit Sub

WriteCreateDamage_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteCreateDamage " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleInfoEvento(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleInfoEvento_Err

    '</EhHeader>
    
    Dim A As Long
    
    With UserList(UserIndex)
        
        If Not Interval_Packet250(UserIndex) Then Exit Sub
            
        Call WriteTournamentList(UserIndex)
        
    End With

    '<EhFooter>
    Exit Sub

HandleInfoEvento_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleInfoEvento " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleDragToPos(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleDragToPos_Err

    '</EhHeader>

    ' @ Author : maTih.-
    '            Drag&Drop de objetos en del inventario a una posición.
            
    Dim X      As Byte

    Dim Y      As Byte

    Dim Slot   As Byte

    Dim Amount As Integer

    Dim tUser  As Integer

    Dim tNpc   As Integer

    X = Reader.ReadInt()
    Y = Reader.ReadInt()
    Slot = Reader.ReadInt()
    Amount = Reader.ReadInt()

    tUser = MapData(UserList(UserIndex).Pos.Map, X, Y).UserIndex

    tNpc = MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex

    If Not Interval_Drop(UserIndex) Then Exit Sub
    If UserList(UserIndex).flags.Comerciando Then Exit Sub
    If UserList(UserIndex).flags.Montando Then Exit Sub
    
    If Not InMapBounds(UserList(UserIndex).Pos.Map, X, Y) Then Exit Sub
    If MapData(UserList(UserIndex).Pos.Map, X, Y).Blocked = 1 Then Exit Sub
    If UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
    If Amount <= 0 Or Amount > UserList(UserIndex).Invent.Object(Slot).Amount Then Exit Sub
    If UserList(UserIndex).Invent.Object(Slot).ObjIndex <= 0 Then Exit Sub
    If tUser = UserIndex Then Exit Sub
    If EsGm(UserIndex) And Not EsGmPriv(UserIndex) Then Exit Sub
    
    If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).NoNada = 1 Then Exit Sub
    
    'If EsNewbie(UserIndex) Then
    'Call WriteConsoleMsg(UserIndex, "Los newbies no pueden dropear objetos.", FontTypeNames.FONTTYPE_INFO)
    'Exit Sub
    'End If
    
    If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).OBJType = otGemaTelep Then
        Call WriteConsoleMsg(UserIndex, "No puedes dropear este objeto.", FontTypeNames.FONTTYPE_INFO)

        Exit Sub

    End If
    
    If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).OBJType = otMonturas Then
        Call WriteConsoleMsg(UserIndex, "No puedes dropear este objeto.", FontTypeNames.FONTTYPE_INFO)

        Exit Sub

    End If
        
    If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).OBJType = otTransformVIP Then
        Call WriteConsoleMsg(UserIndex, "No puedes dropear este objeto.", FontTypeNames.FONTTYPE_INFO)

        Exit Sub

    End If
    
    If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).NoDrop = 1 Then
        Call WriteConsoleMsg(UserIndex, "No puedes dropear este objeto.", FontTypeNames.FONTTYPE_INFO)

        Exit Sub

    End If
        
    If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Plata = 1 Then
        Call WriteConsoleMsg(UserIndex, "No puedes dropear este objeto.", FontTypeNames.FONTTYPE_INFO)

        Exit Sub

    End If
        
    If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Oro = 1 Then
        Call WriteConsoleMsg(UserIndex, "No puedes dropear este objeto.", FontTypeNames.FONTTYPE_INFO)

        Exit Sub

    End If
        
    If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Premium = 1 Then
        Call WriteConsoleMsg(UserIndex, "No puedes dropear este objeto.", FontTypeNames.FONTTYPE_INFO)

        Exit Sub

    End If
    
    If UserList(UserIndex).flags.SlotEvent > 0 Or UserList(UserIndex).flags.SlotReto > 0 Then Exit Sub
    
    If tUser > 0 Then
        If tUser = UserIndex Then Exit Sub
        If EsGm(tUser) Then Exit Sub
         
        If UserList(tUser).flags.DragBlocked Then
            Call WriteConsoleMsg(UserIndex, "La persona no quiere tus objetos.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If UserList(tUser).flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "No puedes arrojar objetos si la persona está comerciando.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        Call mDragAndDrop.DragToUser(UserIndex, tUser, Slot, Amount)

    ElseIf tNpc > 0 Then
        Call mDragAndDrop.DragToNPC(UserIndex, tNpc, Slot, Amount)

    Else
        Call mDragAndDrop.DragToPos(UserIndex, X, Y, Slot, Amount)

    End If

    '<EhFooter>
    Exit Sub

HandleDragToPos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleDragToPos " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleAbandonateFaction(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleAbandonateFaction_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If

        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub

        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If

        If .Faction.Status = 0 Then
            Call WriteConsoleMsg(UserIndex, "¡No perteneces a ninguna facción!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        Call mFacciones.Faction_RemoveUser(UserIndex)

    End With

    '<EhFooter>
    Exit Sub

HandleAbandonateFaction_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleAbandonateFaction " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleEnlist(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleEnlist_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If

        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub

        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If

        If .Faction.Status > 0 Then
            Call WriteConsoleMsg(UserIndex, "Ya eres miembro de una facción y espero que sea la nuestra, sino mis guardias te atacaran!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            If Escriminal(UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "¡¡Sal de aquí, antes de que mis guardias acaben contigo!!", FontTypeNames.FONTTYPE_WARNING)

                Exit Sub

            End If
            
            Call mFacciones.Faction_AddUser(UserIndex, r_Armada)
            Call Guilds_CheckAlineation(UserIndex, a_Armada)
        Else

            If Not Escriminal(UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "¡¡Sal de aquí, antes de que mis guardias acaben contigo!!", FontTypeNames.FONTTYPE_WARNING)

                Exit Sub

            End If
            
            Call mFacciones.Faction_AddUser(UserIndex, r_Caos)
            Call Guilds_CheckAlineation(UserIndex, a_Legion)

        End If
        
    End With

    '<EhFooter>
    Exit Sub

HandleEnlist_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleEnlist " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleReward(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleReward_Err

    '</EhHeader>
    
    With UserList(UserIndex)
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If

        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub

        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
    
        Call mFacciones.Faction_CheckRangeUser(UserIndex)
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleReward_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleReward " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleFianza(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleFianza_Err

    '</EhHeader>
    
    '***************************************************
    'Author: Matías Ezequiel
    'Last Modification: 16/03/2016 by DS
    'Sistema de fianzas TDS.
    '***************************************************
    Dim Fianza      As Long

    Dim ValueFianza As Long
        
    With UserList(UserIndex)
        
        Fianza = Reader.ReadInt
        ValueFianza = Fianza * 5
            
        If Fianza <= 0 Or Fianza > MAXORO Then Exit Sub
        
        If MapInfo(.Pos.Map).Pk Then
            Call WriteConsoleMsg(UserIndex, "Debes estar en zona segura para utilizar este comando.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If .flags.Muerto Then
            Call WriteConsoleMsg(UserIndex, "Estás muerto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If

        If (ValueFianza) > .Stats.Gld Then
            Call WriteConsoleMsg(UserIndex, "Para pagar esa fianza necesitas pagar impuestos. El precio total es: " & (ValueFianza) & " Monedas de Oro. ¡Agradece que no son Eldhires gusano!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        Dim EraCriminal As Boolean
        
        EraCriminal = Escriminal(UserIndex)
        .Reputacion.NobleRep = .Reputacion.NobleRep + Fianza
        .Stats.Gld = .Stats.Gld - (ValueFianza)

        Call WriteConsoleMsg(UserIndex, "Has ganado " & Fianza & " puntos de noble.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Se te han descontado " & ValueFianza & " Monedas de Oro.", FontTypeNames.FONTTYPE_INFO)
        Call WriteUpdateGold(UserIndex)
        
        If EraCriminal And Not Escriminal(UserIndex) Then
            Call RefreshCharStatus(UserIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleFianza_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleFianza " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleHome(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleHome_Err

    '</EhHeader>

    '***************************************************
    'Author: Budi
    'Creation Date: 06/01/2010
    'Last Modification: 05/06/10
    'Pato - 05/06/10: Add the Ucase$ to prevent problems.
    '***************************************************
    With UserList(UserIndex)
            
        ' @ El personaje se asocia a una nueva CIUDAD.
        If .flags.TargetNPC > 0 Then
            If Npclist(.flags.TargetNPC).Ciudad > 0 Then
                Call setHome(UserIndex, Npclist(.flags.TargetNPC).Ciudad, .flags.TargetNPC)
                Exit Sub

            End If

        End If
            
        If .flags.Muerto = 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes usar el comando si estás vivo.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If Not MapInfo(.Pos.Map).Pk Then
            Call WriteConsoleMsg(UserIndex, "Ya te encuentras en Zona Segura", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
            
        If .flags.Traveling = 1 Then
            Call EndTravel(UserIndex, True)
            Exit Sub

        End If
                
        Dim RequiredGld As Long
            
        If Not EsNewbie(UserIndex) Then
            If .Stats.Elv < 20 Then
                RequiredGld = 10 * .Stats.Elv
            ElseIf .Stats.Elv < 35 Then
                RequiredGld = 50 * .Stats.Elv
            Else
                RequiredGld = 150 * .Stats.Elv

            End If

        End If
            
        Select Case .Account.Premium

            Case 0
                
            Case 1, 2
                RequiredGld = RequiredGld / 2

            Case 3
                RequiredGld = 0

        End Select
            
        If .flags.SlotEvent > 0 Then
            WriteConsoleMsg UserIndex, "No puedes usar la restauración si estás en un evento.", FontTypeNames.FONTTYPE_INFO

            Exit Sub

        End If
              
        If .flags.SlotReto > 0 Or .flags.SlotFast > 0 Then
            WriteConsoleMsg UserIndex, "No puede susar este comando si estás en reto.", FontTypeNames.FONTTYPE_INFO

            Exit Sub

        End If
        
        If .Counters.Pena > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes usar la restauración si estás en la carcel.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
            
        If .Stats.Gld < RequiredGld Then
            Call WriteConsoleMsg(UserIndex, "El viaje requiere que dispongas de " & RequiredGld & " Monedas de Oro.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
            
        .Stats.Gld = .Stats.Gld - RequiredGld
        Call WriteUpdateGold(UserIndex)
            
        Call goHome(UserIndex)

    End With

    '<EhFooter>
    Exit Sub

HandleHome_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleHome " & "at line " & Erl

    '</EhFooter>
End Sub

''
' Prepares the "UpdateControlPotas" message and returns it.
'

Public Function PrepareMessageUpdateControlPotas(ByVal charindex As Integer, _
                                                 ByVal MinHp As Integer, _
                                                 ByVal MaxHp As Integer, _
                                                 ByVal MinMan As Integer, _
                                                 ByVal MaxMan As Integer) As String

    '<EhHeader>
    On Error GoTo PrepareMessageUpdateControlPotas_Err

    '</EhHeader>

    '***************************************************
    'Author
    'Last Modification:
    '
    '***************************************************
    Call Writer.WriteInt(ServerPacketID.UpdateControlPotas)
        
    Call Writer.WriteInt(charindex)
    Call Writer.WriteInt(MinHp)
    Call Writer.WriteInt(MaxHp)
    Call Writer.WriteInt(MinMan)
    Call Writer.WriteInt(MaxMan)

    '<EhFooter>
    Exit Function

PrepareMessageUpdateControlPotas_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageUpdateControlPotas " & "at line " & Erl
        
    '</EhFooter>
End Function

'
' Handles the "SendReply" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSendReply(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleSendReply_Err

    '</EhHeader>

    '***************************************************
    'Author:
    'Last Modification:
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim Fight   As tFight

        Dim Users() As String

        Dim Temp    As String

        Dim A       As Long

        Temp = Reader.ReadString8
        Fight.Tipo = Reader.ReadInt8
        Fight.Gld = Reader.ReadInt32
        Fight.Time = (Reader.ReadInt8 * 60)
        Fight.RoundsLimit = Reader.ReadInt8
        Fight.Terreno = Reader.ReadInt8

        For A = LBound(Fight.config) To UBound(Fight.config)
            Fight.config(A) = Reader.ReadInt8
        Next A
    
        Users = Split(Temp, "-")
        
        ReDim Fight.User(LBound(Users) To UBound(Users)) As tFightUser
        
        For A = LBound(Users) To UBound(Users)
            Fight.User(A).Name = Users(A)
        Next A
        
        Call mRetos.SendFight(UserIndex, Fight)
              
    End With
    
    '<EhFooter>
    Exit Sub

HandleSendReply_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleSendReply " & "at line " & Erl
        
    '</EhFooter>
End Sub

'
' Handles the "AcceptReply" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptReply(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleAcceptReply_Err

    '</EhHeader>

    '***************************************************
    'Author:
    'Last Modification:
    '
    '***************************************************

    With UserList(UserIndex)
        
        Dim UserName As String
        
        UserName = Reader.ReadString8
                      
        Call mRetos.AcceptFight(UserIndex, UserName)
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleAcceptReply_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleAcceptReply " & "at line " & Erl
        
    '</EhFooter>
End Sub

'
' Handles the "AbandonateReply" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAbandonateReply(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleAbandonateReply_Err

    '</EhHeader>

    '***************************************************
    'Author:
    'Last Modification:
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.SlotReto > 0 Then
            Call mRetos.UserdieFight(UserIndex, 0, True)

        End If
    
    End With

    '<EhFooter>
    Exit Sub

HandleAbandonateReply_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleAbandonateReply " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleEvents_CreateNew(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleEvents_CreateNew_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        Dim Temp As tEvents
        
        Dim A    As Integer
        
        Temp.Modality = Reader.ReadInt
        Temp.Name = Reader.ReadString8
        Temp.QuotasMin = Reader.ReadInt
        Temp.QuotasMax = Reader.ReadInt
        Temp.LvlMin = Reader.ReadInt
        Temp.LvlMax = Reader.ReadInt
        Temp.InscriptionGld = Reader.ReadInt
        Temp.InscriptionEldhir = Reader.ReadInt
        Temp.TimeInscription = Reader.ReadInt
        Temp.TimeCancel = Reader.ReadInt
        Temp.TeamCant = Reader.ReadInt
                      
        Temp.LimitRed = Reader.ReadInt
        Temp.PrizeGld = Reader.ReadInt
        Temp.PrizeEldhir = Reader.ReadInt
        Temp.PrizeObj.ObjIndex = Reader.ReadInt
        Temp.PrizeObj.Amount = Reader.ReadInt
        
        ReDim Temp.AllowedClasses(1 To NUMCLASES) As Byte
        ReDim Temp.AllowedFaction(1 To 4) As Byte
        
        For A = 1 To 4
            Temp.AllowedFaction(A) = Reader.ReadInt()
        Next A

        For A = 1 To NUMCLASES
            Temp.AllowedClasses(A) = Reader.ReadInt()
        Next A
                      
        Temp.ChangeClass = Reader.ReadInt()
        Temp.ChangeRaze = Reader.ReadInt()
        
        For A = 1 To MAX_EVENTS_CONFIG
            Temp.config(A) = Reader.ReadInt8()
        Next A
        
        Temp.LimitRound = Reader.ReadInt8()
        Temp.LimitRoundFinal = Reader.ReadInt8()
        Temp.GanaSigue = Reader.ReadInt8()
        Temp.ArenasLimit = Reader.ReadInt8()
        Temp.ArenasMin = Reader.ReadInt8()
        Temp.ArenasMax = Reader.ReadInt()
        Temp.ChangeLevel = Reader.ReadInt8()
        Temp.Prob = Reader.ReadInt8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then

            Dim CanEvent As Byte: CanEvent = NewEvent(Temp)
            
            If CanEvent <> 0 Then
                Events(CanEvent).Enabled = True

            Else
                Call WriteConsoleMsg(UserIndex, "No hay más cupos para eventos o bien ya existe un evento con esa modalidad", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleEvents_CreateNew_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleEvents_CreateNew " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleEvents_Close(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleEvents_Close_Err

    '</EhHeader>
    
    With UserList(UserIndex)
        
        Dim Slot As Byte
        
        Slot = Reader.ReadInt
        
        If Slot <= 0 Or Slot > MAX_EVENT_SIMULTANEO Then Exit Sub
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            
            Call EventosDS.CloseEvent(Slot, , True)

        End If
        
    End With

    '<EhFooter>
    Exit Sub

HandleEvents_Close_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleEvents_Close " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandlePro_Seguimiento(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandlePro_Seguimiento_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim Seguir   As Boolean
        
        UserName = Reader.ReadString8
        Seguir = Reader.ReadBool
        
        tUser = NameIndex(UserName)
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            If tUser > 0 Then
                If Seguir Then
                    UserList(tUser).flags.GmSeguidor = UserIndex
                    Call WriteConsoleMsg(UserIndex, "Has comenzado el seguimiento al usuario " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_INFOGREEN)
                Else
                    UserList(tUser).flags.GmSeguidor = 0
                    Call WriteConsoleMsg(UserIndex, "Has reiniciado el seguimiento al usuario " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_INFOGREEN)

                End If

            Else
                Call WriteConsoleMsg(UserIndex, "El personaje está offline", FontTypeNames.FONTTYPE_INFORED)

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandlePro_Seguimiento_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandlePro_Seguimiento " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function PrepareMessageUpdateGroupIndex(ByVal charindex As Integer, _
                                               ByVal GroupIndex As Byte) As String

    '<EhHeader>
    On Error GoTo PrepareMessageUpdateGroupIndex_Err

    '</EhHeader>

    Call Writer.WriteInt(ServerPacketID.UpdateGroupIndex)
    Call Writer.WriteInt(charindex)
    Call Writer.WriteInt(GroupIndex)

    '<EhFooter>
    Exit Function

PrepareMessageUpdateGroupIndex_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageUpdateGroupIndex " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub WriteUpdateInfoIntervals(ByVal UserIndex As Integer, _
                                    ByVal Tipo As Byte, _
                                    ByVal Value As Long, _
                                    ByVal MenuCliente As Byte)

    '<EhHeader>
    On Error GoTo WriteUpdateInfoIntervals_Err

    '</EhHeader>

    Call Writer.WriteInt(ServerPacketID.UpdateInfoIntervals)
        
    Call Writer.WriteInt(Tipo)
    Call Writer.WriteInt(Value)
    Call Writer.WriteInt(MenuCliente)
        
    Call SendData(ToOne, UserIndex, vbNullString)
    
    '<EhFooter>
    Exit Sub

WriteUpdateInfoIntervals_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateInfoIntervals " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleEvent_Participe(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleEvent_Participe_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        Dim Modality As String

        Dim Slot     As Byte

        Dim ErrorMsg As String
        
        Modality = Reader.ReadString8
        Slot = Events_SearchSlotEvent(Modality): If Slot = 0 Then Exit Sub
    
        If Not Event_CheckInscriptions_User(UserIndex, Slot, ErrorMsg) Then
            Call WriteConsoleMsg(UserIndex, ErrorMsg, FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If
            
        If Event_CheckExistUser(UserIndex, Slot) > 0 Then
            Call WriteConsoleMsg(UserIndex, "Ya estás inscripto en esta partida. Espera a que se complete y participarás automáticamente.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If
        
        If .GroupIndex > 0 Then
            Call Events_Group_Set(.GroupIndex, Slot)
        Else
            Call Event_SetNewUser(UserIndex, Slot)

        End If
            
    End With
    
    '<EhFooter>
    Exit Sub

HandleEvent_Participe_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleEvent_Participe " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleUpdateInactive(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleUpdateInactive_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        .Counters.TimeInactive = 0
    
    End With

    '<EhFooter>
    Exit Sub

HandleUpdateInactive_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleUpdateInactive " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleRetos_RewardObj(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRetos_RewardObj_Err

    '</EhHeader>
    
    With UserList(UserIndex)
         
        If .flags.ClainObject = 0 Then Exit Sub
        If MapInfo(.Pos.Map).Pk Then Exit Sub
        
        Call Retos_ReclameObj(UserIndex)
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleRetos_RewardObj_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRetos_RewardObj " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleEvents_KickUser(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleEvents_KickUser_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = Reader.ReadString8
        
        tUser = NameIndex(UserName)
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            If tUser > 0 Then
                Call EventosDS.AbandonateEvent(tUser)
                Call WriteConsoleMsg(UserIndex, "Has kickeado del evento al personaje " & UserList(tUser).Name & ". ¡No abuses de tu poder!", FontTypeNames.FONTTYPE_INFOGREEN)
                Call Logs_User(.Name, eLog.eGm, eNone, "El GM " & .Name & " ha kickeado del evento al personaje " & UserList(tUser).Name & ".")
            Else
                Call WriteConsoleMsg(UserIndex, "El personaje está offline. Actualiza la lista de eventos.", FontTypeNames.FONTTYPE_INFORED)

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleEvents_KickUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleEvents_KickUser " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleGuilds_Required(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleGuilds_Required_Err

    '</EhHeader>
    
    With UserList(UserIndex)
        
        Dim Value As Integer

        Value = Reader.ReadInt
        
        If Value = 0 Then
            Call WriteGuild_List(UserIndex, Guilds_PrepareList)
        ElseIf Value > 0 And Value < MAX_GUILDS Then
            Call WriteGuild_Info(UserIndex, Value, GuildsInfo(Value), GuildsInfo(Value).Members)
        Else
            
            Select Case Value
            
                Case 1000
                    Call Guilds_PrepareInfoUsers(UserIndex)

            End Select
            
        End If
    
    End With

    '<EhFooter>
    Exit Sub

HandleGuilds_Required_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGuilds_Required " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteGuild_List(ByVal UserIndex As Integer, ByRef GuildList() As String)

    '<EhHeader>
    On Error GoTo WriteGuild_List_Err

    '</EhHeader>

    Dim Tmp As String

    Dim A   As Long
    
    Call Writer.WriteInt(ServerPacketID.Guild_List)
    Call Writer.WriteBool(UserList(UserIndex).GuildRange = rLeader Or UserList(UserIndex).GuildRange = rFound)
        
    For A = LBound(GuildList()) To UBound(GuildList())
        Tmp = Tmp & GuildList(A) & SEPARATOR
    Next A
        
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
    Call Writer.WriteString8(Tmp)
    
    For A = 1 To MAX_GUILDS
        Call Writer.WriteInt8(GuildsInfo(A).Alineation)
        Call Writer.WriteInt8(GuildsInfo(A).NumMembers)
        Call Writer.WriteInt8(GuildsInfo(A).MaxMembers)
        Call Writer.WriteInt8(GuildsInfo(A).Lvl)
        Call Writer.WriteInt32(GuildsInfo(A).Exp)
        Call Writer.WriteInt32(GuildsInfo(A).Elu)
    Next A
    
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteGuild_List_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteGuild_List " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleGuilds_Found(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleGuilds_Found_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        Dim Name                        As String, Temp As String

        Dim Alineation                  As eGuildAlineation

        Dim Codex(1 To MAX_GUILD_CODEX) As String

        Dim A                           As Long
        
        Name = Reader.ReadString8
        Alineation = Reader.ReadInt8
        
        Call mGuilds.Guilds_New(UserIndex, Name, Alineation, Codex)
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleGuilds_Found_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGuilds_Found " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleGuilds_Invitation(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleGuilds_Invitation_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        Dim UserName As String

        Dim Tipo     As Byte
        
        UserName = Reader.ReadString8
        Tipo = Reader.ReadInt
        
        Select Case Tipo
        
            Case 0  ' El lider enviá solicitud a un miembro.
                Call Guilds_SendInvitation(UserIndex, UserName)

            Case 1 ' El personaje acepta la solicitud del Lider
                Call Guilds_AcceptInvitation(UserIndex, UserName)

        End Select
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleGuilds_Invitation_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGuilds_Invitation " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteGuild_Info(ByVal UserIndex As Integer, _
                           ByVal GuildIndex As Integer, _
                           ByRef GuildInfo As tGuild, _
                           ByRef MemberInfo() As tGuildMember)

    '<EhHeader>
    On Error GoTo WriteGuild_Info_Err

    '</EhHeader>

    Dim A As Long
    
    Call Writer.WriteInt(ServerPacketID.Guild_Info)
    Call Writer.WriteInt16(GuildIndex)
    
    Call Writer.WriteString8(GuildInfo.Name)
    Call Writer.WriteInt8(GuildInfo.Alineation)
    
    For A = 1 To MAX_GUILD_MEMBER
        Call Writer.WriteString8(MemberInfo(A).Name)
        Call Writer.WriteInt8(MemberInfo(A).Range)
            
        Call Writer.WriteInt16(MemberInfo(A).Char.Body)
        Call Writer.WriteInt16(MemberInfo(A).Char.Head)
        Call Writer.WriteInt16(MemberInfo(A).Char.Helm)
        Call Writer.WriteInt16(MemberInfo(A).Char.Shield)
        Call Writer.WriteInt16(MemberInfo(A).Char.Weapon)
            
    Next A
        
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteGuild_Info_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteGuild_Info " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleGuilds_Online(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleGuilds_Online_Err

    '</EhHeader>
    
    With UserList(UserIndex)
        
        If .GuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, "No perteneces a ningún clan.", FontTypeNames.FONTTYPE_INFORED)
        Else
            Call Guilds_PrepareOnline(UserIndex, .GuildIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleGuilds_Online_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGuilds_Online " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteGuild_InfoUsers(ByVal UserIndex As Integer, _
                                ByVal GuildIndex As Integer, _
                                ByRef MemberInfo() As tGuildMember)

    '<EhHeader>
    On Error GoTo WriteGuild_InfoUsers_Err

    '</EhHeader>

    Dim A As Long
    
    Call Writer.WriteInt(ServerPacketID.Guild_InfoUsers)

    Call Writer.WriteInt16(GuildIndex)
    
    For A = 1 To MAX_GUILD_MEMBER
        Call Writer.WriteString8(MemberInfo(A).Name)
        Call Writer.WriteInt(MemberInfo(A).Range)
            
        Call Writer.WriteInt(MemberInfo(A).Char.Elv)
        Call Writer.WriteInt(MemberInfo(A).Char.Class)
        Call Writer.WriteInt(MemberInfo(A).Char.Raze)
            
        Call Writer.WriteInt(MemberInfo(A).Char.Body)
        Call Writer.WriteInt(MemberInfo(A).Char.Head)
        Call Writer.WriteInt(MemberInfo(A).Char.Helm)
        Call Writer.WriteInt(MemberInfo(A).Char.Shield)
        Call Writer.WriteInt(MemberInfo(A).Char.Weapon)
        Call Writer.WriteInt(MemberInfo(A).Char.Points)
            
    Next A
        
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteGuild_InfoUsers_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteGuild_InfoUsers " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleGuilds_Kick(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleGuilds_Kick_Err

    '</EhHeader>
    
    Dim UserName As String

    With UserList(UserIndex)
        
        UserName = Reader.ReadString8

        Call Guilds_KickUser(UserIndex, UCase$(UserName))

    End With
    
    '<EhFooter>
    Exit Sub

HandleGuilds_Kick_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGuilds_Kick " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleGuilds_Abandonate(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleGuilds_Abandonate_Err

    '</EhHeader>
    
    With UserList(UserIndex)
        
        If .GuildIndex > 0 Then
            Call Guilds_KickMe(UserIndex)
        Else
            Call WriteConsoleMsg(UserIndex, "No perteneces a ningún clan.", FontTypeNames.FONTTYPE_INFORED)

        End If

    End With
    
    '<EhFooter>
    Exit Sub

HandleGuilds_Abandonate_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGuilds_Abandonate " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteFight_PanelAccept(ByVal UserIndex As Integer, _
                                  ByVal UserName As String, _
                                  ByVal TextUsers As String, _
                                  ByRef RetoTemp As tFight)

    '<EhHeader>
    On Error GoTo WriteFight_PanelAccept_Err

    '</EhHeader>

    Dim A    As Long

    Dim Str  As String

    Dim Temp As Byte
    
    Call Writer.WriteInt(ServerPacketID.Fight_PanelAccept)
        
    Call Writer.WriteString8(UserName)
    Call Writer.WriteString8(TextUsers)
    Call Writer.WriteInt32(RetoTemp.Gld)
    Call Writer.WriteInt8(RetoTemp.RoundsLimit)
    Call Writer.WriteInt8(RetoTemp.Terreno)
    
    Temp = Int(RetoTemp.Time / 60)
    Call Writer.WriteInt8(Temp)
    
    For A = 1 To MAX_RETOS_CONFIG
        Call Writer.WriteInt8(RetoTemp.config(A))
    Next A
        
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteFight_PanelAccept_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteFight_PanelAccept " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleFight_CancelInvitation(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleFight_CancelInvitation_Err

    '</EhHeader>

    With UserList(UserIndex)

        .Counters.FightInvitation = 0

    End With

    '<EhFooter>
    Exit Sub

HandleFight_CancelInvitation_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleFight_CancelInvitation " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleGuilds_Talk(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleGuilds_Talk_Err

    '</EhHeader>

    With UserList(UserIndex)

        Dim chat      As String

        Dim IsSupport As Boolean
            
        Dim CanTalk   As Boolean
              
        chat = Reader.ReadString8()
        IsSupport = Reader.ReadBool()
              
        If LenB(chat) <> 0 Then
                  
            CanTalk = True

            If .flags.SlotEvent > 0 Then
                If Events(.flags.SlotEvent).Modality = eModalityEvent.DeathMatch Then
                    CanTalk = False

                End If

            End If
                  
            If CanTalk Then
                If .GuildIndex > 0 Then
                    If IsSupport Then
                        If GuildsInfo(.GuildIndex).Lvl >= 3 Then
                            Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageConsoleMsg("[AYUDA] " & .Name & "> " & chat, FontTypeNames.FONTTYPE_INFORED))

                        End If

                    Else
                        Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageConsoleMsg("[CLANES]" & .Name & "> " & chat, FontTypeNames.FONTTYPE_GUILDMSG))

                    End If
                    
                End If

            End If

        End If

    End With
    
    '<EhFooter>
    Exit Sub

HandleGuilds_Talk_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGuilds_Talk " & "at line " & Erl

    '</EhFooter>
End Sub

Sub LoadPictureConMatrizDeBytes(ByRef Arrai() As Byte)

    '<EhHeader>
    On Error GoTo LoadPictureConMatrizDeBytes_Err

    '</EhHeader>

    Dim FilePath  As String: FilePath = App.Path & "PRUEBA.BMP"
  
    Dim FileIndex As Integer

    Dim A         As Long
    
    FileIndex = FreeFile
  
    Open FilePath For Output As FileIndex

    For A = LBound(Arrai) To UBound(Arrai)
        Print #FileIndex, Arrai(A)
    Next A

    Close FileIndex
  
    '<EhFooter>
    Exit Sub

LoadPictureConMatrizDeBytes_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.LoadPictureConMatrizDeBytes " & "at line " & Erl
        
    '</EhFooter>
End Sub
 
Private Sub HandleSendPic(ByVal UserIndex As Integer)
   
End Sub

Public Sub HandleLoginAccount(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleLoginAccount_Err

    '</EhHeader>
        
    Dim Version   As String, Email As String, Passwd As String

    Dim Time      As Long

    Dim SERIAL(7) As String

    Dim Temp      As tAccountSecurity

    'Dim Key_Encrypt As String: Key_Encrypt = mEncrypt_B.XOREncryption("ILMWNlOOvtUkOjo6bu")
    'Dim Key_Decrypt As String: Key_Decrypt = "ILMWNlOOvtUkOjo6bu"
    
    Version = CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt())
    
    'Email = mEncrypt_A.AesDecryptString(Reader.ReadString8, mEncrypt_B.XOR_CHARACTER)
    'Passwd = mEncrypt_A.AesDecryptString(Reader.ReadString8, mEncrypt_B.XOR_CHARACTER)
    Email = Reader.ReadString8
    Passwd = Reader.ReadString8
    Temp.SERIAL_BIOS = Reader.ReadString8 ' Serial_Bios
    Temp.SERIAL_DISK = Reader.ReadString8 ' Serial_DISK
    Temp.SERIAL_MAC = Reader.ReadString8 ' Serial_MAC
    Temp.SERIAL_MOTHERBOARD = Reader.ReadString8 ' Serial_MOTHERBOARD
    Temp.SERIAL_PROCESSOR = Reader.ReadString8 ' Serial_PROCESSOR
    Temp.SYSTEM_DATA = Reader.ReadString8 ' System Data
    Temp.IP_Local = Reader.ReadString8 ' IP LOCAL
    Temp.IP_Public = Reader.ReadString8 ' IP PUBLICA
    
    Time = GetTime
        
    'If SERIAL(0) <> vbNullString Then Temp.SERIAL_BIOS = mEncrypt_A.AesDecryptString(SERIAL(0), Key_Decrypt)
    'If SERIAL(1) <> vbNullString Then Temp.SERIAL_DISK = mEncrypt_A.AesDecryptString(SERIAL(1), Key_Encrypt)
    'If SERIAL(2) <> vbNullString Then Temp.SERIAL_MAC = mEncrypt_A.AesDecryptString(SERIAL(2), Key_Decrypt)
    'If SERIAL(3) <> vbNullString Then Temp.SERIAL_MOTHERBOARD = mEncrypt_A.AesDecryptString(SERIAL(3), Key_Encrypt)
    'If SERIAL(4) <> vbNullString Then Temp.SERIAL_PROCESSOR = mEncrypt_A.AesDecryptString(SERIAL(4), Key_Decrypt)
    'If SERIAL(5) <> vbNullString Then Temp.SYSTEM_DATA = mEncrypt_A.AesDecryptString(SERIAL(5), Key_Encrypt)
    'If SERIAL(6) <> vbNullString Then Temp.IpAddress_Local = mEncrypt_A.AesDecryptString(SERIAL(6), Key_Encrypt)
    'If SERIAL(7) <> vbNullString Then Temp.IpAddress_Public = mEncrypt_A.AesDecryptString(SERIAL(7), Key_Decrypt)
    
    Email = LCase$(Email)
        
    Dim Testing    As Boolean

    Const TIMER_MS As Byte = 250

    If (Time - TIMER_MS) <= TIMER_MS Then Exit Sub
        
    #If Testeo = 1 Then
        Testing = True
    #End If
    
    SLOT_TERMINAL_ARCHIVE = 1 'LwK
        
    If SLOT_TERMINAL_ARCHIVE = 0 And Not Testing Then
        Call Protocol.Kick(UserIndex, "Servidor en mantenimiento. Consulta otros servidores para disfrutar y pasar el rato.")
        
    Else

        If ServerSoloGMs > 0 Then
            If Not Email_Is_Testing_Pro(Email) Then
                Call Protocol.Kick(UserIndex, "Servidor en mantenimiento. Consulta otros servidores para disfrutar y pasar el rato.")
                Exit Sub
        
            End If

        End If
        
        frmMain.lstDebug.AddItem "Iniciando " & Email
    
        If Not VersionOK(Version) Then
            Call Protocol.Kick(UserIndex, "Se ha detectado una versión obsoleta. Compruebe actualizaciones.")
        Else

            If mAccount.LoginAccount(UserIndex, LCase$(Email), Passwd) Then
                UserList(UserIndex).Account.Sec = Temp
                UserList(UserIndex).Account.Sec.IP_Address = UserList(UserIndex).IpAddress
                'UserList(UserIndex).IpAddress = UserList(UserIndex).Account.Sec.IP_Public
                Call Logs_Account_SettingData(UserIndex, "LOGIN", LCase$(Email))
                Call WriteRequestID(LCase$(Email))

            End If

        End If

    End If
        
    UserList(UserIndex).LastRequestLogin = Time

    '<EhFooter>
    Exit Sub

HandleLoginAccount_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleLoginAccount " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Actualiza Slot de Mercado y/o Monedas de Oro de la Cuenta, como así también PREMIUM.
Public Sub WriteAccountInfo(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteAccountInfo_Err

    '</EhHeader>
    Call Writer.WriteInt(ServerPacketID.AccountInfo)
    
    Call Writer.WriteInt32(UserList(UserIndex).Account.Gld)
    Call Writer.WriteInt32(UserList(UserIndex).Account.Eldhir)
    Call Writer.WriteInt8(UserList(UserIndex).Account.Premium)
    Call Writer.WriteInt16(UserList(UserIndex).Account.MercaderSlot)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Points)
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteAccountInfo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteAccountInfo " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteLoggedAccount(ByVal UserIndex As Integer, ByRef Temp() As tAccountChar)

    '<EhHeader>
    On Error GoTo WriteLoggedAccount_Err

    '</EhHeader>

    Call Writer.WriteInt(ServerPacketID.loggedaccount)
    
    Dim A As Long
    
    Call Writer.WriteInt32(UserList(UserIndex).Account.Gld)
    Call Writer.WriteInt32(UserList(UserIndex).Account.Eldhir)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Points)
    
    Call Writer.WriteInt8(UserList(UserIndex).Account.Premium)
    Call Writer.WriteInt16(UserList(UserIndex).Account.MercaderSlot)
    
    Call Writer.WriteInt8(UserList(UserIndex).Account.CharsAmount)
    
    For A = 1 To ACCOUNT_MAX_CHARS
        Call Writer.WriteInt8(A)
        Call Writer.WriteString8(Temp(A).Name)
        Call Writer.WriteInt8(Temp(A).Blocked)
        Call Writer.WriteString8(Temp(A).Guild)
            
        Call Writer.WriteInt16(Temp(A).Body)
        Call Writer.WriteInt16(Temp(A).Head)
        Call Writer.WriteInt16(Temp(A).Weapon)
        Call Writer.WriteInt16(Temp(A).Shield)
        Call Writer.WriteInt16(Temp(A).Helm)
            
        Call Writer.WriteInt8(Temp(A).Ban)
            
        Call Writer.WriteInt8(Temp(A).Class)
        Call Writer.WriteInt8(Temp(A).Raze)
        Call Writer.WriteInt8(Temp(A).Elv)
            
        Call Writer.WriteInt16(Temp(A).Map)
        Call Writer.WriteInt8(Temp(A).posX)
        Call Writer.WriteInt8(Temp(A).posY)
            
        Call Writer.WriteInt8(Temp(A).Faction)
        Call Writer.WriteInt8(Temp(A).FactionRange)
    Next A
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteLoggedAccount_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteLoggedAccount " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteLoggedAccount_DataChar(ByVal UserIndex As Integer, _
                                       ByVal Slot As Byte, _
                                       ByRef DataChar As tAccountChar)

    '<EhHeader>
    On Error GoTo WriteLoggedAccount_DataChar_Err

    '</EhHeader>

    Call Writer.WriteInt(ServerPacketID.LoggedAccount_DataChar)

    Call Writer.WriteInt8(Slot)
    Call Writer.WriteString8(DataChar.Name)
    Call Writer.WriteString8(DataChar.Guild)
        
    Call Writer.WriteInt16(DataChar.Body)
    Call Writer.WriteInt16(DataChar.Head)
    Call Writer.WriteInt16(DataChar.Weapon)
    Call Writer.WriteInt16(DataChar.Shield)
    Call Writer.WriteInt16(DataChar.Helm)
        
    Call Writer.WriteInt8(DataChar.Ban)
        
    Call Writer.WriteInt8(DataChar.Class)
    Call Writer.WriteInt8(DataChar.Raze)
    Call Writer.WriteInt8(DataChar.Elv)
        
    Call Writer.WriteInt16(DataChar.Map)
    Call Writer.WriteInt8(DataChar.posX)
    Call Writer.WriteInt8(DataChar.posY)
        
    Call Writer.WriteInt8(DataChar.Faction)
    Call Writer.WriteInt8(DataChar.FactionRange)
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteLoggedAccount_DataChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteLoggedAccount_DataChar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteConnectedMessage(ByVal UserIndex As Integer, ByVal ServerSelected As Byte)

    '<EhHeader>
    On Error GoTo WriteConnectedMessage_Err

    '</EhHeader>

    Call Writer.WriteInt(ServerPacketID.Connected)
    Call Writer.WriteInt8(ServerSelected)
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteConnectedMessage_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteConnectedMessage " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteLoggedRemoveChar(ByVal UserIndex As Integer, ByVal SlotUserName As Byte)

    '<EhHeader>
    On Error GoTo WriteLoggedRemoveChar_Err

    '</EhHeader>

    Dim A As Long
    
    Call Writer.WriteInt(ServerPacketID.LoggedRemoveChar)
    Call Writer.WriteInt(SlotUserName)
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteLoggedRemoveChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteLoggedRemoveChar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleLoginChar(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleLoginChar_Err

    '</EhHeader>

    Dim UserName As String

    Dim Version  As String

    Dim Key      As String
    
    Dim Slot     As Byte
    
    Version = CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt())
    UserName = Reader.ReadString8()
    Key = Reader.ReadString8
    Slot = Reader.ReadInt8
    
    If Not VersionOK(Version) Then
        Call Protocol.Kick(UserIndex, "Se ha detectado una versión obsoleta. Compruebe actualizaciones.")
    ElseIf PuedeConectarPersonajes = 0 Then
        Call Protocol.Kick(UserIndex, "No está permitido el ingreso de personajes al juego.")
    ElseIf CheckUserLogged(UCase$(UserName)) Then
        Call WriteErrorMsg(UserIndex, "El personaje se encuentra online.")
    Else
        Call mAccount.LoginAccount_Char(UserIndex, UserName, Key, Slot, False)

    End If

    '<EhFooter>
    Exit Sub

HandleLoginChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleLoginChar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleDisconnectForced(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleDisconnectForced_Err

    '</EhHeader>

    Dim Account As String

    Dim Key     As String

    Dim Version As String

    Version = CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt())
    Account = Reader.ReadString8()
    Key = Reader.ReadString8()
    
    If Not VersionOK(Version) Then
        Call Protocol.Kick(UserIndex, "Se ha detectado una versión obsoleta. Compruebe actualizaciones.")
    Else
        Call mAccount.DisconnectForced(UserIndex, LCase$(Account), Key)

    End If

    '<EhFooter>
    Exit Sub

HandleDisconnectForced_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleDisconnectForced " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleLoginCharNew(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleLoginCharNew_Err

    '</EhHeader>

    Dim Key       As String

    Dim Version   As String

    Dim UserName  As String

    Dim UserClase As Byte

    Dim UserRaza  As Byte

    Dim UserSexo  As Byte
    
    Dim UserHead  As Integer
    
    Dim Slot      As Byte
    
    Version = CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt())
    UserName = Reader.ReadString8()
    UserClase = Reader.ReadInt8()
    UserRaza = Reader.ReadInt8()
    UserSexo = Reader.ReadInt8()
    UserHead = Reader.ReadInt16()
    
    If Not VersionOK(Version) Then
        Call Protocol.Kick(UserIndex, "Se ha detectado una versión obsoleta. Compruebe actualizaciones.")
    Else
        Call mAccount.LoginAccount_CharNew(UserIndex, UserName, UserClase, UserRaza, UserSexo, UserHead)

    End If

    '<EhFooter>
    Exit Sub

HandleLoginCharNew_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleLoginCharNew " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleLoginName(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleLoginName_Err

    '</EhHeader>

    Dim Version  As String

    Dim UserName As String
    
    Version = CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt())
    UserName = Reader.ReadString8()
    
    #If Classic = 1 Then
        Exit Sub
    #End If
    
    If Not VersionOK(Version) Then
        Call Protocol.Kick(UserIndex, "Se ha detectado una versión obsoleta. Compruebe actualizaciones.")
    Else
        Call mAccount.LoginAccount_ChangeAlias(UserIndex, UserName)

    End If

    '<EhFooter>
    Exit Sub

HandleLoginName_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleLoginName " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleLoginRemove(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleLoginRemove_Err

    '</EhHeader>

    Dim Key     As String

    Dim Version As String
    
    Dim Slot    As Byte
    
    Version = CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt()) & "." & CStr(Reader.ReadInt())
    Key = Reader.ReadString8
    Slot = Reader.ReadInt8
    
    If Not VersionOK(Version) Then
        Call Protocol.Kick(UserIndex, "Se ha detectado una versión obsoleta. Compruebe actualizaciones.")
    Else
        Call mAccount.LoginAccount_Remove(UserIndex, Key, Slot)

    End If

    '<EhFooter>
    Exit Sub

HandleLoginRemove_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleLoginRemove " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleMercader_New(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleMercader_New_Err

    '</EhHeader>

    Dim Key          As String, Passwd As String
    
    Dim Chars()      As Byte
    
    Dim Mercader     As tMercaderChar
    
    Dim Gld          As Long, Dsp As Long
    
    Dim A            As Long, Desc As String
    
    Dim SaleCost     As Long
    
    Dim Blocked      As Byte
    
    Dim CantChars    As Byte
        
    Dim SlotMercader As Integer 'Slot al que queremos ofrecer

    Passwd = Reader.ReadString8
    Key = Reader.ReadString8
          
    SlotMercader = Reader.ReadInt16
          
    Gld = Reader.ReadInt32
    Dsp = Reader.ReadInt32
    Desc = Reader.ReadString8
        
    SaleCost = Reader.ReadInt32
    Blocked = Reader.ReadInt8
    
    Call Reader.ReadSafeArrayInt8(Chars)
            
    If Not StrComp(Key, UserList(UserIndex).Account.Key) = 0 Then
        Call WriteErrorMsg(UserIndex, "Has escrito una clave de seguridad erronea.")
        Exit Sub

    End If
        
    If Not StrComp(Passwd, UserList(UserIndex).Account.Passwd) = 0 Then
        Call WriteErrorMsg(UserIndex, "Has escrito una contraseña incorrecta.")
        Exit Sub

    End If
        
    If SlotMercader < 0 Or SlotMercader > mMao.MERCADER_MAX_LIST Then Exit Sub
        
    If MercaderActivate Then
        
        Mercader.Account = UserList(UserIndex).Account.Email
        Mercader.Gld = Gld
        Mercader.Dsp = Dsp
        Mercader.Desc = Desc
                
        If SlotMercader = 0 Then
            If UserList(UserIndex).Account.MercaderSlot > 0 Then
                Call WriteErrorMsg(UserIndex, "¡Ya tienes una publicación vigente! Elimina la que tienes para crear otra...")
                Exit Sub

            End If

        Else

            If StrComp(MercaderList(SlotMercader).Chars.Account, UserList(UserIndex).Account.Email) = 0 Then
                Call WriteErrorMsg(UserIndex, "¡No puedes hacer intercambios contigo mismo!")
                Exit Sub

            End If

        End If
                
        If SlotMercader = 0 Then
            Call mMao.Mercader_AddList(UserIndex, Chars, Mercader, Blocked)
        Else
            Call mMao.Mercader_AddOffer(UserIndex, Chars, SlotMercader, Mercader, Blocked)

        End If

    Else
        Call WriteErrorMsg(UserIndex, "Mercado desactivado temporalmente.")

    End If

    '<EhFooter>
    Exit Sub

HandleMercader_New_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleMercader_New " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleMercader_Required(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleMercader_Required_Err

    '</EhHeader>
    
    Dim Version  As String

    Dim Required As Byte

    Dim Value    As Long, Value1 As Long

    Required = Reader.ReadInt
    Value = Reader.ReadInt
    Value1 = Reader.ReadInt

    Select Case Required

        Case 0 ' Remover publicación.

            If UserList(UserIndex).Account.MercaderSlot > 0 Then
                Call Mercader_Remove(UserList(UserIndex).Account.MercaderSlot, UserList(UserIndex).Account.Email)
                Call WriteErrorMsg(UserIndex, "¡Has eliminado la publicación!")
                Call WriteAccountInfo(UserIndex)

            End If
                
        Case 1 ' Enviar Lista del Mercado.

            If Value <= 0 Then Exit Sub
            If Value1 <= 0 Then Exit Sub
            If Value > MERCADER_MAX_LIST Then Value = MERCADER_MAX_LIST
            If Value1 > MERCADER_MAX_LIST Then Value1 = MERCADER_MAX_LIST
            Call WriteMercader_List(UserIndex, Value, Value1, 0)

        Case 2 ' Enviar Información del listado seleccionado.

            If Value <= 0 Or Value > MERCADER_MAX_LIST Then Exit Sub
            If Value1 <= 0 Or Value1 > ACCOUNT_MAX_CHARS Then Exit Sub
            If MercaderList(Value).Chars.Account = vbNullString Then Exit Sub
            If MercaderList(Value).Chars.Count = 0 Then Exit Sub
            Call WriteMercader_ListChar(UserIndex, Value, Value1, False)

        Case 3 ' Enviar la lista de ofertas

            If Value <= 0 Or Value > MERCADER_MAX_LIST Then Exit Sub
            Call WriteMercader_List(UserIndex, 1, 50, Value)
                    
        Case 4 ' Envia la información de las ofertas

            If Value <= 0 Or Value > MERCADER_MAX_LIST Then Exit Sub
            If Value1 <= 0 Or Value1 > MERCADER_MAX_OFFER Then Exit Sub
            If MercaderList(Value).Offer(Value1).Account = vbNullString Then Exit Sub
            If MercaderList(Value).Offer(Value1).Count = 0 Then Exit Sub
            Call WriteMercader_ListChar(UserIndex, Value, Value1, True)
            
        Case 5 ' Acepta una oferta recibida

            If Value <= 0 Or Value > MERCADER_MAX_OFFER Then Exit Sub
            If UserList(UserIndex).Account.MercaderSlot = 0 Then Exit Sub
            If MercaderList(UserList(UserIndex).Account.MercaderSlot).Offer(Value).Account = vbNullString Then Exit Sub
            Call mMao.Mercader_AcceptOffer(UserIndex, UserList(UserIndex).Account.MercaderSlot, Value)

    End Select

    '<EhFooter>
    Exit Sub

HandleMercader_Required_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleMercader_Required " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub WriteMercader_List(ByVal UserIndex As Integer, _
                              ByVal aBound As Integer, _
                              ByVal bBound As Integer, _
                              ByVal MercaderSlot As Integer)

    '<EhHeader>
    On Error GoTo WriteMercader_List_Err

    '</EhHeader>

    Call Writer.WriteInt(ServerPacketID.Mercader_List)
    
    Dim A        As Long, B As Long

    Dim Text     As String
    
    Dim Mercader As tMercaderChar
    
    Call Writer.WriteInt16(aBound)
    Call Writer.WriteInt16(bBound)

    Call Writer.WriteInt16(UserList(UserIndex).Account.MercaderSlot)
    
    For A = aBound To bBound

        If MercaderSlot = 0 Then
            Mercader = MercaderList(A).Chars
        Else
            Mercader = MercaderList(MercaderSlot).Offer(A)

        End If
        
        With Mercader
            Call Writer.WriteInt16(A)
            Call Writer.WriteInt8(.Count)
            Call Writer.WriteString8(.Desc)
            Call Writer.WriteInt32(.Dsp)
            Call Writer.WriteInt32(.Gld)
            
            For B = 1 To .Count
                Call Writer.WriteString8(.NameU(B))
                Call Writer.WriteInt8(.Info(B).Class)
                Call Writer.WriteInt8(.Info(B).Raze)
                
                Call Writer.WriteInt8(.Info(B).Elv)
                Call Writer.WriteInt32(.Info(B).Exp)
                Call Writer.WriteInt32(.Info(B).Elu)
                    
                Call Writer.WriteInt16(.Info(B).Hp)
                Call Writer.WriteInt8(.Info(B).Constitucion)

                Call Writer.WriteInt16(.Info(B).Body)
                Call Writer.WriteInt16(.Info(B).Head)
                Call Writer.WriteInt16(.Info(B).Weapon)
                Call Writer.WriteInt16(.Info(B).Shield)
                Call Writer.WriteInt16(.Info(B).Helm)

            Next B

        End With

    Next A
    
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteMercader_List_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteMercader_List " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteMercader_ListChar(ByVal UserIndex As Integer, _
                                  ByVal Slot As Integer, _
                                  ByVal SlotChar As Integer, _
                                  ByVal InfoOffer As Boolean)

    '<EhHeader>
    On Error GoTo WriteMercader_ListChar_Err

    '</EhHeader>

    Call Writer.WriteInt(ServerPacketID.Mercader_ListInfo)
    
    Dim A        As Long, B As Long

    Dim Text     As String

    Dim Mercader As tMercaderChar
    
    Call Writer.WriteInt16(SlotChar)
    
    If Not InfoOffer Then
        Mercader = MercaderList(Slot).Chars
    Else
        Mercader = MercaderList(Slot).Offer(SlotChar)

    End If
        
    With Mercader
        Call Writer.WriteString8(.NameU(SlotChar))
        Call Writer.WriteString8(GuildsInfo(.Info(SlotChar).GuildIndex).Name)
        
        Call Writer.WriteInt32(.Info(SlotChar).Gld)
        
        Call Writer.WriteInt16(.Info(SlotChar).Body)
        Call Writer.WriteInt16(.Info(SlotChar).Head)
        Call Writer.WriteInt16(.Info(SlotChar).Weapon)
        Call Writer.WriteInt16(.Info(SlotChar).Shield)
        Call Writer.WriteInt16(.Info(SlotChar).Helm)
        
        Call Writer.WriteInt8(.Info(SlotChar).Faction)
        Call Writer.WriteInt8(.Info(SlotChar).FactionRange)
        Call Writer.WriteInt16(.Info(SlotChar).FragsCiu)
        Call Writer.WriteInt16(.Info(SlotChar).FragsCri)
        
        For A = 1 To MAX_INVENTORY_SLOTS
            Call Writer.WriteInt16(.Info(SlotChar).Object(A).ObjIndex)
            Call Writer.WriteInt16(.Info(SlotChar).Object(A).Amount)
        Next A
            
        For A = 1 To MAX_BANCOINVENTORY_SLOTS
            Call Writer.WriteInt16(.Info(SlotChar).Bank(A).ObjIndex)
            Call Writer.WriteInt16(.Info(SlotChar).Bank(A).Amount)
        Next A
        
        For A = 1 To 35

            If .Info(SlotChar).Spells(A) > 0 Then
                Call Writer.WriteString8(Hechizos(.Info(SlotChar).Spells(A)).Nombre)
            Else
                Call Writer.WriteString8(vbNullString)

            End If

        Next A
        
        For A = 1 To NUMSKILLS
            Call Writer.WriteInt8(.Info(SlotChar).Skills(A))
        Next A
        
    End With
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteMercader_ListChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteMercader_ListChar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteMercader_ListOffer(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteMercader_ListOffer_Err

    '</EhHeader>

    Call Writer.WriteInt(ServerPacketID.Mercader_ListOffer)
    
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteMercader_ListOffer_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteMercader_ListOffer " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleForgive_Faction(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleForgive_Faction_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If

        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub

        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If

        If .Faction.Status > 0 Then
            Call WriteConsoleMsg(UserIndex, "Ya eres miembro de una facción y espero que sea la nuestra, sino mis guardias te atacaran!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            If Escriminal(UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "¡¡Sal de aquí, antes de que mis guardias acaben contigo!!", FontTypeNames.FONTTYPE_WARNING)

                Exit Sub

            End If
            
            If Not TieneObjetos(1086, 1, UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "¡No has reclamado tu recompensa! Debes hacer la misión que te otorga el fragmento necesario para el perdón", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub

            End If
            
            Call QuitarObjetos(1086, 1, UserIndex)
            
            UserList(UserIndex).Faction.FragsCiu = 0
            
        Else

            If Not Escriminal(UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "¡¡Sal de aquí, antes de que mis guardias acaben contigo!!", FontTypeNames.FONTTYPE_WARNING)

                Exit Sub

            End If
            
            If Not TieneObjetos(1087, 1, UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "¡No has reclamado tu recompensa! Debes hacer la misión que te otorga el fragmento necesario para el perdón", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub

            End If
            
            Call QuitarObjetos(1087, 1, UserIndex)

        End If
        
        Call Faction_RemoveUser(UserIndex)
        Call WriteConsoleMsg(UserIndex, "¡Te hemos perdonado, pero no abuses de nuestra bondad. Nuestras tropas son fieles y no toleran estupideces!", FontTypeNames.FONTTYPE_DEMONIO)
        
    End With

    '<EhFooter>
    Exit Sub

HandleForgive_Faction_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleForgive_Faction " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleMap_RequiredInfo(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleMap_RequiredInfo_Err

    '</EhHeader>
    
    Dim Map As Integer
    
    Map = Reader.ReadInt
    
    If Map = 0 Or Map > NumMaps Then Exit Sub
    
    Call WriteMiniMap_InfoCriature(UserIndex, Map)

    '<EhFooter>
    Exit Sub

HandleMap_RequiredInfo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleMap_RequiredInfo " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteRender_CountDown(ByVal UserIndex As Integer, ByVal CountDown As Long)

    '<EhHeader>
    On Error GoTo WriteRender_CountDown_Err

    '</EhHeader>
    Call SendData(ToOne, UserIndex, PrepareMessageRender_CountDown(CountDown))
    '<EhFooter>
    Exit Sub

WriteRender_CountDown_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteRender_CountDown " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function PrepareMessageRender_CountDown(ByVal Time As Long) As String

    '<EhHeader>
    On Error GoTo PrepareMessageRender_CountDown_Err

    '</EhHeader>

    Call Writer.WriteInt(ServerPacketID.Render_CountDown)
    Call Writer.WriteInt(Time)
        
    '<EhFooter>
    Exit Function

PrepareMessageRender_CountDown_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageRender_CountDown " & "at line " & Erl
        
    '</EhFooter>
End Function

Private Sub WriteMiniMap_InfoCriature(ByVal UserIndex As Integer, ByVal Map As Integer)

    '<EhHeader>
    On Error GoTo WriteMiniMap_InfoCriature_Err

    '</EhHeader>

    Dim A    As Long, B As Long

    Dim Str  As String

    Dim Temp As String

    Call Writer.WriteInt(ServerPacketID.MiniMap_InfoCriature)
    Call Writer.WriteInt(Map)
    Call Writer.WriteInt(MiniMap(Map).NpcsNum)
    Call Writer.WriteString8(MiniMap(Map).Name)
    Call Writer.WriteBool(MiniMap(Map).Pk)
    Call Writer.WriteInt(MiniMap(Map).LvlMin)
    Call Writer.WriteInt(MiniMap(Map).LvlMax)
            
    If MiniMap(Map).NpcsNum Then

        With MiniMap(Map)
        
            For A = 1 To MiniMap(Map).NpcsNum
                Call Writer.WriteInt16(MiniMap(Map).Npcs(A).NpcIndex)
                Call Writer.WriteString8(MiniMap(Map).Npcs(A).Name)
                Call Writer.WriteInt(MiniMap(Map).Npcs(A).Body)
                Call Writer.WriteInt(MiniMap(Map).Npcs(A).Head)
                Call Writer.WriteInt(MiniMap(Map).Npcs(A).Hp)
                Call Writer.WriteInt(MiniMap(Map).Npcs(A).MinHit)
                Call Writer.WriteInt(MiniMap(Map).Npcs(A).MaxHit)
                Call Writer.WriteInt(MiniMap(Map).Npcs(A).Exp)
                Call Writer.WriteInt(MiniMap(Map).Npcs(A).Gld)
                Call Writer.WriteInt(MiniMap(Map).Npcs(A).Eldhir)
                
                'Spells
                Call Writer.WriteInt(MiniMap(Map).Npcs(A).NroSpells)
                
                If MiniMap(Map).Npcs(A).NroSpells Then
    
                    For B = 1 To MiniMap(Map).Npcs(A).NroSpells
                        Call Writer.WriteString8(Hechizos(MiniMap(Map).Npcs(A).Spells(B)).Nombre)
                    Next B
    
                End If
                
                ' Inventario de la Criatura
                Call Writer.WriteInt(MiniMap(Map).Npcs(A).NroItems)
                
                For B = 1 To MiniMap(Map).Npcs(A).NroItems
                    Temp = vbNullString
                        
                    If MiniMap(Map).Npcs(A).Invent.Object(B).ObjIndex > 0 Then
                        Temp = ObjData(MiniMap(Map).Npcs(A).Invent.Object(B).ObjIndex).Name

                    End If
                        
                    Call Writer.WriteString8(Temp)
                    Call Writer.WriteInt(MiniMap(Map).Npcs(A).Invent.Object(B).Amount)
                Next B
                
                ' Drops de la Criatura
                Call Writer.WriteInt(MiniMap(Map).Npcs(A).NroDrops)
                
                For B = 1 To MiniMap(Map).Npcs(A).NroDrops
                    Temp = vbNullString
                        
                    If MiniMap(Map).Npcs(A).Drop(B).ObjIndex > 0 Then
                        Temp = ObjData(MiniMap(Map).Npcs(A).Drop(B).ObjIndex).Name

                    End If
                        
                    Call Writer.WriteString8(Temp)
                    Call Writer.WriteInt(MiniMap(Map).Npcs(A).Drop(B).Amount)
                    Call Writer.WriteInt(MiniMap(Map).Npcs(A).Drop(B).Probability)
                Next B
            
            Next A
        
        End With
        
    End If
        
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteMiniMap_InfoCriature_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteMiniMap_InfoCriature " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleWherePower(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleWherePower_Err

    '</EhHeader>
    
    If Power.UserIndex = 0 Then
        Call WriteConsoleMsg(UserIndex, "Ningún usuario posee el don.", FontTypeNames.FONTTYPE_INFORED)
    Else
        Call WriteConsoleMsg(UserIndex, "El poseedor del poder es el personaje " & UserList(Power.UserIndex).Name & " en el mapa " & MapInfo(UserList(Power.UserIndex).Pos.Map).Name, FontTypeNames.FONTTYPE_INFOGREEN)

    End If

    '<EhFooter>
    Exit Sub

HandleWherePower_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleWherePower " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleAuction_New(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleAuction_New_Err

    '</EhHeader>

    Dim Slot   As Byte

    Dim Amount As Integer

    Dim Gld    As Long

    Dim Eldhir As Long
    
    Slot = Reader.ReadInt
    Amount = Reader.ReadInt
    Gld = Reader.ReadInt
    Eldhir = Reader.ReadInt
    
    If Slot <= 0 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub
    If Amount <= 0 Or Amount > 10000 Then Exit Sub
    If Gld < 0 Or Gld > 100000000 Then Exit Sub
    If Eldhir < 0 Or Eldhir > 1000 Then Exit Sub
    If UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
    If UserList(UserIndex).Invent.Object(Slot).Amount < Amount Then Exit Sub
    If UserList(UserIndex).flags.Bronce = 0 Then
        Call WriteConsoleMsg(UserIndex, "Debes ser [BRONCE] para poder subastar objetos.", FontTypeNames.FONTTYPE_USERBRONCE)
        Exit Sub

    End If
    
    Call Auction_CreateNew(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex, Amount, Gld, Eldhir)
    '<EhFooter>
    Exit Sub

HandleAuction_New_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleAuction_New " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleAuction_Info(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleAuction_Info_Err

    '</EhHeader>
    
    If Auction.ObjIndex = 0 Then
        Call WriteConsoleMsg(UserIndex, "¡No hay ninguna subasta en trámite!", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
    
    Call WriteConsoleMsg(UserIndex, "El personaje " & Auction.Name & " está subastando " & ObjData(Auction.ObjIndex).Name & " (x" & Auction.Amount & "). Deberás ofrecer como mínimo: " & Auction.Offer.Gld * 1.1 & " Monedas de Oro Y " & Auction.Offer.Eldhir & " Monedas de Eldhir.", FontTypeNames.FONTTYPE_INFOGREEN)
    '<EhFooter>
    Exit Sub

HandleAuction_Info_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleAuction_Info " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleAuction_Offer(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleAuction_Offer_Err

    '</EhHeader>
    
    Dim Gld    As Long

    Dim Eldhir As Long
    
    Gld = Reader.ReadInt
    Eldhir = Reader.ReadInt
    
    If Gld < 0 Or Gld > 1000000000 Then Exit Sub
    If Eldhir < 0 Or Eldhir > 5000 Then Exit Sub
    
    Call mAuction.Auction_Offer(UserIndex, Gld, Eldhir)
    '<EhFooter>
    Exit Sub

HandleAuction_Offer_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleAuction_Offer " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleGoInvation(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleGoInvation_Err

    '</EhHeader>
    
    With UserList(UserIndex)
        
        Dim Slot As Byte
        
        Slot = Reader.ReadInt8
        
        If Slot <= 0 Or Slot > UBound(Invations) Then Exit Sub
        
        If Not .Pos.Map = Ullathorpe.Map Then
            Call WriteConsoleMsg(UserIndex, "Solo puedes ingresar a la invasión estando en Ullathorpe.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Invations(Slot).Run = False Then
            Call WriteConsoleMsg(UserIndex, "El evento no se encuentra disponible.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Call EventWarpUser(UserIndex, Invations(Slot).InitialMap, Invations(Slot).InitialX, Invations(Slot).InitialY)
        Call WriteConsoleMsg(UserIndex, "¡Bienvenido a " & Invations(Slot).Name & "! Esperemos que te diviertas y compartas tu experiencia con el resto de los usuarios. ¡Suerte!", FontTypeNames.FONTTYPE_INVASION)
        
        .Counters.Shield = 3
        
    End With
    
    '<EhFooter>
    Exit Sub

HandleGoInvation_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleGoInvation " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleSendDataUser(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleSendDataUser_Err

    '</EhHeader>
    
    Dim UserName As String

    Dim tUser    As Integer

    Dim Message  As String
    
    UserName = Reader.ReadString8
    
    If Not EsGmPriv(UserIndex) Then Exit Sub
    
    tUser = NameIndex(UserName)
    
    If tUser > 0 Then

        With UserList(tUser)
            Message = "'DATOS DE " & UserName & "'"
            Message = Message & vbCrLf & "IP PUBLICA: " & .Account.Sec.IP_Public
            Message = Message & vbCrLf & "IP ADDRESS: " & .Account.Sec.IP_Address
            Message = Message & vbCrLf & "IP LOCAL: " & .Account.Sec.IP_Local
            Message = Message & vbCrLf & "MAC ADDRESS: " & .Account.Sec.SERIAL_MAC
            Message = Message & vbCrLf & "DISCO: " & .Account.Sec.SERIAL_DISK
            Message = Message & vbCrLf & "BIOS: " & .Account.Sec.SERIAL_BIOS
            Message = Message & vbCrLf & "MOTHERBOARD: " & .Account.Sec.SERIAL_MOTHERBOARD
            Message = Message & vbCrLf & "PROCESSOR: " & .Account.Sec.SERIAL_PROCESSOR
            Message = Message & vbCrLf & "SYSTEM DATA " & .Account.Sec.SYSTEM_DATA

        End With
        
        Call WriteConsoleMsg(UserIndex, Message, FontTypeNames.FONTTYPE_INFOGREEN)
    Else
        Call WriteConsoleMsg(UserIndex, "El personaje está offline.", FontTypeNames.FONTTYPE_INFORED)

    End If
    
    '<EhFooter>
    Exit Sub

HandleSendDataUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleSendDataUser " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleSearchDataUser(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleSearchDataUser_Err

    '</EhHeader>
    Dim Data     As String

    Dim Selected As eSearchData
    
    Selected = Reader.ReadInt8
    Data = Reader.ReadString8
    
    If Not EsGmPriv(UserIndex) Then Exit Sub
    If Data = vbNullString Then Exit Sub
    If Selected <= 0 Or Selected > 3 Then Exit Sub
    
    Call Security_SearchData(UserIndex, Selected, Data)
    '<EhFooter>
    Exit Sub

HandleSearchDataUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleSearchDataUser " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleChangeModoArgentum(ByVal UserIndex As Integer)

    If Not EsGmPriv(UserIndex) Then Exit Sub
    
    ' Cambia el Uso de Paquetes
    If PacketUseItem = ClientPacketID.UseItem Then
        PacketUseItem = ClientPacketID.UseItemTwo
        
        EsModoEvento = 1
        Call WriteConsoleMsg(UserIndex, "¡Has pasado al MODO EVENTO!", FontTypeNames.FONTTYPE_INFO)
    Else
        PacketUseItem = ClientPacketID.UseItem
        
        Call WriteConsoleMsg(UserIndex, "¡Has vuelto al MODO Default!", FontTypeNames.FONTTYPE_INFO)
        EsModoEvento = 0
        
    End If
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageUpdateEvento(EsModoEvento))

End Sub

Public Function WriteUpdateEffect(ByVal UserIndex As Integer) As String

    '<EhHeader>
    On Error GoTo WriteUpdateEffect_Err

    '</EhHeader>

    '***************************************************
    ' Actualiza distintos tipos de efectos
    ' Efecto n°1: Veneno (Efecto Verde)
    '
    '***************************************************

    Call Writer.WriteInt(ServerPacketID.UpdateEffectPoison)
        
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Function

WriteUpdateEffect_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateEffect " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Prepares the "CreateFXMap" message and returns it.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCreateFXMap(ByVal X As Byte, _
                                          ByVal Y As Byte, _
                                          ByVal FX As Integer, _
                                          ByVal FXLoops As Integer) As String

    '<EhHeader>
    On Error GoTo PrepareMessageCreateFXMap_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification:
    'Prepares the "CreateFXMap" message and returns it
    '***************************************************
    
    Call Writer.WriteInt(ServerPacketID.CreateFXMap)
    Call Writer.WriteInt(X)
    Call Writer.WriteInt(Y)
    Call Writer.WriteInt(FX)
    Call Writer.WriteInt(FXLoops)

    '<EhFooter>
    Exit Function

PrepareMessageCreateFXMap_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageCreateFXMap " & "at line " & Erl
        
    '</EhFooter>
End Function

Private Sub HandleEvents_DonateObject(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleEvents_DonateObject_Err

    '</EhHeader>

    Dim Slot      As Byte

    Dim SlotEvent As Byte

    Dim Amount    As Integer

    SlotEvent = Reader.ReadInt8
    Slot = Reader.ReadInt8
    Amount = Reader.ReadInt16
    
    If SlotEvent <= 0 Or SlotEvent > UBound(Events) Then Exit Sub
    If Slot <= 0 Or Slot >= MAX_INVENTORY_SLOTS Then Exit Sub
    If Amount > UserList(UserIndex).Invent.Object(Slot).Amount Or Amount >= MAX_INVENTORY_OBJS Then Exit Sub
    If UserList(UserIndex).Invent.Object(Slot).ObjIndex <= 0 Then Exit Sub
    If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Donable = 0 Then
        Call WriteConsoleMsg(UserIndex, "¡Ha ha ha tu objeto no es tolerado aquí!", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If

    Call EventosDS_Reward.Events_Reward_Add(UserIndex, SlotEvent, Slot, Amount)
   
    '<EhFooter>
    Exit Sub

HandleEvents_DonateObject_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleEvents_DonateObject " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function PrepareMessageRenderConsole(ByVal Text As String, _
                                            ByVal DamageType As eDamageType, _
                                            ByVal Duration As Long, _
                                            ByVal Slot As Byte)

    '<EhHeader>
    On Error GoTo PrepareMessageRenderConsole_Err

    '</EhHeader>
 
    Writer.WriteInt ServerPacketID.RenderConsole
    Writer.WriteString8 Text
    Writer.WriteInt8 DamageType
    Writer.WriteInt32 Duration
    Writer.WriteInt8 Slot

    '<EhFooter>
    Exit Function

PrepareMessageRenderConsole_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageRenderConsole " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub WriteViewListQuest(ByVal UserIndex As Integer, _
                              ByRef Quest() As Byte, _
                              ByVal NameNpc As String)

    '<EhHeader>
    On Error GoTo WriteViewListQuest_Err

    '</EhHeader>
    
    Dim A As Long, B As Long
    
    Call Writer.WriteInt(ServerPacketID.ViewListQuest)
    Call Writer.WriteInt8(UBound(Quest))
    Call Writer.WriteString8(NameNpc)
    
    For A = LBound(Quest) To UBound(Quest)
        Call Writer.WriteInt8(Quest(A))
    Next A

    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteViewListQuest_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteViewListQuest " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteUpdateUserDead(ByVal UserIndex As Integer, ByVal UserMuerto As Byte)

    '<EhHeader>
    On Error GoTo WriteUpdateUserDead_Err

    '</EhHeader>
                                
    Call Writer.WriteInt(ServerPacketID.UpdateUserDead)
    Call Writer.WriteInt8(UserMuerto)
    
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteUpdateUserDead_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateUserDead " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteQuestInfo(ByVal UserIndex As Integer, _
                          ByVal Visible As Boolean, _
                          ByVal Slot As Integer)

    '<EhHeader>
    On Error GoTo WriteQuestInfo_Err

    '</EhHeader>
                                
    Call Writer.WriteInt(ServerPacketID.QuestInfo)
        
    Dim A          As Long
        
    Dim i          As Long

    Dim QuestIndex As Integer
    
    Call Writer.WriteBool(Visible)
    Call Writer.WriteInt16(Slot)
    
    If Slot <> 0 Then
        QuestIndex = UserList(UserIndex).QuestStats(Slot).QuestIndex
                
        Call Writer.WriteInt16(QuestIndex)
                
        If UserList(UserIndex).QuestStats(Slot).QuestIndex > 0 Then
            
            If QuestList(QuestIndex).RequiredNPCs > 0 Then
    
                For i = LBound(UserList(UserIndex).QuestStats(Slot).NPCsKilled) To UBound(UserList(UserIndex).QuestStats(Slot).NPCsKilled)
                    Call Writer.WriteInt32(UserList(UserIndex).QuestStats(Slot).NPCsKilled(i))
                Next i
    
            End If
                        
            If QuestList(QuestIndex).RequiredSaleOBJs > 0 Then
    
                For i = LBound(UserList(UserIndex).QuestStats(Slot).ObjsSale) To UBound(UserList(UserIndex).QuestStats(Slot).ObjsSale)
                    Call Writer.WriteInt32(UserList(UserIndex).QuestStats(Slot).ObjsSale(i))
                Next i
    
            End If
                        
            If QuestList(QuestIndex).RequiredChestOBJs > 0 Then
    
                For i = LBound(UserList(UserIndex).QuestStats(Slot).ObjsPick) To UBound(UserList(UserIndex).QuestStats(Slot).ObjsPick)
                    Call Writer.WriteInt32(UserList(UserIndex).QuestStats(Slot).ObjsPick(i))
                Next i
    
            End If
                        
        End If

    Else
    
        For A = 1 To MAXUSERQUESTS
            QuestIndex = UserList(UserIndex).QuestStats(A).QuestIndex
                
            Call Writer.WriteInt16(QuestIndex)
                
            If UserList(UserIndex).QuestStats(A).QuestIndex > 0 Then
            
                If QuestList(QuestIndex).RequiredNPCs > 0 Then
    
                    For i = LBound(UserList(UserIndex).QuestStats(A).NPCsKilled) To UBound(UserList(UserIndex).QuestStats(A).NPCsKilled)
                        Call Writer.WriteInt32(UserList(UserIndex).QuestStats(A).NPCsKilled(i))
                    Next i
    
                End If
                        
                If QuestList(QuestIndex).RequiredSaleOBJs > 0 Then
    
                    For i = LBound(UserList(UserIndex).QuestStats(A).ObjsSale) To UBound(UserList(UserIndex).QuestStats(A).ObjsSale)
                        Call Writer.WriteInt32(UserList(UserIndex).QuestStats(A).ObjsSale(i))
                    Next i
    
                End If
                        
                If QuestList(QuestIndex).RequiredChestOBJs > 0 Then
    
                    For i = LBound(UserList(UserIndex).QuestStats(A).ObjsPick) To UBound(UserList(UserIndex).QuestStats(A).ObjsPick)
                        Call Writer.WriteInt32(UserList(UserIndex).QuestStats(A).ObjsPick(i))
                    Next i
    
                End If
                        
            End If

        Next A
    
    End If

    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteQuestInfo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteQuestInfo " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Private Sub HandleQuestRequired(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleQuestRequired_Err

    '</EhHeader>
    
    Dim Tipo As Byte

    Tipo = Reader.ReadInt8
    
    Call WriteQuestInfo(UserIndex, True, 0)
    '<EhFooter>
    Exit Sub

HandleQuestRequired_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleQuestRequired " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteUpdateGlobalCounter(ByVal UserIndex As Integer, _
                                    ByVal Tipo As Byte, _
                                    ByVal Counter As Integer)

    '<EhHeader>
    On Error GoTo WriteUpdateGlobalCounter_Err

    '</EhHeader>
    Call Writer.WriteInt(ServerPacketID.UpdateGlobalCounter)
    
    Call Writer.WriteInt8(Tipo)
    Call Writer.WriteInt16(Counter)

    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteUpdateGlobalCounter_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateGlobalCounter " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteSendIntervals(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteSendIntervals_Err

    '</EhHeader>

    Call Writer.WriteInt(ServerPacketID.SendIntervals)

    Call Writer.WriteInt16(IntervaloUserPuedeAtacar)
    Call Writer.WriteInt16(IntervaloUserPuedeUsar)
    Call Writer.WriteInt16(IntervaloUserPuedeUsarClick)
    Call Writer.WriteInt16(2000) ' Actualizar POS
    Call Writer.WriteInt16(IntervaloUserPuedeCastear)
    Call Writer.WriteInt16(IntervaloUserPuedeShiftear)
    Call Writer.WriteInt16(IntervaloFlechasCazadores)
    Call Writer.WriteInt16(IntervaloMagiaGolpe)
    Call Writer.WriteInt16(IntervaloGolpeMagia)
    Call Writer.WriteInt16(IntervaloGolpeUsar)
    Call Writer.WriteInt16(IntervaloUserPuedeTrabajar)
    Call Writer.WriteInt16(IntervalDrop)
    Call Writer.WriteReal32(IntervaloCaminar)
          
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteSendIntervals_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteSendIntervals " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteSendInfoNpc(ByVal UserIndex As Integer, ByVal number As Integer)

    '<EhHeader>
    On Error GoTo WriteSendInfoNpc_Err

    '</EhHeader>

    Call Writer.WriteInt(ServerPacketID.SendInfoNpc)

    Call Writer.WriteInt16(number)
   
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteSendInfoNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteSendInfoNpc " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteUpdatePosGuild(ByVal UserIndex As Integer, _
                               ByVal SlotMember As Byte, _
                               ByVal tUser As Integer)

    '<EhHeader>
    On Error GoTo WriteUpdatePosGuild_Err

    '</EhHeader>
    Call Writer.WriteInt(ServerPacketID.UpdatePosGuild)

    Call Writer.WriteInt8(SlotMember)
    
    If tUser > 0 Then
        Call Writer.WriteInt8(UserList(tUser).Pos.X)
        Call Writer.WriteInt8(UserList(tUser).Pos.Y)
    Else
        Call Writer.WriteInt8(0)
        Call Writer.WriteInt8(0)

    End If
    
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteUpdatePosGuild_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdatePosGuild " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function PrepareUpdateLevelGuild(ByVal LevelGuild As Byte)

    '<EhHeader>
    On Error GoTo PrepareUpdateLevelGuild_Err

    '</EhHeader>
    Call Writer.WriteInt(ServerPacketID.UpdateLevelGuild)
    Call Writer.WriteInt8(LevelGuild)
    '<EhFooter>
    Exit Function

PrepareUpdateLevelGuild_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareUpdateLevelGuild " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub WriteUpdateStatusMAO(ByVal UserIndex As Integer, ByVal Status As Byte)

    '<EhHeader>
    On Error GoTo WriteUpdateStatusMAO_Err

    '</EhHeader>
    Call Writer.WriteInt(ServerPacketID.UpdateStatusMAO)
    Call Writer.WriteInt8(Status)
    Call SendData(ToOne, UserIndex, vbNullString)

    '<EhFooter>
    Exit Sub

WriteUpdateStatusMAO_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateStatusMAO " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub HandleChangeClass(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeClass_Err

    '</EhHeader>
    Dim Clase  As Byte

    Dim Raza   As Byte

    Dim Genero As Byte
    
    Clase = Reader.ReadInt8
    Raza = Reader.ReadInt8
    Genero = Reader.ReadInt8
    
    If Clase <= 0 Or Clase > NUMCLASES Then Exit Sub
    If Raza <= 0 Or Raza > NUMRAZAS Then Exit Sub
    If Genero <= 0 Or Genero > 2 Then Exit Sub
    
    #If Classic = 1 Then
        Exit Sub
    #End If
    
    With UserList(UserIndex)
        .Clase = Clase
        .Raza = Raza
        .Genero = Genero
        .flags.Muerto = 0
        
        Call InitialUserStats(UserList(UserIndex))
        Call UserLevelEditation(UserList(UserIndex), STAT_MAXELV, 0)
        Call WriteUpdateUserStats(UserIndex)
        Call WriteUpdateHungerAndThirst(UserIndex)
        Call LoadSetInitial_Class(UserIndex)
        'Call LoadSetInitial_Class(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Ahora eres un " & ListaClases(.Clase) & " " & ListaRazas(.Raza) & ".", FontTypeNames.FONTTYPE_INFO)

    End With
    
    '<EhFooter>
    Exit Sub

HandleChangeClass_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeClass " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function PrepareMessageUpdateOnline() As String

    '<EhHeader>
    On Error GoTo PrepareMessageUpdateOnline_Err

    '</EhHeader>

    Call Writer.WriteInt(ServerPacketID.UpdateOnline)
    Call Writer.WriteInt16(NumUsers + UsersBot)
        
    '<EhFooter>
    Exit Function

PrepareMessageUpdateOnline_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageUpdateOnline " & "at line " & Erl
        
    '</EhFooter>
End Function

' SEGURIDAD
Public Function PrepareMessageUpdateEvento(ByVal ModoEvento As Byte) As String

    '<EhHeader>
    On Error GoTo PrepareMessageUpdateEvento_Err

    '</EhHeader>

    Call Writer.WriteInt(ServerPacketID.UpdateEvento)
    Call Writer.WriteInt8(ModoEvento)
        
    '<EhFooter>
    Exit Function

PrepareMessageUpdateEvento_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageUpdateEvento " & "at line " & Erl
        
    '</EhFooter>
End Function

Private Sub HandleModoStreamer(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleModoStreamer_Err

    '</EhHeader>
    
    With UserList(UserIndex)

        If .flags.StreamUrl = vbNullString Then
            Call WriteConsoleMsg(UserIndex, "Por favor setea primero una URL con el comando /STREAMLINK. ¡Vende tu contenido! Sé hábil para poner alguna frase que haga que las personas ingresen a tu canal haciendo clic! NO muy largo.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If
            
        .flags.ModoStream = Not .flags.ModoStream
            
        If .flags.ModoStream Then
            Call WriteMultiMessage(UserIndex, eMessages.ModoStreamOn)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.ModoStreamOff)

        End If
                
        ' Call Streamer_Can(UserIndex)
            
        'Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
    
    End With
  
    '<EhFooter>
    Exit Sub

HandleModoStreamer_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleModoStreamer " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleStreamerSetLink(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleStreamerSetLink_Err

    '</EhHeader>
    
    Dim Url As String
    
    Url = Reader.ReadString8
    
    With UserList(UserIndex)
        .flags.StreamUrl = Url
        Call WriteConsoleMsg(UserIndex, "¡La URL del Twitch ha pasado a ser " & Url & "!", FontTypeNames.FONTTYPE_INFOGREEN)
        
        Call Logs_Security(eSecurity, eAntiHack, "La cuenta " & .Account.Email & " con personaje: " & .Name & " ha cambiado su link de Twitch a " & Url & ".")

    End With
   
    '<EhFooter>
    Exit Sub

HandleStreamerSetLink_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleStreamerSetLink " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleChangeNick(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleChangeNick_Err

    '</EhHeader>
    Dim UserName As String

    Dim Leader   As Boolean
    
    UserName = Reader.ReadString8
    Leader = Reader.ReadBool
    
    If Leader Then
        Call mGuilds.ChangeLeader(UserIndex, UCase$(UserName))
    Else
        Call mAccount.ChangeNickChar(UserIndex, UCase$(UserName))

    End If
   
    '<EhFooter>
    Exit Sub

HandleChangeNick_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleChangeNick " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleConfirmTransaccion(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleConfirmTransaccion_Err

    '</EhHeader>
    Dim Waiting As tShopWaiting
    
    Waiting.Email = Reader.ReadString8
    Waiting.Promotion = Reader.ReadInt8
    Waiting.Bank = Reader.ReadString8
    
    Call mShop.Transaccion_Add(UserIndex, Waiting)
    
    '<EhFooter>
    Exit Sub

HandleConfirmTransaccion_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleConfirmTransaccion " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleConfirmItem(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleConfirmItem_Err

    '</EhHeader>
    Dim ID            As Integer

    Dim PrioriceValue As Byte
        
    ID = Reader.ReadInt16
    PrioriceValue = Reader.ReadInt8
          
    If ID <= 0 Or ID > ShopLast Then Exit Sub ' Anti Hack
    If PrioriceValue > 1 Then Exit Sub
          
    Call mShop.ConfirmItem(UserIndex, ID, PrioriceValue)
    
    '<EhFooter>
    Exit Sub

HandleConfirmItem_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleConfirmItem " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleConfirmTier(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleConfirmTier_Err

    '</EhHeader>
    Dim Tier As Byte
    
    Tier = Reader.ReadInt8
    
    Call mShop.ConfirmTier(UserIndex, Tier)
    '<EhFooter>
    Exit Sub

HandleConfirmTier_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleConfirmTier " & "at line " & Erl
        
    '</EhFooter>
End Sub

' SEGURIDAD
Public Function PrepareMessageUpdateMeditation(ByRef MeditationUser() As Integer, _
                                               ByVal MeditationAnim As Byte) As String

    '<EhHeader>
    On Error GoTo PrepareMessageUpdateMeditation_Err

    '</EhHeader>

    Call Writer.WriteInt(ServerPacketID.UpdateMeditation)
    Call Writer.WriteInt16(MeditationAnim)
    
    Call Writer.WriteInt8(MAX_MEDITATION)
    
    Dim A As Long
    
    For A = 1 To MAX_MEDITATION
        Call Writer.WriteInt8(MeditationUser(A))
    Next A
    
    '<EhFooter>
    Exit Function

PrepareMessageUpdateMeditation_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.PrepareMessageUpdateMeditation " & "at line " & Erl
        
    '</EhFooter>
End Function

Private Sub HandleRequiredShopChars(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleRequiredShopChars_Err

    '</EhHeader>
    If Not Interval_Packet250(UserIndex) Then Exit Sub
    Call WriteShopChars(UserIndex)
    '<EhFooter>
    Exit Sub

HandleRequiredShopChars_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleRequiredShopChars " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteShopChars(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteShopChars_Err

    '</EhHeader>
    Call Writer.WriteInt(ServerPacketID.SendShopChars)
    
    Dim A As Long
    
    Call Writer.WriteInt8(ShopCharLast)
    
    For A = 1 To ShopCharLast

        With ShopChars(A)
            Call Writer.WriteString8(.Name)
            Call Writer.WriteInt16(.Dsp)
            
            Call Writer.WriteInt8(.Elv)
            Call Writer.WriteInt8(.Porc)
            Call Writer.WriteInt8(.Class)
            Call Writer.WriteInt8(.Raze)
            Call Writer.WriteInt16(.Head)
            Call Writer.WriteInt16(.Hp)
            Call Writer.WriteInt16(.Man)

        End With

    Next A
    
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteShopChars_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteShopChars " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleConfirmChar(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleConfirmChar_Err

    '</EhHeader>
    
    Dim ID As Byte
    
    ID = Reader.ReadInt8
    
    If ID <= 0 Or ID > ShopCharLast Then Exit Sub
    
    Call mShop.ConfirmChar(UserIndex, ID)
    
    '<EhFooter>
    Exit Sub

HandleConfirmChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleConfirmChar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleConfirmQuest(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleConfirmQuest_Err

    '</EhHeader>
    
    Dim Tipo  As Byte

    Dim Quest As Byte
    
    Tipo = Reader.ReadInt8
    Quest = Reader.ReadInt8
    
    Select Case Tipo
    
        Case 1 ' Reclamar Mision
                 
            If UserList(UserIndex).QuestStats(Quest).QuestIndex Then
                If Quests_CheckFinish(UserIndex, Quest) Then
                    Call mQuests.Quests_Next(UserIndex, Quest)

                End If
                  
            End If
            
        Case 2 ' Confirmar para hacer una de las de alto riesgo
    
    End Select
    
    '<EhFooter>
    Exit Sub

HandleConfirmQuest_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleConfirmQuest " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WriteUpdateFinishQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUpdateFinishQuest_Err

    '</EhHeader>
                                              
    Call Writer.WriteInt(ServerPacketID.UpdateFinishQuest)
    Call Writer.WriteInt16(QuestIndex)
    
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteUpdateFinishQuest_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateFinishQuest " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleRequiredSkin(ByVal UserIndex As Integer)
        
    If Not Interval_Packet250(UserIndex) Then Exit Sub
        
    Dim ObjIndex As Integer

    Dim Modo     As Byte
    
    ObjIndex = Reader.ReadInt16
    Modo = Reader.ReadInt8
    
    If ObjIndex > 0 Then
        If Modo = 3 Then    ' Desequipar
            Call Skins_Desequipar(UserIndex, ObjIndex)
        Else
            Call mSkins.Skins_AddNew(UserIndex, ObjIndex)

        End If
        
    Else
        WriteUpdateDataSkin UserIndex, UserList(UserIndex).Skins.Last

    End If
        
End Sub

Public Sub WriteUpdateDataSkin(ByVal UserIndex As Integer, ByVal Last As Integer)

    '<EhHeader>
    On Error GoTo WriteUpdateFinishQuest_Err

    '</EhHeader>
                                              
    Call Writer.WriteInt(ServerPacketID.UpdateDataSkin)

    Dim Data As tSkins

    Dim A    As Long
            
    Data = UserList(UserIndex).Skins
    Call Writer.WriteInt16(Last)
            
    If Last > 0 Then
        Data.Last = Last
                
        If Data.Last > 0 Then
                    
            For A = 1 To Data.Last
                Call Writer.WriteInt16(Data.ObjIndex(A))
            Next A

        End If

    End If
               
    Call Writer.WriteInt16(Data.ArmourIndex)
    Call Writer.WriteInt16(Data.HelmIndex)
    Call Writer.WriteInt16(Data.ShieldIndex)
    Call Writer.WriteInt16(Data.WeaponIndex)
    Call Writer.WriteInt16(Data.WeaponArcoIndex)
    Call Writer.WriteInt16(Data.WeaponDagaIndex)
            
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteUpdateFinishQuest_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateFinishQuest " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Mueve al personaje solo desde otro
Public Sub WriteRequiredMoveChar(ByVal UserIndex As Integer, ByVal Heading As Byte)
    
    Call Writer.WriteInt(ServerPacketID.RequiredMoveChar)
    Call Writer.WriteInt8(Heading)
    Call SendData(ToOne, UserIndex, vbNullString)
    
End Sub

Private Sub HandleStreamerBotSetting(ByVal UserIndex As Integer)
    
    Dim Delay      As Long ' Delay entre sum & sum

    Dim Mode       As eStreamerMode

    Dim DelayIndex As Long
    
    Delay = Reader.ReadInt32
    Mode = Reader.ReadInt8
    DelayIndex = Reader.ReadInt32
    
    If Not EsGm(UserIndex) Then Exit Sub
    
    If Delay < 0 Or Delay > 320000 Then Exit Sub  ' Más de un minuto no se puede poner
    If DelayIndex < 0 Or DelayIndex > 320000 Then Exit Sub  ' Más de un minuto no se puede poner
      
    ' @ Si no es seteo, comprueba de que sea él, para que otro no le modifique..
    If (Delay > 0 And Mode > 0) Then
    
        Call mStreamer.Streamer_Initial(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
        Call WriteConsoleMsg(UserIndex, "MODO STREAMER BOT", FontTypeNames.FONTTYPE_INFOGREEN)
        ' Call WriteConsoleMsg(UserIndex, "Tiempo de Warp: " & PonerPuntos(Delay), FontTypeNames.FONTTYPE_INFOGREEN)
        'Call WriteConsoleMsg(UserIndex, "Modo Seleccionado: " & Streamer_Mode_String(Mode), FontTypeNames.FONTTYPE_INFOGREEN)
        Call Streamer_CheckPosition
        'Exit Sub
    ElseIf Delay = 0 And Mode = 0 And DelayIndex = 0 Then
        Call mStreamer.Streamer_Initial(0, 0, 0, 0)
        Call WriteConsoleMsg(UserIndex, "DESACTIVADO MODO STREAMER BOT", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub
    Else

        If StreamerBot.Active <> UserIndex Then Exit Sub

    End If
    
    If Delay > 0 Then
        StreamerBot.Config_TimeWarp = Delay
        Call WriteConsoleMsg(UserIndex, "Tiempo de Warp: " & PonerPuntos(Delay), FontTypeNames.FONTTYPE_INFOGREEN)

    End If
    
    If Mode > 0 Then
        If Mode > eStreamerMode.e_LAST - 1 Then Exit Sub
        
        StreamerBot.Mode = Mode
        Call WriteConsoleMsg(UserIndex, "Modo Seleccionado: " & Streamer_Mode_String(Mode), FontTypeNames.FONTTYPE_INFOGREEN)

    End If

    If DelayIndex > 0 Then
        
        StreamerBot.Config_TimeCanIndex = DelayIndex
        Call WriteConsoleMsg(UserIndex, "Tiempo para buscar a la misma persona: " & PonerPuntos(DelayIndex), FontTypeNames.FONTTYPE_INFOGREEN)

    End If
    
End Sub

Public Function PrepareMessageUpdateBar(ByVal charindex As Integer, _
                                        ByRef Tipo As eTypeBar, _
                                        ByVal Min As Long, _
                                        ByVal max As Long) As String

    Call Writer.WriteInt(ServerPacketID.UpdateBar)
    
    Call Writer.WriteInt8(Tipo)
    Call Writer.WriteInt16(charindex)
    
    Call Writer.WriteInt32(Min)
    Call Writer.WriteInt32(max)

End Function

Public Function PrepareMessageUpdateBarTerrain(ByVal X As Integer, _
                                               ByVal Y As Integer, _
                                               ByRef Tipo As eTypeBar, _
                                               ByVal Min As Long, _
                                               ByVal max As Long) As String

    Call Writer.WriteInt(ServerPacketID.UpdateBarTerrain)
    
    Call Writer.WriteInt8(Tipo)
    Call Writer.WriteInt16(X)
    Call Writer.WriteInt16(Y)
    
    Call Writer.WriteInt32(Min)
    Call Writer.WriteInt32(max)

End Function

Private Sub HandleRequiredLive(ByVal UserIndex As Integer)
    
    If Not Interval_Message(UserIndex) Then Exit Sub

    Call mStreamer.Streamer_RequiredBOT(UserIndex)
    
End Sub

Public Sub WriteVelocidadToggle(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteVelocidadToggle_Err

    '</EhHeader>
    Call Writer.WriteInt(ServerPacketID.VelocidadToggle)
    Call Writer.WriteReal32(UserList(UserIndex).Char.speeding)
    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteVelocidadToggle_Err:
    Call Writer.Clear

    '</EhFooter>
End Sub

Public Function PrepareMessageSpeedingACT(ByVal charindex As Integer, _
                                          ByVal speeding As Single)

    '<EhHeader>
    On Error GoTo PrepareMessageSpeedingACT_Err

    '</EhHeader>
    Call Writer.WriteInt(ServerPacketID.SpeedToChar)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteReal32(speeding)
    '<EhFooter>
    Exit Function

PrepareMessageSpeedingACT_Err:
    Call Writer.Clear
       
    '</EhFooter>
End Function

Public Function PrepareMessageMeditateToggle(ByVal charindex As Integer, _
                                             ByVal FX As Integer, _
                                             Optional ByVal X As Integer = 0, _
                                             Optional ByVal Y As Integer = 0, _
                                             Optional ByVal IMeditar As Boolean = True)

    '<EhHeader>
    On Error GoTo PrepareMessageMeditateToggle_Err

    '</EhHeader>
    Call Writer.WriteInt(ServerPacketID.MeditateToggle)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt16(FX)
    Call Writer.WriteInt16(X)
    Call Writer.WriteInt16(Y)
    Call Writer.WriteBool(IMeditar)
    '<EhFooter>
    Exit Function

PrepareMessageMeditateToggle_Err:
    Call Writer.Clear

    '</EhFooter>
End Function

Private Sub HandleAcelerationChar(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleAcelerationChar_Err

    '</EhHeader>
    
    #If Classic = 1 Then
        Exit Sub
    #End If
    
    With UserList(UserIndex)
        
        If Not IntervaloPermiteShiftear(UserIndex) Then Exit Sub
        
        If Not .Stats.MinSta >= (.Stats.MaxSta * 0.3) Then Exit Sub
        .Counters.BuffoAceleration = 10
        Call ActualizarVelocidadDeUsuario(UserIndex, True)
        
        .Stats.MinSta = .Stats.MinSta - (.Stats.MaxSta * 0.3)
        Call WriteUpdateSta(UserIndex)

    End With

    '<EhFooter>
    Exit Sub

HandleAcelerationChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleAcelerationChar " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Public Sub WriteUpdateUserTrabajo(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WriteUserIndexInServer_Err

    Call Writer.WriteInt(ServerPacketID.UpdateUserTrabajo)

    Call SendData(ToOne, UserIndex, vbNullString)
    '<EhFooter>
    Exit Sub

WriteUserIndexInServer_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUserIndexInServer " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub HandleAlquilarComerciante(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleAlquilarComerciante_Err

    '</EhHeader>}
        
    Dim Tipo As Byte
        
    Tipo = Reader.ReadInt8
        
    With UserList(UserIndex)
        
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Selecciona la criatura que alquilarás!", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.eCommerceChar Then
            Call WriteChatOverHead(UserIndex, "¡Ey, yo no alquilo mi mercado!", Npclist(.flags.TargetNPC).Char.charindex, vbWhite)
            Exit Sub

        End If
            
        Exit Sub
            
        If Tipo = 1 Then
            Call mComerciantes.Commerce_SetNew(.flags.TargetNPC, UserIndex)
        ElseIf Tipo = 2 Then
            Call mComerciantes.Commerce_ViewBalance(.flags.TargetNPC, UserIndex)
        ElseIf Tipo = 3 Then
            Call mComerciantes.Commerce_ReclamarGanancias(.flags.TargetNPC, UserIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleAlquilarComerciante_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.HandleAlquilarComerciante " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Private Sub HandleTirarRuleta(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Dim Mode As Byte
        
        Mode = Reader.ReadInt8
        
        If Mode <> 1 And Mode <> 2 Then Exit Sub
        
        ' Call mRuleta.Ruleta_Tirada(UserIndex, Mode)
    
    End With

End Sub

Private Sub HandleLotteryNew(ByVal UserIndex As Integer)
    
    Dim TempLottery As tLottery
    
    TempLottery.Name = Reader.ReadString8
    TempLottery.Desc = Reader.ReadString8
    TempLottery.DateFinish = Reader.ReadString8
    TempLottery.PrizeChar = Reader.ReadString8
    TempLottery.PrizeObj = Reader.ReadInt16
    TempLottery.PrizeObjAmount = Reader.ReadInt16
    
    If Len(TempLottery.Name) <= 0 Then Exit Sub
    If Len(TempLottery.Desc) <= 0 Then Exit Sub
    If Len(TempLottery.PrizeChar) <= 0 And TempLottery.PrizeObj <= 0 Then Exit Sub
    If TempLottery.PrizeObj > 0 And TempLottery.PrizeObjAmount <= 0 Then Exit Sub

    If Not EsGmPriv(UserIndex) Then Exit Sub

    Call mLottery.Lottery_New(TempLottery)

End Sub

Public Sub WriteTournamentList(ByVal UserIndex As Integer)

    On Error GoTo WriteTournamentList_Err

    Call Writer.WriteInt(ServerPacketID.TournamentList)

    Dim A As Long, B As Long
        
    For A = 1 To MAX_EVENT_SIMULTANEO

        With Events(A)

            If .Name <> vbNullString Then
                Call Writer.WriteBool(True)
                    
                Call Writer.WriteString8(.Name)
                Call Writer.WriteInt8(.config(eConfigEvent.eFuegoAmigo))
                Call Writer.WriteInt8(.LimitRound)
                Call Writer.WriteInt8(.LimitRoundFinal)
                Call Writer.WriteInt16(.PrizePoints)
                Call Writer.WriteInt8(.LvlMin)
                Call Writer.WriteInt8(.LvlMax)
                    
                For B = 1 To NUMCLASES
                    Call Writer.WriteInt8(.AllowedClasses(B))
                Next B
                    
                Call Writer.WriteInt16(.InscriptionGld)
                Call Writer.WriteInt16(.InscriptionEldhir)
                    
                Call Writer.WriteInt16(.PrizeGld)
                Call Writer.WriteInt16(.PrizeEldhir)
                Call Writer.WriteInt16(.PrizeObj.ObjIndex)
                Call Writer.WriteInt16(.PrizeObj.Amount)
                    
                Call Writer.WriteInt8(.config(eConfigEvent.eCascoEscudo))
                    
                Call Writer.WriteInt8(.config(eConfigEvent.eResu))
                Call Writer.WriteInt8(.config(eConfigEvent.eInvisibilidad))
                Call Writer.WriteInt8(.config(eConfigEvent.eOcultar))
                Call Writer.WriteInt8(.config(eConfigEvent.eInvocar))
            Else
                Call Writer.WriteBool(False)

            End If

        End With

    Next A

    Call SendData(ToOne, UserIndex, vbNullString)

    Exit Sub

WriteTournamentList_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteTournamentList " & "at line " & Erl

End Sub

Public Sub WriteStatsUser(ByVal UserIndex As Integer, ByRef IUser As User)

    On Error GoTo WriteStatsUser_Err

    Call Writer.WriteInt(ServerPacketID.StatsUser)

    Dim A As Long
    
    With IUser
        Call Writer.WriteString8(.Name)
        Call Writer.WriteInt8(.Clase)
        Call Writer.WriteInt8(.Raza)
        Call Writer.WriteInt8(.Genero)
        Call Writer.WriteInt8(.Stats.Elv)
        Call Writer.WriteInt32(.Stats.Exp)
        Call Writer.WriteInt32(.Stats.Elu)
        
        Call Writer.WriteInt8(.Blocked)
        Call Writer.WriteInt32(.BlockedHasta)
        
        Call Writer.WriteInt32(.Stats.Gld)
        Call Writer.WriteInt32(.Stats.Eldhir)
        Call Writer.WriteInt32(.Stats.Points)
        
        Call Writer.WriteInt16(.Faction.FragsOther)
        Call Writer.WriteInt16(.Faction.FragsCiu)
        Call Writer.WriteInt16(.Faction.FragsCri)
        
        Call Writer.WriteInt16(.Pos.Map)
        Call Writer.WriteInt8(.Pos.X)
        Call Writer.WriteInt8(.Pos.Y)
        
        Call Writer.WriteInt16(.Stats.MaxHp)

    End With

    Call SendData(ToOne, UserIndex, vbNullString)

    Exit Sub

WriteStatsUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteStatsUser " & "at line " & Erl

End Sub

' # Inventario de un personaje
Public Sub WriteStatsUser_Inventory(ByVal UserIndex As Integer, ByRef IUser As Inventario)

    On Error GoTo WriteStatsUser_Inventory_Err

    Call Writer.WriteInt(ServerPacketID.StatsUser_Inventory)

    Dim A As Long
    
    With IUser
        Call Writer.WriteInt8(.NroItems)
        
        If .NroItems > 0 Then

            For A = 1 To .NroItems
                Call Writer.WriteInt16(.Object(A).ObjIndex)
                Call Writer.WriteInt16(.Object(A).Amount)
                Call Writer.WriteInt8(.Object(A).Equipped)
            Next A

        End If

    End With
    
    Call SendData(ToOne, UserIndex, vbNullString)

    Exit Sub

WriteStatsUser_Inventory_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteStatsUser_Inventory " & "at line " & Erl

End Sub

' # Banco de un personaje
Public Sub WriteStatsUser_Bank(ByVal UserIndex As Integer, ByRef IUser As BancoInventario)

    On Error GoTo WriteStatsUser_Inventory_Err

    Call Writer.WriteInt(ServerPacketID.StatsUser_Bank)

    Dim A As Long
    
    With IUser
        Call Writer.WriteInt8(.NroItems)
        
        If .NroItems > 0 Then

            For A = 1 To .NroItems
                Call Writer.WriteInt16(.Object(A).ObjIndex)
                Call Writer.WriteInt16(.Object(A).Amount)
                Call Writer.WriteInt8(.Object(A).Equipped)
            Next A

        End If

    End With
    
    Call SendData(ToOne, UserIndex, vbNullString)

    Exit Sub

WriteStatsUser_Inventory_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteStatsUser_Inventory " & "at line " & Erl

End Sub

' # Hechizos de un personaje
Public Sub WriteStatsUser_Spells(ByVal UserIndex As Integer, ByRef IUser() As Integer)

    On Error GoTo WriteStatsUser_Spells_Err

    Call Writer.WriteInt(ServerPacketID.StatsUser_Spells)

    Dim A As Long

    For A = LBound(IUser) To UBound(IUser)
        Call Writer.WriteInt16(IUser(A))
    Next A

    Call SendData(ToOne, UserIndex, vbNullString)

    Exit Sub

WriteStatsUser_Spells_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteStatsUser_Spells " & "at line " & Erl

End Sub

' # Habilidades de un personaje
Public Sub WriteStatsUser_Skills(ByVal UserIndex As Integer, ByRef IUser() As Integer)

    On Error GoTo WriteStatsUser_Skills_Err

    Call Writer.WriteInt(ServerPacketID.StatsUser_Skills)
    
    Dim A As Long

    For A = LBound(IUser) To UBound(IUser)
        Call Writer.WriteInt16(IUser(A))
    Next A
    
    Call SendData(ToOne, UserIndex, vbNullString)

    Exit Sub

WriteStatsUser_Skills_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteStatsUser_Skills " & "at line " & Erl

End Sub

' # Bonificaciones de un personaje por tiempo de duración
Public Sub WriteStatsUser_Bonus(ByVal UserIndex As Integer, ByRef IUser As UserStats)

    On Error GoTo WriteStatsUser_Bonos_Err

    Call Writer.WriteInt(ServerPacketID.StatsUser_Bonos)

    Dim A As Long
    
    With IUser
        Call Writer.WriteInt8(.BonusLast)
        
        If .BonusLast > 0 Then

            For A = 1 To .BonusLast

                With .Bonus(A)
                    Call Writer.WriteInt8(.Tipo)
                    Call Writer.WriteInt(.Value)
                    Call Writer.WriteInt(.Amount)
                    Call Writer.WriteInt(.DurationSeconds)
                    Call Writer.WriteString8(.DurationDate)

                End With

            Next A

        End If

    End With
    
    Call SendData(ToOne, UserIndex, vbNullString)

    Exit Sub

WriteStatsUser_Bonos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteStatsUser_Bonos_Err " & "at line " & Erl

End Sub

' # Penas de un personaje por tiempo de duración
Public Sub WriteStatsUser_Penas(ByVal UserIndex As Integer, ByRef IUser As User)

    On Error GoTo WriteStatsUser_Penas_Err

    Call Writer.WriteInt(ServerPacketID.StatsUser_Penas)

    Dim A As Long
    
    With IUser
        Call Writer.WriteInt16(.Counters.Pena)
        
        Call Writer.WriteInt8(.PenasLast)
        
        ' # Cargar Penas
        If .PenasLast > 0 Then

            For A = 1 To .PenasLast
                Call Writer.WriteString8(.Penas(A))
            Next A

        End If

    End With
    
    Call SendData(ToOne, UserIndex, vbNullString)

    Exit Sub

WriteStatsUser_Penas_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteStatsUser_Penas " & "at line " & Erl

End Sub

' # Skins de un personaje
Public Sub WriteStatsUser_Skins(ByVal UserIndex As Integer, ByRef IUser As tSkins)

    On Error GoTo WriteStatsUser_Skins_Err

    Call Writer.WriteInt(ServerPacketID.StatsUser_Skins)

    Dim A As Long
    
    With IUser
        Call Writer.WriteInt8(.Last)
        
        If .Last > 0 Then

            For A = 1 To .Last
                Call Writer.WriteInt16(.ObjIndex(A))
            Next A

        End If

    End With
    
    Call SendData(ToOne, UserIndex, vbNullString)

    Exit Sub

WriteStatsUser_Skins_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteStatsUser_Skins " & "at line " & Erl

End Sub

Public Sub WriteUpdateClient(ByVal UserIndex As Integer)

    On Error GoTo WriteStatsUser_Err

    Call Writer.WriteInt(ServerPacketID.UpdateClient)
    
    Call SendData(ToOne, UserIndex, vbNullString, , True)

    Exit Sub

WriteStatsUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Protocol.WriteUpdateClient " & "at line " & Erl

End Sub

Private Sub HandleCastleInfo(ByVal UserIndex As Integer)
    
    On Error GoTo ErrHandler
    
    Dim A           As Long

    Dim Text        As String
    
    Dim CastleIndex As Byte
    
    CastleIndex = Reader.ReadInt8
    
    If CastleIndex > 4 Then Exit Sub
    
    If CastleIndex = 0 Then
        If Not Interval_Message(UserIndex) Then Exit Sub
        
        For A = 1 To CastleLast

            With Castle(A)
                Text = .Name & "» " & .Desc & IIf(.GuildIndex > 0, " (Conquistado por: " & .GuildName & ")", " NO está conquistado.")
                
                Call WriteConsoleMsg(UserIndex, Text, FontTypeNames.FONTTYPE_USERBRONCE)
                
            End With

        Next A
        
        Call WriteConsoleMsg(UserIndex, "BONUS 10% EXP+ORO» " & IIf(CastleBonus > 0, "(Obtenido por: " & GuildsInfo(CastleBonus).Name & ")", "Ningún clan es poseedor."), FontTypeNames.FONTTYPE_USERPLATA)

        Call WriteConsoleMsg(UserIndex, "Utiliza los comandos /NORTE /SUR /ESTE /OESTE una vez que seas poseedor del Castillo.", FontTypeNames.FONTTYPE_USERGOLD)
    Else
        Call mCastle.Castle_Travel(UserIndex, CastleIndex)
    
    End If
    
    Exit Sub
ErrHandler:
    
End Sub

Private Sub HandleRequiredStatsUser(ByVal UserIndex As Integer)
    
    On Error GoTo ErrHandler
    
    Dim Tipo  As Byte

    Dim Name  As String

    Dim IUser As User

    Dim tUser As Integer
    
    Tipo = Reader.ReadInt8
    Name = Reader.ReadString8
    
    If Tipo < 0 Then Exit Sub
    
    ' # Chequea el intervalo con el que lo hace
    If Not Interval_Packet500(UserIndex) Then
        Call Logs_Security(eSecurity, eAntiHack, "El requerimiento de Stats está siendo alto por parte de " & UserList(UserIndex).Account.Email)
        Exit Sub

    End If
    
    ' # No existe el personaje
    If Not PersonajeExiste(Name) Then
        Call WriteConsoleMsg(UserIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
    
    ' # Está online
    tUser = NameIndex(Name)
    
    If tUser > 0 Then
        IUser = UserList(tUser)
        
        ' Información confidencial
        If MapInfo(IUser.Pos.Map).Pk Then
            IUser.Pos.Map = 0
            IUser.Invent.NroItems = 0

        End If
        
        If EsGm(tUser) Then Exit Sub
    Else

        If EsDios(Name) Or EsSemiDios(Name) Or EsAdmin(Name) Then Exit Sub
        IUser = Load_UserList_Offline(Name)         ' # Cargamos el personaje offline
        
    End If
    
    Select Case Tipo

        Case 0 ' Inventario
            Call WriteStatsUser_Inventory(UserIndex, IUser.Invent)
            
        Case 1 ' Spells
            Call WriteStatsUser_Spells(UserIndex, IUser.Stats.UserHechizos)
            
        Case 2 ' Boveda
            Call WriteStatsUser_Bank(UserIndex, IUser.BancoInvent)
            
        Case 3 ' Skills
            Call WriteStatsUser_Skills(UserIndex, IUser.Stats.UserSkills)
            
        Case 4 ' Bonus
            Call WriteStatsUser_Bonus(UserIndex, IUser.Stats)
            
        Case 5 ' Penas
            Call WriteStatsUser_Penas(UserIndex, IUser)
            
        Case 6 ' Skins
            Call WriteStatsUser_Skins(UserIndex, IUser.Skins)
            
        Case 7 ' Logros
            ' # Proximamente
            
        Case 197 ' Formulario principal
            Call WriteStatsUser(UserIndex, IUser)
            
    End Select
    
    ' # Numeros que puede solicitar
    'eInventory = 0
    'eSpells = 1
    'eBank = 2
    'eAbilities = 3
    'eBonus = 4
    'ePenas = 5
    'eSkins = 6
    'eLogros = 7
    
ErrHandler:
    Exit Sub
    
End Sub

