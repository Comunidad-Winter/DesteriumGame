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

Option Explicit

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

''
'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
'Specially usefull to create a message once and send it over to several clients.
Private auxiliarBuffer As New clsByteQueue


Private Enum ServerPacketID
    PacketGambleSv
    SendRetos
    ShortMsj
    DescNpcs
    PalabrasMagicas
    SendPartyData
    Logged                  ' LOGGED
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    MontateToggle
    CreateDamage            ' CDMG
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    CanjeInit
    InfoCanje
    UserCommerceInit        ' INITCOMUSU
    UserCommerceEnd         ' FINCOMUSUOK
    UserOfferConfirm
    CommerceChat
    ShowBlacksmithForm      ' SFH
    ShowCarpenterForm       ' SFC
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
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
    CharacterRemove         ' BP
    CharacterChangeNick
    CharacterMove           ' MP, +, * and _ '
    ForceCharMove
    CharacterChange         ' CP
    ObjectCreate            ' HO
    ObjectDelete            ' BO
    BlockPosition           ' BQ
    PlayMIDI                ' TM
    PlayWave                ' TW
    guildList               ' GL
    AreaChanged             ' CA
    PauseToggle             ' BKW
    UserInEvent
    CreateFX                ' CFX
    UpdateUserStats         ' EST
    WorkRequestTarget       ' T01
    ChangeInventorySlot     ' CSI
    ChangeBankSlot          ' SBO
    ChangeSpellSlot         ' SHS
    Atributes               ' ATR
    BlacksmithWeapons       ' LAH
    BlacksmithArmors        ' LAR
    CarpenterObjects        ' OBR
    RestOK                  ' DOK
    ErrorMsg                ' ERR
    Blind                   ' CEGU
    Dumb                    ' DUMB
    ShowSignal              ' MCAR
    ChangeNPCInventorySlot  ' NPCI
    UpdateHungerAndThirst   ' EHYS
    Fame                    ' FAMA
    MiniStats               ' MEST
    LevelUp                 ' SUNI
    AddForumMsg             ' FMSG
    ShowForumForm           ' MFOR
    SetInvisible            ' NOVER
    DiceRoll                ' DADOS
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    TrainerCreatureList     ' LSTCRI
    guildNews               ' GUILDNE
    OfferDetails            ' PEACEDE & ALLIEDE
    AlianceProposalsList    ' ALLIEPR
    PeaceProposalsList      ' PEACEPR
    CharacterInfo           ' CHRINFO
    GuildLeaderInfo         ' LEADERI
    GuildMemberInfo
    GuildDetails            ' CLANDET
    ShowGuildFundationForm  ' SHOWFUN
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    TradeOK                 ' TRANSOK
    BankOK                  ' BANCOOK
    ChangeUserTradeSlot     ' COMUSUINV
    SendNight               ' NOC
    Pong
    UpdateTagAndStatus

    MovimientSW
    rCaptions
    ShowCaptions
    rThreads
    ShowThreads

    'GM messages
    SpawnList               ' SPL
    ShowSOSForm             ' MSOS
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU

    MiniPekka
    SeeInProcess

    ShowGuildAlign
    ShowPartyForm
    UpdateStrenghtAndDexterity
    UpdateStrenght
    UpdateDexterity
    MultiMessage
    StopWorking
    CancelOfferItem
    UpdateSeguimiento
    ShowPanelSeguimiento
    EnviarDatosRanking
    QuestDetails
    QuestListSend
    FormViajes
    ApagameLaPCmono
    SvMercado
    RequestFormRostro
    ShowMenu
    EventPacketSv
End Enum

Private Enum SvEventPacketID
    SendListEvent = 1
    SendDataEvent = 2

End Enum


Private Enum ClientPacketID
    PacketGamble
    UseItemPacket
    RequestPositionUpdate   'RPU
    PickUp                  'AG
    Lookprocess
    RequestFame             'FAMA
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    UserCommerceConfirm

    RequestSkills           'ESKI

    CommerceChat
    PacketRetos
    CanjeItem

    ThrowDices
    Talk                    ';
    Yell                    '-
    ReportCheat
    Whisper                 '\
    Walk                    'M

    SendProcessList
    CombatModeToggle        'TAB        - SHOULD BE HANLDED JUST BY THE CLIENT!!
    SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle
    RequestGuildLeaderInfo  'GLINFO
    RequestAtributes        'ATR

    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    Work
    LogeaNuevoPj
    CraftBlacksmith         'CNS
    CraftCarpenter          'CNC
    CanjeInfo
    ChangeNick
    WorkLeftClick           'WLC
    CreateNewGuild
    GuildOfferPeace         'PEACEOFF
    GuildOfferAlliance      'ALLIEOFF
    GuildAllianceDetails    'ALLIEDET
    GuildPeaceDetails       'PEACEDET
    GuildRequestJoinerInfo  'ENVCOMEN
    Useitem
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    ModifySkills            'SKSE
    Train                   'ENTR
    Attack
    CommerceBuy
    BankExtractItem
    ClanCodexUpdate         'DESCOD
    UserCommerceOffer       'OFRECER
    GuildAcceptPeace        'ACEPPEAT
    GuildRejectAlliance     'RECPALIA
    GuildRejectPeace        'RECPPEAT
    GuildAcceptAlliance     'ACEPALIA

    GuildAlliancePropList   'ENVALPRO
    GuildPeacePropList      'ENVPROPP
    GuildDeclareWar         'DECGUERR
    GuildLeave              '/SALIRCLAN
    
    GuildNewWebsite         'NEWWEBSI
    CommerceSell
    PaqueteEncriptado
    BankDeposit
    ForumPost
    MoveSpell
    MoveBank

    GuildAcceptNewMember    'ACEPTARI
    Drop                    'T
    DoubleClick
    Meditate                '/MEDITAR
    GuildRejectNewMember    'RECHAZAR

    GuildOpenElections      'ABREELEC
    GuildRequestMembership  'SOLICITUD
    GuildRequestDetails     'CLANDETAILS
    TrainList               '/ENTRENAR
    Rest                    '/DESCANSAR
    CastSpell               'LH
    Online                  '/ONLINE
    Quit                    '/SALIR
    LeftClick               'LC

    RequestAccountState     '/BALANCE            'RC
    RequestInfoEvento
    PetStand                '/QUIETO
    UseSpellMacro           'UMH
    PetFollow               '/ACOMPAÑAR
    ReleasePet              '/LIBERAR
    Oro
    Plata
    Bronce
    Limpiar
    GlobalMessage
    GMCommands
    GlobalStatus
    CuentaRegresiva
    Nivel
    ResetearPj
    BorrarPJ
    RecuperarPJ
    Verpenas
    DropItems
    Fianzah
    GuildKickMember         'ECHARCLA
    GuildUpdateNews         'ACTGNEWS
    GuildMemberInfo         '1HRINFO<
    Resucitate
    Heal
    Help
    RequestStats
    CommerceStart
    BankStart
    Enlist
    Information
    Reward
    UpTime
    PartyLeave
    PartyCreate
    PartyJoin
    Inquiry
    GuildMessage
    PartyMessage
    CentinelReport
    GuildOnline
    PartyOnline
    CouncilMessage
    RoleMasterRequest
    GMRequest
    ChangeDescription
    GuildVote
    Punishments
    ChangePassword
    ChangePin
    rCaptions
    SCaptions
    InquiryVote
    LeaveFaction
    BankExtractGold
    BankDepositGold
    Denounce
    rThreads
    SThreads
    Gamble
    GuildFundate
    GuildFundation
    PartyKick
    PartySetLeader
    PartyAcceptMember
    Ping
    Cara
    Viajar
    ItemUpgrade
    InitCrafting
    Home
    ShowGuildNews
    ConectarUsuarioE
    ShareNpc
    StopSharingNpc
    Consulta
    SolicitaRranking
    solicitudes
    WherePower
    Premium
    Mercado
    RightClick
    EventPacket
    GuildDisolution
    Quest
    QuestAccept
    QuestListRequest
    QuestDetailsRequest
    QuestAbandon

    DragToPos
    DragToggle
    UsePotions
    SetMenu
    dragInventory
    SetPartyPorcentajes
    RequestPartyForm
    PpARTY
    usarbono

End Enum


Public Enum EventPacketID
    eNewEvent = 1
    eCloseEvent = 2
    RequiredEvents = 3
    RequiredDataEvent = 4
    eParticipeEvent = 5
    eAbandonateEvent = 6
End Enum

Private Enum MercadoPacketID
    RequestMercado = 1
    SendMercado = 2
    RequestTipoMAO = 3
    SendTipoMAO = 4
    RequestInfoCharMAO = 5
    PublicationPj = 6
    InvitationChange = 7
    BuyPj = 8
    QuitarPj = 9
    RequestOfferSent = 10
    SendOfferSent = 11
    RequestOffer = 12
    SendOffer = 13
    AcceptInvitation = 14
    RechaceInvitation = 15
    CancelInvitation = 16
End Enum

''
'The last existing client packet id.
Private Const LAST_CLIENT_PACKET_ID As Byte = 190

Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
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
    fonttype_dios
    FONTTYPE_CONTEO
    FONTTYPE_CONTEOS
    FONTTYPE_ADMIN
    FONTTYPE_PREMIUM
    FONTTYPE_RETO
    FONTTYPE_ORO
    FONTTYPE_PLATA
    FONTTYPE_BRONCE
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

Private Function CheckCRC(ByVal Userindex As Integer, ByVal CRC As Integer) As Boolean
'***************************************************
'Author: Damián
'Last Modification: 24/08/2013
'***************************************************
    
    With UserList(Userindex)
        .CRC = .CRC + 1
        If .CRC = 501 Then .CRC = 0
    
        If SeguridadCRC(.CRC) <> CRC Then
            Call LogCRC("El usuario " & .Name & " de IP: " & .ip & " mandó un CRC equivocado.")
            CheckCRC = False
            Exit Function
        End If
        
        CheckCRC = True
    End With
End Function


''
' Handles incoming data.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIncomingData(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'
'***************************************************
    On Error Resume Next
    Dim PacketID As Byte

    PacketID = UserList(Userindex).incomingData.PeekByte()

    'Does the packet requires a logged user??
    If Not (PacketID = ClientPacketID.ThrowDices _
            Or PacketID = ClientPacketID.ConectarUsuarioE _
            Or PacketID = ClientPacketID.LogeaNuevoPj _
            Or PacketID = ClientPacketID.BorrarPJ _
            Or PacketID = ClientPacketID.RecuperarPJ _
            Or PacketID = ClientPacketID.PaqueteEncriptado) Then


        'Is the user actually logged?
        If Not UserList(Userindex).flags.UserLogged Then
            Call CloseSocket(Userindex)
            Exit Sub

            'He is logged. Reset idle counter if id is valid.
        ElseIf PacketID <= LAST_CLIENT_PACKET_ID Then
            UserList(Userindex).Counters.IdleCount = 0
        End If
    ElseIf PacketID <= LAST_CLIENT_PACKET_ID Then
        UserList(Userindex).Counters.IdleCount = 0

        'Is the user logged?
        If UserList(Userindex).flags.UserLogged Then
            Call CloseSocket(Userindex)
            Exit Sub
        End If
    End If

    ' Ante cualquier paquete, pierde la proteccion de ser atacado.
    UserList(Userindex).flags.NoPuedeSerAtacado = False


    Select Case PacketID
    
    Case ClientPacketID.PacketGamble
        Call HandlePacketGamble(Userindex)
    
    Case ClientPacketID.RequestInfoEvento
        Call HandleRequestInfoEvento(Userindex)
        
        
    Case ClientPacketID.PacketRetos
        Call HandlePacketRetos(Userindex)
    
    Case ClientPacketID.CanjeItem
        Call HandleCanjeItem(Userindex)
        
    Case ClientPacketID.CanjeInfo
        Call HandleCanjeInfo(Userindex)
        
    Case ClientPacketID.ChangeNick
        Call HandleChangeNick(Userindex)
        
    Case ClientPacketID.ReportCheat
        Call HandleReportCheat(Userindex)
        
   ' Case ClientPacketID.PacketCofres
      '  Call HandleCofres(UserIndex)
        
    Case ClientPacketID.PaqueteEncriptado
        Call HandlePaqueteEncriptado(Userindex)
        
    Case ClientPacketID.GuildDisolution
        Call HandleDisolutionGuild(Userindex)
        
    Case ClientPacketID.EventPacket
        Call HandleEventPacket(Userindex)

    Case ClientPacketID.ThrowDices              'TIRDAD
        Call HandleThrowDices(Userindex)

    Case ClientPacketID.LogeaNuevoPj            'NLOGIN
        Call HandleLogeaNuevoPj(Userindex)

    Case ClientPacketID.BorrarPJ                'BORROK
        Call HandleKillChar(Userindex)

    Case ClientPacketID.RecuperarPJ             'RECPAS
        Call HandleRenewPassChar(Userindex)

    Case ClientPacketID.Talk                    ';
        Call HandleTalk(Userindex)

    Case ClientPacketID.Yell                    '-
        Call HandleYell(Userindex)

    Case ClientPacketID.Whisper                 '\
        Call HandleWhisper(Userindex)

    Case ClientPacketID.Walk                    'M
        Call HandleWalk(Userindex)

    Case ClientPacketID.Lookprocess
        Call HandleLookProcess(Userindex)
    Case ClientPacketID.SendProcessList
        Call HandleSendProcessList(Userindex)


    Case ClientPacketID.RequestPositionUpdate   'RPU
        Call HandleRequestPositionUpdate(Userindex)
        
    Case ClientPacketID.UseItemPacket
        Call HandleUseItemPacket(Userindex)

    Case ClientPacketID.Attack                  'AT
        Call HandleAttack(Userindex)

    Case ClientPacketID.PickUp                  'AG
        Call HandlePickUp(Userindex)

    Case ClientPacketID.CombatModeToggle        'TAB        - SHOULD BE HANLDED JUST BY THE CLIENT!!
        Call HanldeCombatModeToggle(Userindex)

    Case ClientPacketID.SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
        Call HandleSafeToggle(Userindex)

    Case ClientPacketID.ResuscitationSafeToggle
        Call HandleResuscitationToggle(Userindex)

    Case ClientPacketID.RequestGuildLeaderInfo  'GLINFO
        Call HandleRequestGuildLeaderInfo(Userindex)

    Case ClientPacketID.RequestAtributes        'ATR
        Call HandleRequestAtributes(Userindex)

    Case ClientPacketID.RequestFame             'FAMA
        Call HandleRequestFame(Userindex)

    Case ClientPacketID.RequestSkills           'ESKI
        Call HandleRequestSkills(Userindex)

    Case ClientPacketID.RequestMiniStats        'FEST
        Call HandleRequestMiniStats(Userindex)

    Case ClientPacketID.CommerceEnd             'FINCOM
        Call HandleCommerceEnd(Userindex)

    Case ClientPacketID.CommerceChat
        Call HandleCommerceChat(Userindex)

    Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
        Call HandleUserCommerceEnd(Userindex)

    Case ClientPacketID.UserCommerceConfirm
        Call HandleUserCommerceConfirm(Userindex)

    Case ClientPacketID.BankEnd                 'FINBAN
        Call HandleBankEnd(Userindex)

    Case ClientPacketID.UserCommerceOk          'COMUSUOK
        Call HandleUserCommerceOk(Userindex)

    Case ClientPacketID.UserCommerceReject      'COMUSUNO
        Call HandleUserCommerceReject(Userindex)

    Case ClientPacketID.Drop                    'TI
        Call HandleDrop(Userindex)

    Case ClientPacketID.CastSpell               'LH
        Call HandleCastSpell(Userindex)

    Case ClientPacketID.LeftClick               'LC
        Call HandleLeftClick(Userindex)

    Case ClientPacketID.DoubleClick             'RC
        Call HandleDoubleClick(Userindex)

    Case ClientPacketID.Work                    'UK
        Call HandleWork(Userindex)

    Case ClientPacketID.UseSpellMacro           'UMH
        Call HandleUseSpellMacro(Userindex)


    Case ClientPacketID.ConectarUsuarioE       'OLOGIN
        Call HandleConectarUsuarioE(Userindex)

    Case ClientPacketID.Useitem                 'USA
        Call HandleUseItem(Userindex)

    Case ClientPacketID.CraftBlacksmith         'CNS
        Call HandleCraftBlacksmith(Userindex)

    Case ClientPacketID.CraftCarpenter          'CNC
        Call HandleCraftCarpenter(Userindex)

    Case ClientPacketID.WorkLeftClick           'WLC
        Call HandleWorkLeftClick(Userindex)

    Case ClientPacketID.CreateNewGuild          'CIG
        Call HandleCreateNewGuild(Userindex)

    Case ClientPacketID.SpellInfo               'INFS
        Call HandleSpellInfo(Userindex)

    Case ClientPacketID.EquipItem               'EQUI
        Call HandleEquipItem(Userindex)

    Case ClientPacketID.ChangeHeading           'CHEA
        Call HandleChangeHeading(Userindex)

    Case ClientPacketID.ModifySkills            'SKSE
        Call HandleModifySkills(Userindex)

    Case ClientPacketID.Train                   'ENTR
        Call HandleTrain(Userindex)

    Case ClientPacketID.CommerceBuy             'COMP
        Call HandleCommerceBuy(Userindex)

    Case ClientPacketID.BankExtractItem         'RETI
        Call HandleBankExtractItem(Userindex)

    Case ClientPacketID.CommerceSell            'VEND
        Call HandleCommerceSell(Userindex)

    Case ClientPacketID.BankDeposit             'DEPO
        Call HandleBankDeposit(Userindex)

    Case ClientPacketID.ForumPost               'DEMSG
        Call HandleForumPost(Userindex)

    Case ClientPacketID.MoveSpell               'DESPHE
        Call HandleMoveSpell(Userindex)

    Case ClientPacketID.MoveBank
        Call HandleMoveBank(Userindex)

    Case ClientPacketID.ClanCodexUpdate         'DESCOD
        Call HandleClanCodexUpdate(Userindex)

    Case ClientPacketID.UserCommerceOffer       'OFRECER
        Call HandleUserCommerceOffer(Userindex)

    Case ClientPacketID.GuildAcceptPeace        'ACEPPEAT
        Call HandleGuildAcceptPeace(Userindex)

    Case ClientPacketID.GuildRejectAlliance     'RECPALIA
        Call HandleGuildRejectAlliance(Userindex)

    Case ClientPacketID.GuildRejectPeace        'RECPPEAT
        Call HandleGuildRejectPeace(Userindex)

    Case ClientPacketID.GuildAcceptAlliance     'ACEPALIA
        Call HandleGuildAcceptAlliance(Userindex)

    Case ClientPacketID.GuildOfferPeace         'PEACEOFF
        Call HandleGuildOfferPeace(Userindex)

    Case ClientPacketID.GuildOfferAlliance      'ALLIEOFF
        Call HandleGuildOfferAlliance(Userindex)

    Case ClientPacketID.GuildAllianceDetails    'ALLIEDET
        Call HandleGuildAllianceDetails(Userindex)

    Case ClientPacketID.GuildPeaceDetails       'PEACEDET
        Call HandleGuildPeaceDetails(Userindex)

    Case ClientPacketID.GuildRequestJoinerInfo  'ENVCOMEN
        Call HandleGuildRequestJoinerInfo(Userindex)

    Case ClientPacketID.GuildAlliancePropList   'ENVALPRO
        Call HandleGuildAlliancePropList(Userindex)

    Case ClientPacketID.GuildPeacePropList      'ENVPROPP
        Call HandleGuildPeacePropList(Userindex)

    Case ClientPacketID.GuildDeclareWar         'DECGUERR
        Call HandleGuildDeclareWar(Userindex)

    Case ClientPacketID.GuildNewWebsite         'NEWWEBSI
        Call HandleGuildNewWebsite(Userindex)

    Case ClientPacketID.GuildAcceptNewMember    'ACEPTARI
        Call HandleGuildAcceptNewMember(Userindex)

    Case ClientPacketID.GuildRejectNewMember    'RECHAZAR
        Call HandleGuildRejectNewMember(Userindex)

    Case ClientPacketID.GuildKickMember         'ECHARCLA
        Call HandleGuildKickMember(Userindex)

    Case ClientPacketID.GuildUpdateNews         'ACTGNEWS
        Call HandleGuildUpdateNews(Userindex)

    Case ClientPacketID.GuildMemberInfo         '1HRINFO<
        Call HandleGuildMemberInfo(Userindex)

    Case ClientPacketID.GuildOpenElections      'ABREELEC
        Call HandleGuildOpenElections(Userindex)

    Case ClientPacketID.GuildRequestMembership  'SOLICITUD
        Call HandleGuildRequestMembership(Userindex)

    Case ClientPacketID.GuildRequestDetails     'CLANDETAILS
        Call HandleGuildRequestDetails(Userindex)

    Case ClientPacketID.Online                  '/ONLINE
        Call HandleOnline(Userindex)

    Case ClientPacketID.Quit                    '/SALIR
        Call HandleQuit(Userindex)

    Case ClientPacketID.GuildLeave              '/SALIRCLAN
        Call HandleGuildLeave(Userindex)


    Case ClientPacketID.RequestAccountState     '/BALANCE
        Call HandleRequestAccountState(Userindex)

    Case ClientPacketID.PetStand                '/QUIETO
        Call HandlePetStand(Userindex)

    Case ClientPacketID.PetFollow               '/ACOMPAÑAR
        Call HandlePetFollow(Userindex)

    Case ClientPacketID.ReleasePet              '/LIBERAR
        Call HandleReleasePet(Userindex)

    Case ClientPacketID.TrainList               '/ENTRENAR
        Call HandleTrainList(Userindex)

    Case ClientPacketID.Rest                    '/DESCANSAR
        Call HandleRest(Userindex)

    Case ClientPacketID.Meditate                '/MEDITAR
        Call HandleMeditate(Userindex)

    Case ClientPacketID.Verpenas
        Call Handleverpenas(Userindex)

    Case ClientPacketID.DropItems           '/CAER
        Call HandleDropItems(Userindex)

    Case ClientPacketID.Fianzah
        Call HandleFianzah(Userindex)

    Case ClientPacketID.Resucitate              '/RESUCITAR
        Call HandleResucitate(Userindex)

    Case ClientPacketID.Heal                    '/CURAR
        Call HandleHeal(Userindex)

    Case ClientPacketID.Help                    '/AYUDA
        Call HandleHelp(Userindex)

    Case ClientPacketID.RequestStats            '/EST
        Call HandleRequestStats(Userindex)

    Case ClientPacketID.CommerceStart           '/COMERCIAR
        Call HandleCommerceStart(Userindex)

    Case ClientPacketID.BankStart               '/BOVEDA
        Call HandleBankStart(Userindex)

    Case ClientPacketID.Enlist                  '/ENLISTAR
        Call HandleEnlist(Userindex)

    Case ClientPacketID.Information             '/INFORMACION
        Call HandleInformation(Userindex)

    Case ClientPacketID.Reward                  '/RECOMPENSA
        Call HandleReward(Userindex)

    Case ClientPacketID.UpTime                  '/UPTIME
        Call HandleUpTime(Userindex)

    Case ClientPacketID.PartyLeave              '/SALIRPARTY
        Call HandlePartyLeave(Userindex)

    Case ClientPacketID.PartyCreate             '/CREARPARTY
        Call HandlePartyCreate(Userindex)

    Case ClientPacketID.PartyJoin               '/PARTY
        Call HandlePartyJoin(Userindex)

    Case ClientPacketID.Inquiry                 '/ENCUESTA ( with no params )
        Call HandleInquiry(Userindex)

    Case ClientPacketID.GuildMessage            '/CMSG
        Call HandleGuildMessage(Userindex)

    Case ClientPacketID.PartyMessage            '/PMSG
        Call HandlePartyMessage(Userindex)

    Case ClientPacketID.CentinelReport          '/CENTINELA
        Call HandleCentinelReport(Userindex)

    Case ClientPacketID.GuildOnline             '/ONLINECLAN
        Call HandleGuildOnline(Userindex)

    Case ClientPacketID.PartyOnline             '/ONLINEPARTY
        Call HandlePartyOnline(Userindex)

    Case ClientPacketID.CouncilMessage          '/BMSG
        Call HandleCouncilMessage(Userindex)

    Case ClientPacketID.RoleMasterRequest       '/ROL
        Call HandleRoleMasterRequest(Userindex)

    Case ClientPacketID.GMRequest               '/GM
        Call HandleGMRequest(Userindex)

    Case ClientPacketID.ChangeDescription       '/DESC
        Call HandleChangeDescription(Userindex)

    Case ClientPacketID.GuildVote               '/VOTO
        Call HandleGuildVote(Userindex)

    Case ClientPacketID.Punishments             '/PENAS
        Call HandlePunishments(Userindex)

    Case ClientPacketID.ChangePassword          '/CONTRASEÑA
        Call HandleChangePassword(Userindex)

    Case ClientPacketID.ChangePin         '/CONTRASEÑA
        Call HandleChangePin(Userindex)

    Case ClientPacketID.Gamble                  '/APOSTAR
        Call HandleGamble(Userindex)

    Case ClientPacketID.InquiryVote             '/ENCUESTA ( with parameters )
        Call HandleInquiryVote(Userindex)

    Case ClientPacketID.LeaveFaction            '/RETIRAR ( with no arguments )
        Call HandleLeaveFaction(Userindex)

    Case ClientPacketID.BankExtractGold         '/RETIRAR ( with arguments )
        Call HandleBankExtractGold(Userindex)

    Case ClientPacketID.BankDepositGold         '/DEPOSITAR
        Call HandleBankDepositGold(Userindex)

    Case ClientPacketID.Denounce                '/DENUNCIAR
        Call HandleDenounce(Userindex)

    Case ClientPacketID.GuildFundate            '/FUNDARCLAN
        Call HandleGuildFundate(Userindex)

    Case ClientPacketID.GuildFundation
        Call HandleGuildFundation(Userindex)

    Case ClientPacketID.PartyKick               '/ECHARPARTY
        Call HandlePartyKick(Userindex)

    Case ClientPacketID.PartySetLeader          '/PARTYLIDER
        Call HandlePartySetLeader(Userindex)

    Case ClientPacketID.PartyAcceptMember       '/ACCEPTPARTY
        Call HandlePartyAcceptMember(Userindex)

    Case ClientPacketID.rCaptions
        Call HandleRequieredCaptions(Userindex)

    Case ClientPacketID.SCaptions
        Call HandleSendCaptions(Userindex)

    Case ClientPacketID.Ping                    '/PING
        Call HandlePing(Userindex)

    Case ClientPacketID.Cara                    '/Cara
        Call HandleCara(Userindex)

    Case ClientPacketID.Viajar
        Call HandleViajar(Userindex)

    Case ClientPacketID.ItemUpgrade
        Call HandleItemUpgrade(Userindex)

    Case ClientPacketID.GMCommands              'GM Messages
        Call HandleGMCommands(Userindex)

    Case ClientPacketID.InitCrafting
        Call HandleInitCrafting(Userindex)

    Case ClientPacketID.Home
        Call HandleHome(Userindex)

    Case ClientPacketID.ShowGuildNews
        Call HandleShowGuildNews(Userindex)

    Case ClientPacketID.ShareNpc
        Call HandleShareNpc(Userindex)

    Case ClientPacketID.StopSharingNpc
        Call HandleStopSharingNpc(Userindex)

    Case ClientPacketID.Consulta
        Call HandleConsultation(Userindex)

    Case ClientPacketID.SolicitaRranking
        Call HandleSolicitarRanking(Userindex)

    Case ClientPacketID.Quest                   '/QUEST
        Call HandleQuest(Userindex)

    Case ClientPacketID.QuestAccept
        Call HandleQuestAccept(Userindex)

    Case ClientPacketID.QuestListRequest
        Call HandleQuestListRequest(Userindex)

    Case ClientPacketID.QuestDetailsRequest
        Call HandleQuestDetailsRequest(Userindex)

    Case ClientPacketID.QuestAbandon
        Call HandleQuestAbandon(Userindex)

    Case ClientPacketID.ResetearPj                 '/RESET
        Call HandleResetearPJ(Userindex)

    Case ClientPacketID.Nivel               '/RESET
        Call HanDlenivel(Userindex)


    Case ClientPacketID.usarbono
        Call HandleUsarBono(Userindex)

    Case ClientPacketID.Oro
        Call HandleOro(Userindex)
        
    Case ClientPacketID.Premium
        Call HandlePremium(Userindex)
        
    Case ClientPacketID.Mercado
        Call HandleMercado(Userindex)
        
    Case ClientPacketID.RightClick
        Call HandleRightClick(Userindex)

    Case ClientPacketID.Plata
        Call HandlePlata(Userindex)

    Case ClientPacketID.Bronce
        Call HandleBronce(Userindex)

    Case ClientPacketID.GlobalMessage
        Call HandleGlobalMessage(Userindex)

    Case ClientPacketID.GlobalStatus
        Call HandleGlobalStatus(Userindex)

    Case ClientPacketID.CuentaRegresiva      '/CR
        Call HandleCuentaRegresiva(Userindex)

    Case ClientPacketID.dragInventory        'DINVENT
        Call HandleDragInventory(Userindex)

    Case ClientPacketID.DragToPos               'DTOPOS
        Call HandleDragToPos(Userindex)

    Case ClientPacketID.DragToggle
        Call HandleDragToggle(Userindex)

    Case ClientPacketID.SetPartyPorcentajes
        Call handleSetPartyPorcentajes(Userindex)

    Case ClientPacketID.RequestPartyForm                  '205
        Call handleRequestPartyForm(Userindex)

    Case ClientPacketID.solicitudes              '/DENUNCIAR
        Call HandleSolicitud(Userindex)

    Case ClientPacketID.UsePotions
        Call HandleUsePotions(Userindex)

    Case ClientPacketID.SetMenu
        Call HandleSetMenu(Userindex)
    
    Case ClientPacketID.WherePower
        Call HandleWherePower(Userindex)


        #If SeguridadAlkon Then
        Case Else
            Do While HandleIncomingDataEx(Userindex)
            Loop
        #Else
        Case Else
            'ERROR : Abort!
            Call CloseSocket(Userindex)
        #End If
    End Select

    'Done with this packet, move on to next one or send everything if no more packets found
    If UserList(Userindex).incomingData.length > 0 And Err.Number = 0 Then
        Err.Clear
        Call HandleIncomingData(Userindex)
    
    ElseIf Err.Number <> 0 And Not Err.Number = UserList(Userindex).incomingData.NotEnoughDataErrCode Then
        'An error ocurred, log it and kick player.
        Call LogError("Error: " & Err.Number & " [" & Err.Description & "] " & " Source: " & Err.source & _
                        vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & _
                        vbTab & " LastDllError: " & Err.LastDllError & vbTab & _
                        " - UserIndex: " & Userindex & " - producido al manejar el paquete: " & CStr(PacketID))
        Call CloseSocket(Userindex)
    
    Else
        'Flush buffer - send everything that has been written
        Call FlushBuffer(Userindex)
    End If
End Sub

Public Sub WriteMultiMessage(ByVal Userindex As Integer, ByVal MessageIndex As Integer, Optional ByVal Arg1 As Long, Optional ByVal Arg2 As Long, Optional ByVal Arg3 As Long, Optional ByVal StringArg1 As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.MultiMessage)
        Call .WriteByte(MessageIndex)

        Select Case MessageIndex
        Case eMessages.DontSeeAnything, eMessages.NPCSwing, eMessages.NPCKillUser, eMessages.BlockedWithShieldUser, _
             eMessages.BlockedWithShieldother, eMessages.UserSwing, eMessages.SafeModeOn, eMessages.SafeModeOff, eMessages.DragOnn, eMessages.DragOff, _
             eMessages.ResuscitationSafeOff, eMessages.ResuscitationSafeOn, eMessages.NobilityLost, _
             eMessages.CantUseWhileMeditating, eMessages.CancelHome, eMessages.FinishHome

        Case eMessages.NPCHitUser
            Call .WriteByte(Arg1)    'Target
            Call .WriteInteger(Arg2)    'damage

        Case eMessages.UserHitNPC
            Call .WriteLong(Arg1)    'damage

        Case eMessages.UserAttackedSwing
            Call .WriteInteger(UserList(Arg1).Char.CharIndex)

        Case eMessages.UserHittedByUser
            Call .WriteInteger(Arg1)    'AttackerIndex
            Call .WriteByte(Arg2)    'Target
            Call .WriteInteger(Arg3)    'damage

        Case eMessages.UserHittedUser
            Call .WriteInteger(Arg1)    'AttackerIndex
            Call .WriteByte(Arg2)    'Target
            Call .WriteInteger(Arg3)    'damage

        Case eMessages.WorkRequestTarget
            Call .WriteByte(Arg1)    'skill

        Case eMessages.HaveKilledUser    '"Has matado a " & UserList(VictimIndex).name & "!" "Has ganado " & DaExp & " puntos de experiencia."
            Call .WriteInteger(UserList(Arg1).Char.CharIndex)    'VictimIndex
            Call .WriteLong(Arg2)    'Expe

        Case eMessages.UserKill    '"¡" & .name & " te ha matado!"
            Call .WriteInteger(UserList(Arg1).Char.CharIndex)    'AttackerIndex

        Case eMessages.EarnExp

        Case eMessages.Home
            Call .WriteByte(CByte(Arg1))
            Call .WriteInteger(CInt(Arg2))
            'El cliente no conoce nada sobre nombre de mapas y hogares, por lo tanto _
             hasta que no se pasen los dats e .INFs al cliente, esto queda así.
            Call .WriteASCIIString(StringArg1)    'Call .WriteByte(CByte(Arg2))

        End Select
    End With
    Exit Sub    ''

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
Private Sub HandleGMCommands(ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo Errhandler

    Dim Command As Byte

    With UserList(Userindex)
        Call .incomingData.ReadByte
        
        Command = .incomingData.PeekByte

        If Not EsGM(Userindex) Then
            
            FlushBuffer Userindex
        End If
        
        Select Case Command
        
        Case eGMCommands.GMMessage                '/GMSG
            Call HandleGMMessage(Userindex)

        Case eGMCommands.showName                '/SHOWNAME
            Call HandleShowName(Userindex)

        Case eGMCommands.OnlineRoyalArmy
            Call HandleOnlineRoyalArmy(Userindex)

        Case eGMCommands.OnlineChaosLegion       '/ONLINECAOS
            Call HandleOnlineChaosLegion(Userindex)

        Case eGMCommands.GoNearby                '/IRCERCA
            Call HandleGoNearby(Userindex)

        Case eGMCommands.SeBusca                '/SEBUSCA
            Call Elmasbuscado(Userindex)

        Case eGMCommands.comment                 '/REM
            Call HandleComment(Userindex)

        Case eGMCommands.serverTime              '/HORA
            Call HandleServerTime(Userindex)

        Case eGMCommands.Where                   '/DONDE
            Call HandleWhere(Userindex)

        Case eGMCommands.CreaturesInMap          '/NENE
            Call HandleCreaturesInMap(Userindex)

        Case eGMCommands.WarpMeToTarget          '/TELEPLOC
            Call HandleWarpMeToTarget(Userindex)

        Case eGMCommands.WarpChar                '/TELEP
            Call HandleWarpChar(Userindex)

        Case eGMCommands.Silence                 '/SILENCIAR
            Call HandleSilence(Userindex)

        Case eGMCommands.SOSShowList             '/SHOW SOS
            Call HandleSOSShowList(Userindex)

        Case eGMCommands.SOSRemove               'SOSDONE
            Call HandleSOSRemove(Userindex)

        Case eGMCommands.GoToChar                '/IRA
            Call HandleGoToChar(Userindex)

        Case eGMCommands.invisible               '/INVISIBLE
            Call HandleInvisible(Userindex)

        Case eGMCommands.GMPanel                 '/PANELGM
            Call HandleGMPanel(Userindex)

        Case eGMCommands.RequestUserList         'LISTUSU
            Call HandleRequestUserList(Userindex)

        Case eGMCommands.Working                 '/TRABAJANDO
            Call HandleWorking(Userindex)

        Case eGMCommands.Hiding                  '/OCULTANDO
            Call HandleHiding(Userindex)

        Case eGMCommands.Jail                    '/CARCEL
            Call HandleJail(Userindex)

        Case eGMCommands.KillNPC                 '/RMATA
            Call HandleKillNPC(Userindex)

        Case eGMCommands.WarnUser                '/ADVERTENCIA
            Call HandleWarnUser(Userindex)

        Case eGMCommands.RequestCharInfo         '/INFO
            Call HandleRequestCharInfo(Userindex)

        Case eGMCommands.RequestCharStats        '/STAT
            Call HandleRequestCharStats(Userindex)

        Case eGMCommands.RequestCharGold         '/BAL
            Call HandleRequestCharGold(Userindex)

        Case eGMCommands.RequestCharInventory    '/INV
            Call HandleRequestCharInventory(Userindex)

        Case eGMCommands.RequestCharBank         '/BOV
            Call HandleRequestCharBank(Userindex)

        Case eGMCommands.RequestCharSkills       '/SKILLS
            Call HandleRequestCharSkills(Userindex)

        Case eGMCommands.ReviveChar              '/REVIVIR
            Call HandleReviveChar(Userindex)

        Case eGMCommands.OnlineGM                '/ONLINEGM
            Call HandleOnlineGM(Userindex)

        Case eGMCommands.OnlineMap               '/ONLINEMAP
            Call HandleOnlineMap(Userindex)

        Case eGMCommands.Forgive                 '/PERDON
            Call HandleForgive(Userindex)

        Case eGMCommands.Kick                    '/ECHAR
            Call HandleKick(Userindex)

        Case eGMCommands.Execute                 '/EJECUTAR
            Call HandleExecute(Userindex)

        Case eGMCommands.banChar                 '/BAN
            Call HandleBanChar(Userindex)

        Case eGMCommands.UnbanChar               '/UNBAN
            Call HandleUnbanChar(Userindex)

        Case eGMCommands.NPCFollow               '/SEGUIR
            Call HandleNPCFollow(Userindex)

        Case eGMCommands.SummonChar              '/SUM
            Call HandleSummonChar(Userindex)

        Case eGMCommands.SpawnListRequest        '/CC
            Call HandleSpawnListRequest(Userindex)

        Case eGMCommands.SpawnCreature           'SPA
            Call HandleSpawnCreature(Userindex)

        Case eGMCommands.ResetNPCInventory       '/RESETINV
            Call HandleResetNPCInventory(Userindex)

        Case eGMCommands.cleanworld              '/LIMPIAR
            Call HandleCleanWorld(Userindex)

        Case eGMCommands.ServerMessage           '/RMSG
            Call HandleServerMessage(Userindex)

        Case eGMCommands.RolMensaje           '/ROLEANDO
            Call HandleRolMensaje(Userindex)

        Case eGMCommands.nickToIP                '/NICK2IP
            Call HandleNickToIP(Userindex)

        Case eGMCommands.IPToNick                '/IP2NICK
            Call HandleIPToNick(Userindex)

        Case eGMCommands.GuildOnlineMembers      '/ONCLAN
            Call HandleGuildOnlineMembers(Userindex)

        Case eGMCommands.TeleportCreate          '/CT
            Call HandleTeleportCreate(Userindex)

        Case eGMCommands.TeleportDestroy         '/DT
            Call HandleTeleportDestroy(Userindex)

        Case eGMCommands.RainToggle              '/LLUVIA
            Call HandleRainToggle(Userindex)

        Case eGMCommands.SetCharDescription      '/SETDESC
            Call HandleSetCharDescription(Userindex)

        Case eGMCommands.ForceMIDIToMap          '/FORCEMIDIMAP
            Call HanldeForceMIDIToMap(Userindex)

        Case eGMCommands.ForceWAVEToMap          '/FORCEWAVMAP
            Call HandleForceWAVEToMap(Userindex)

        Case eGMCommands.RoyalArmyMessage        '/REALMSG
            Call HandleRoyalArmyMessage(Userindex)

        Case eGMCommands.ChaosLegionMessage      '/CAOSMSG
            Call HandleChaosLegionMessage(Userindex)

        Case eGMCommands.CitizenMessage          '/CIUMSG
            Call HandleCitizenMessage(Userindex)

        Case eGMCommands.CriminalMessage         '/CRIMSG
            Call HandleCriminalMessage(Userindex)

        Case eGMCommands.TalkAsNPC               '/TALKAS
            Call HandleTalkAsNPC(Userindex)

        Case eGMCommands.DestroyAllItemsInArea   '/MASSDEST
            Call HandleDestroyAllItemsInArea(Userindex)

        Case eGMCommands.AcceptRoyalCouncilMember    '/ACEPTCONSE
            Call HandleAcceptRoyalCouncilMember(Userindex)

        Case eGMCommands.AcceptChaosCouncilMember    '/ACEPTCONSECAOS
            Call HandleAcceptChaosCouncilMember(Userindex)

        Case eGMCommands.ItemsInTheFloor         '/PISO
            Call HandleItemsInTheFloor(Userindex)

        Case eGMCommands.MakeDumb                '/ESTUPIDO
            Call HandleMakeDumb(Userindex)

        Case eGMCommands.MakeDumbNoMore          '/NOESTUPIDO
            Call HandleMakeDumbNoMore(Userindex)

        Case eGMCommands.dumpIPTables            '/DUMPSECURITY
            Call HandleDumpIPTables(Userindex)

        Case eGMCommands.CouncilKick             '/KICKCONSE
            Call HandleCouncilKick(Userindex)

        Case eGMCommands.SetTrigger              '/TRIGGER
            Call HandleSetTrigger(Userindex)

        Case eGMCommands.AskTrigger              '/TRIGGER with no args
            Call HandleAskTrigger(Userindex)

        Case eGMCommands.BannedIPList            '/BANIPLIST
            Call HandleBannedIPList(Userindex)

        Case eGMCommands.BannedIPReload          '/BANIPRELOAD
            Call HandleBannedIPReload(Userindex)

        Case eGMCommands.GuildMemberList         '/MIEMBROSCLAN
            Call HandleGuildMemberList(Userindex)

        Case eGMCommands.GuildBan                '/BANCLAN
            Call HandleGuildBan(Userindex)

        Case eGMCommands.BanIP                   '/BANIP
            Call HandleBanIP(Userindex)

        Case eGMCommands.UnbanIP                 '/UNBANIP
            Call HandleUnbanIP(Userindex)

        Case eGMCommands.CreateItem              '/CI
            Call HandleCreateItem(Userindex)

        Case eGMCommands.DestroyItems            '/DEST
            Call HandleDestroyItems(Userindex)

        Case eGMCommands.ChaosLegionKick         '/NOCAOS
            Call HandleChaosLegionKick(Userindex)

        Case eGMCommands.RoyalArmyKick           '/NOREAL
            Call HandleRoyalArmyKick(Userindex)

        Case eGMCommands.ForceMIDIAll            '/FORCEMIDI
            Call HandleForceMIDIAll(Userindex)

        Case eGMCommands.ForceWAVEAll            '/FORCEWAV
            Call HandleForceWAVEAll(Userindex)

        Case eGMCommands.RemovePunishment        '/BORRARPENA
            Call HandleRemovePunishment(Userindex)

        Case eGMCommands.TileBlockedToggle       '/BLOQ
            Call HandleTileBlockedToggle(Userindex)

        Case eGMCommands.KillNPCNoRespawn        '/MATA
            Call HandleKillNPCNoRespawn(Userindex)

        Case eGMCommands.KillAllNearbyNPCs       '/MASSKILL
            Call HandleKillAllNearbyNPCs(Userindex)

        Case eGMCommands.lastip                  '/LASTIP
            Call HandleLastIP(Userindex)

        Case eGMCommands.SystemMessage           '/SMSG
            Call HandleSystemMessage(Userindex)

        Case eGMCommands.CreateNPC               '/ACC
            Call HandleCreateNPC(Userindex)

        Case eGMCommands.CreateNPCWithRespawn    '/RACC
            Call HandleCreateNPCWithRespawn(Userindex)

        Case eGMCommands.ImperialArmour          '/AI1 - 4
            Call HandleImperialArmour(Userindex)

        Case eGMCommands.ChaosArmour             '/AC1 - 4
            Call HandleChaosArmour(Userindex)

        Case eGMCommands.NavigateToggle          '/NAVE
            Call HandleNavigateToggle(Userindex)

        Case eGMCommands.ServerOpenToUsersToggle    '/HABILITAR
            Call HandleServerOpenToUsersToggle(Userindex)

        Case eGMCommands.TurnOffServer           '/APAGAR
            Call HandleTurnOffServer(Userindex)

        Case eGMCommands.TurnCriminal            '/CONDEN
            Call HandleTurnCriminal(Userindex)

        Case eGMCommands.ResetFactionCaos           '/RAJAR
            Call HandleResetFactionCaos(Userindex)

        Case eGMCommands.ResetFactionReal           '/RAJAR
            Call HandleResetFactionReal(Userindex)

        Case eGMCommands.RemoveCharFromGuild     '/RAJARCLAN
            Call HandleRemoveCharFromGuild(Userindex)

        Case eGMCommands.RequestCharMail         '/LASTEMAIL
            Call HandleRequestCharMail(Userindex)

        Case eGMCommands.AlterPassword           '/APASS
            Call HandleAlterPassword(Userindex)

        Case eGMCommands.AlterMail               '/AEMAIL
            Call HandleAlterMail(Userindex)

        Case eGMCommands.AlterName               '/ANAME
            Call HandleAlterName(Userindex)

        Case eGMCommands.ToggleCentinelActivated    '/CENTINELAACTIVADO
            Call HandleToggleCentinelActivated(Userindex)

        Case Declaraciones.eGMCommands.DoBackUp               '/DOBACKUP
            Call HandleDoBackUp(Userindex)

        Case eGMCommands.ShowGuildMessages       '/SHOWCMSG
            Call HandleShowGuildMessages(Userindex)

        Case eGMCommands.SaveMap                 '/GUARDAMAPA
            Call HandleSaveMap(Userindex)

        Case eGMCommands.ChangeMapInfoPK         '/MODMAPINFO PK
            Call HandleChangeMapInfoPK(Userindex)

        Case eGMCommands.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
            Call HandleChangeMapInfoBackup(Userindex)

        Case eGMCommands.ChangeMapInfoRestricted    '/MODMAPINFO RESTRINGIR
            Call HandleChangeMapInfoRestricted(Userindex)

        Case eGMCommands.ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
            Call HandleChangeMapInfoNoMagic(Userindex)

        Case eGMCommands.ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
            Call HandleChangeMapInfoNoInvi(Userindex)

        Case eGMCommands.ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
            Call HandleChangeMapInfoNoResu(Userindex)

        Case eGMCommands.ChangeMapInfoLand       '/MODMAPINFO TERRENO
            Call HandleChangeMapInfoLand(Userindex)

        Case eGMCommands.ChangeMapInfoZone       '/MODMAPINFO ZONA
            Call HandleChangeMapInfoZone(Userindex)

        Case eGMCommands.ChangeMapInfoStealNpc   '/MODMAPINFO ROBONPC
            Call HandleChangeMapInfoStealNpc(Userindex)

        Case eGMCommands.ChangeMapInfoNoOcultar  '/MODMAPINFO OCULTARSINEFECTO
            Call HandleChangeMapInfoNoOcultar(Userindex)

        Case eGMCommands.ChangeMapInfoNoInvocar  '/MODMAPINFO INVOCARSINEFECTO
            Call HandleChangeMapInfoNoInvocar(Userindex)

        Case eGMCommands.SaveChars               '/GRABAR
            Call HandleSaveChars(Userindex)

        Case eGMCommands.CleanSOS                '/BORRAR SOS
            Call HandleCleanSOS(Userindex)

        Case eGMCommands.ShowServerForm          '/SHOW INT
            Call HandleShowServerForm(Userindex)

        Case eGMCommands.night                   '/NOCHE
            Call HandleNight(Userindex)


        Case eGMCommands.KickAllChars            '/ECHARTODOSPJS
            Call HandleKickAllChars(Userindex)

        Case eGMCommands.ReloadNPCs              '/RELOADNPCS
            Call HandleReloadNPCs(Userindex)

        Case eGMCommands.ReloadServerIni         '/RELOADSINI
            Call HandleReloadServerIni(Userindex)

        Case eGMCommands.ReloadSpells            '/RELOADHECHIZOS
            Call HandleReloadSpells(Userindex)

        Case eGMCommands.ReloadObjects           '/RELOADOBJ
            Call HandleReloadObjects(Userindex)

        Case eGMCommands.Restart                 '/REINICIAR
            Call HandleRestart(Userindex)

        Case eGMCommands.ResetAutoUpdate         '/AUTOUPDATE
            Call HandleResetAutoUpdate(Userindex)

        Case eGMCommands.ChatColor               '/CHATCOLOR
            Call HandleChatColor(Userindex)

        Case eGMCommands.Ignored                 '/IGNORADO
            Call HandleIgnored(Userindex)

        Case eGMCommands.CheckSlot               '/SLOT
            Call HandleCheckSlot(Userindex)


        Case eGMCommands.SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
            Call HandleSetIniVar(Userindex)


        Case eGMCommands.Seguimiento
            Call HandleSeguimiento(Userindex)

            '//Disco.
        Case eGMCommands.CheckHD                 '/VERHD NICKUSUARIO
            Call HandleCheckHD(Userindex)

        Case eGMCommands.BanHD                   '/BANHD NICKUSUARIO
            Call HandleBanHD(Userindex)

        Case eGMCommands.UnBanHD                 '/UNBANHD NICKUSUARIO
            Call HandleUnbanHD(Userindex)
            '///Disco.

        Case eGMCommands.MapMessage              '/MAPMSG
            Call HandleMapMessage(Userindex)

        Case eGMCommands.Impersonate             '/IMPERSONAR
            Call HandleImpersonate(Userindex)

        Case eGMCommands.Imitate                 '/MIMETIZAR
            Call HandleImitate(Userindex)


        Case eGMCommands.CambioPj                '/CAMBIO
            Call HandleCambioPj(Userindex)

        Case eGMCommands.CheckCPU_ID                 '/VERCPUID NICKUSUARIO
            Call HandleCheckCPU_ID(Userindex)

        Case eGMCommands.BanT0                   '/BANT0 NICKUSUARIO
            Call HandleBanT0(Userindex)

        Case eGMCommands.UnBanT0                 '/UNBANT0 NICKUSUARIO
            Call HandleUnbanT0(Userindex)
            
        Case eGMCommands.LarryMataNiños
            Call HandleLarryMataNiños(Userindex)
            
        Case eGMCommands.ComandoPorDias
            Call HandleComandoPorDias(Userindex)
            
        Case eGMCommands.DarPoints
            Call HandleDarPoints(Userindex)

        End Select
    End With

    Exit Sub

Errhandler:
    Call LogError("Error en GmCommands. Error: " & Err.Number & " - " & Err.Description & _
                  ". Paquete: " & Command)

End Sub

' ME VOY A FUMAR 340 PAQUETES POR VOS ALAN EMPEZANDO YA

''
' Handles the "Home" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleHome(ByVal Userindex As Integer)
'***************************************************
'Author: Budi
'Creation Date: 06/01/2010
'Last Modification: 05/06/10
'Pato - 05/06/10: Add the Ucase$ to prevent problems.
'***************************************************
    With UserList(Userindex)
        Call .incomingData.ReadByte
        
        If .flags.SlotEvent > 0 Then
            WriteConsoleMsg Userindex, "No puedes usar la restauración si estás en un evento.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
        If .flags.InCVC Then
            WriteConsoleMsg Userindex, "No puedes usar la restauración si estás en CVC.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
        If .flags.SlotRetoUser > 0 Then
            WriteConsoleMsg Userindex, "No puede susar este comando si estás en reto.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
        If .Pos.map = 66 Then
            Call WriteConsoleMsg(Userindex, "No puedes usar la restauración si estás en la carcel.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .Pos.map = 191 Then
            Call WriteConsoleMsg(Userindex, "No puedes usar la restauración si estás en los retos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .Pos.map = 176 Then
            Call WriteConsoleMsg(Userindex, "No puedes usar la restauración si estás en los retos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If .flags.Muerto = 0 Then
            Call WriteConsoleMsg(Userindex, "No puedes usar el comando si estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(Userindex).Stats.Gld < 7000 Then
            Call WriteConsoleMsg(Userindex, "No tienes suficientes monedas de oro, necesitas 7.000 monedas para usar la restauración de personaje.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld - 7000

        Call WriteUpdateGold(Userindex)
        WriteUpdateUserStats (Userindex)

        If .flags.Muerto = 1 Then
            Call WarpUserChar(Userindex, 1, 59, 45, True)
            Call WriteConsoleMsg(Userindex, "Has sido transportado Ullathorpe.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End With
End Sub

''
' Handles the "ConectarUsuarioE" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleConectarUsuarioE(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    #If SeguridadAlkon Then
        If UserList(Userindex).incomingData.length < 53 Then
            Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub
        End If
    #Else
        If UserList(Userindex).incomingData.length < 6 Then
            Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub
        End If
    #End If

    On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(Userindex).incomingData)

    'Remove packet ID
    Call buffer.ReadByte

    Dim DsForInblue As String

    Dim UserName As String
    Dim Password As String
    Dim version As String
    Dim HD     As String    '//Disco.
    Dim CPU_ID As String

    UserName = buffer.ReadASCIIString()
    Password = buffer.ReadASCIIString()
    DsForInblue = buffer.ReadASCIIString()

    'Convert version number to string
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())

    Dim MD5K   As String
    MD5K = buffer.ReadASCIIString()
    ELMD5 = MD5K

    If MD5ok(MD5K) = False Then
        WriteErrorMsg Userindex, "Versión obsoleta, verifique actualizaciones en la web o en el Autoupdater."
        If VersionOK(version) Then LogMD5 UserName & " ha intentado logear con un cliente NO VÁLIDO, MD5:" & MD5K
        FlushBuffer Userindex
        CloseSocket Userindex
        Exit Sub
    End If

    HD = buffer.ReadASCIIString()    '//Disco.
    CPU_ID = buffer.ReadASCIIString()
    
    Dim code As Integer
    code = buffer.ReadInteger
    If Not CheckCRC(Userindex, code - 42) Then Exit Sub

    If Not DsForInblue = "dasd#ewew%/#!4" Then
        Call WriteErrorMsg(Userindex, "Error crítico en el cliente. Por favor reinstale el juego.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        Exit Sub
    End If

    If Not AsciiValidos(UserName) Then
        Call WriteErrorMsg(Userindex, "Nombre inválido.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)

        Exit Sub
    End If

    If Not PersonajeExiste(UserName) Then
        Call WriteErrorMsg(Userindex, "El personaje no existe.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)

        Exit Sub
    End If

    If BuscarRegistroHD(HD) > 0 Or (BuscarRegistroT0(CPU_ID) > 0) Then
        Call WriteErrorMsg(Userindex, "Se te ha prohibido la entrada a Desterium AO. Baneado por " & ban_Reason(UserName))
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        Exit Sub
    End If


    If BANCheck(UserName) Then
        Call WriteErrorMsg(Userindex, "Se te ha prohibido la entrada a Desterium AO. Baneado por " & ban_Reason(UserName))
    ElseIf Not VersionOK(version) Then
        Call WriteErrorMsg(Userindex, "Versión obsoleta, descarga la nueva actualización " & ULTIMAVERSION & " desde la web o ejecuta el autoupdate.")
    Else
        Call ConnectUser(Userindex, UserName, Password, HD, CPU_ID)    '//Disco.
    End If


    'If we got here then packet is complete, copy data back to original queue
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "ThrowDices" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleThrowDices(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    'If Not CheckCRC(Userindex, UserList(Userindex).incomingData.ReadInteger - 42) Then Exit Sub
    
    With UserList(Userindex).Stats
        .UserAtributos(eAtributos.Fuerza) = RandomNumber(17, 18)
        .UserAtributos(eAtributos.Agilidad) = RandomNumber(17, 18)
        .UserAtributos(eAtributos.Inteligencia) = RandomNumber(17, 18)
        .UserAtributos(eAtributos.Carisma) = RandomNumber(17, 18)
        .UserAtributos(eAtributos.Constitucion) = RandomNumber(17, 18)
    End With

    Call WriteDiceRoll(Userindex)
End Sub

''
' Handles the "LogeaNuevoPj" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLogeaNuevoPj(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    If UserList(Userindex).incomingData.length < 15 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(Userindex).incomingData)

    'Remove packet ID
    Call buffer.ReadByte

    Dim UserName As String
    Dim Password As String
    Dim version As String
    Dim race   As eRaza
    Dim gender As eGenero
    Dim homeland As eCiudad
    Dim Class  As eClass
    Dim Head   As Integer
    Dim mail   As String
    Dim Pin    As String
    Dim HD     As String    '//Disco.
    Dim CPU_ID As String
    Const SecuritySV As String = "A$%bcdS4557Es7"
    

    If PuedeCrearPersonajes = 0 Then
        Call WriteErrorMsg(Userindex, "La creación de personajes en este servidor se ha deshabilitado.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)

        Exit Sub
    End If

    If ServerSoloGMs <> 0 Then
        Call WriteErrorMsg(Userindex, "Servidor restringido a administradores. Consulte la página oficial o el foro oficial para más información.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)

        Exit Sub
    End If

    If aClon.MaxPersonajes(UserList(Userindex).ip) Then
        Call WriteErrorMsg(Userindex, "Has creado demasiados personajes.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)

        Exit Sub
    End If

    UserName = buffer.ReadASCIIString()
    Password = buffer.ReadASCIIString()

    'Pin para el Personaje
    Pin = buffer.ReadASCIIString

    'Convert version number to string
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())

    race = buffer.ReadByte()
    gender = buffer.ReadByte()
    Class = buffer.ReadByte()
    Head = buffer.ReadInteger
    mail = buffer.ReadASCIIString()
    homeland = buffer.ReadByte()
    HD = buffer.ReadASCIIString()
    CPU_ID = buffer.ReadASCIIString()
    
    If buffer.ReadASCIIString() <> SecuritySV Then
        Call WriteErrorMsg(Userindex, "Cliente inválido")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        Exit Sub
    End If
    
    '//Disco.
    If (BuscarRegistroHD(HD) > 0) Or (BuscarRegistroT0(CPU_ID) > 0) Then    '//Disco.
        Call WriteErrorMsg(Userindex, "Se te ha prohibido la entrada a Desterium AO. Baneado por " & ban_Reason(UserName))
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        Exit Sub
    End If
    
    If Not VersionOK(version) Then
        Call WriteErrorMsg(Userindex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". La misma se encuentra disponible en https://www.desterium.com")
    Else
        Call ConnectNewUser(Userindex, UserName, Password, race, gender, Class, mail, homeland, Head, Pin, HD, CPU_ID)
    End If

    'If we got here then packet is complete, copy data back to original queue
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub
Public Sub HandleKillChar(Userindex)
' @@ 18/01/2015
' @@ Fix bug que explotaba el vb al detectar de que no existia un personaje

    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    Call buffer.ReadByte


    Dim UN     As String
    Dim PASS   As String
    Dim Pin    As String


    'leemos el nombre del usuario
    UN = buffer.ReadASCIIString
    'leemos el Password
    PASS = buffer.ReadASCIIString
    'leemos el PIN del usuario
    Pin = buffer.ReadASCIIString

    ' @@ Arregle la comprobación
    If Not FileExist(App.Path & "\CHARFILE\" & UN & ".chr") Then    'cacona
        Call WriteErrorMsg(Userindex, "El personaje no existe.")

        Call FlushBuffer(Userindex)
        Call TCP.CloseSocket(Userindex)
        Exit Sub
    Else
        'Call WriteErrorMsg(UserIndex, "El personaje existe.")
    End If


    'Comprobamos que los datos mandados sean iguales a lo que tenemos.
    If UCase(PASS) = UCase(GetVar(App.Path & "\CHARFILE\" & UN & ".chr", "INIT", "Password")) And _
       UCase(Pin) = UCase(GetVar(App.Path & "\CHARFILE\" & UN & ".chr", "INIT", "Pin")) Then
    
        If CBool(GetVar(App.Path & "\CHARFILE\" & UN & ".chr", "MERCADO", "InList")) Then
            WriteErrorMsg Userindex, "No puedes borrar un personaje que está en el mercado."
        Else
            'Borramos
            Call KillCharINFO(UN)
            Call WriteErrorMsg(Userindex, "Personaje Borrado Exitosamente.")
        End If
    Else
    
        'Mandamos el error por msgbox
        Call WriteErrorMsg(Userindex, "Los datos proporcionados no son correctos. Asegurese de haberlos ingresado bien.")
    End If

    Call UserList(Userindex).incomingData.CopyBuffer(buffer)

End Sub

Public Sub HandleRenewPassChar(Userindex)

    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    Call buffer.ReadByte

    Dim UN     As String
    Dim email  As String
    Dim Pin    As String


    'leemos el nombre del usuario
    UN = buffer.ReadASCIIString
    'leemos el email
    email = buffer.ReadASCIIString
    'leemos el PIN del usuario
    Pin = buffer.ReadASCIIString
        
    ' @@Ahora si editan paquetes no pueden _
    tirar el servidor por un loop infinito. Ahora a corregir 500 charfiles xd no
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)

    ' @@ Arregle la comprobación
    If Not FileExist(App.Path & "\CHARFILE\" & UN & ".chr") Then    'cacona
        Call WriteErrorMsg(Userindex, "El personaje no existe.")

        Call FlushBuffer(Userindex)
        Call TCP.CloseSocket(Userindex)
        
    Else
        'Comprobamos que los datos mandados sean iguales a lo que tenemos.
        If UCase(email) = UCase(GetVar(App.Path & "\CHARFILE\" & UN & ".chr", "CONTACTO", "Email")) And _
           UCase(Pin) = UCase(GetVar(App.Path & "\CHARFILE\" & UN & ".chr", "INIT", "Pin")) Then
    
            'Enviamos la nueva password
            Call WriteErrorMsg(Userindex, "Su personaje ha sido recuperado. Su nueva password es: " & "'" & GenerateRandomKey(UN) & "'")
    
        Else
            'Mandamos el error por Msgbox
            Call WriteErrorMsg(Userindex, "Los datos proporcionados no son correctos. Asegurese de haberlos ingresado bien.")
        End If
    End If

    

End Sub

Private Sub HandleTalk(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/09/2009
'15/07/2009: ZaMa - Now invisible admins talk by console.
'23/09/2009: ZaMa - Now invisible admins can't send empty chat.
'***************************************************

    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Dim CanTalk As Boolean
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim chat As String

        chat = buffer.ReadASCIIString()
        
        If CheckCRC(Userindex, buffer.ReadInteger - 42) Then
        
            '[Consejeros & GMs]
            If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                Call LogGM(.Name, "Dijo: " & chat)
            End If
    
            'I see you....
            If .flags.Oculto > 0 Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(Userindex, UserList(Userindex).Char.CharIndex, False)
                    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                    Call WriteConsoleMsg(Userindex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
    
            If LenB(chat) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(chat)
    
                ' If Not (.flags.AdminInvisible = 1) Then
                
                CanTalk = True
                If .flags.SlotEvent > 0 Then
                    If Events(.flags.SlotEvent).Modality = DeathMatch Then
                        CanTalk = False
                    End If
                End If
                    
                If CanTalk Then
                    If .flags.Muerto = 1 Then
                        Call SendData(SendTarget.ToDeadArea, Userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                    Else
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, .flags.ChatColor))
                    End If
                End If
    
            End If
        
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "Yell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleYell(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010 (ZaMa)
'15/07/2009: ZaMa - Now invisible admins yell by console.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim chat As String
        Dim CanTalk As Boolean
        
        chat = buffer.ReadASCIIString()
        
        If CheckCRC(Userindex, buffer.ReadInteger - 42) Then

            '[Consejeros & GMs]
            If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                Call LogGM(.Name, "Grito: " & chat)
            End If
    
            'I see you....
            If .flags.Oculto > 0 Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
    
                If .flags.Navegando = 1 Then
                    If .clase = eClass.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
                        Call ToggleBoatBody(Userindex)
                        Call WriteConsoleMsg(Userindex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.Heading, NingunArma, _
                                            NingunEscudo, NingunCasco)
                    End If
                Else
                    If .flags.invisible = 0 Then
                        Call UsUaRiOs.SetInvisible(Userindex, .Char.CharIndex, False)
                        Call WriteConsoleMsg(Userindex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
    
            If LenB(chat) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(chat)
                
                CanTalk = True
                If .flags.SlotEvent > 0 Then
                    If Events(.flags.SlotEvent).Modality = DeathMatch Then
                        CanTalk = False
                    End If
                End If
                
                If CanTalk Then
                    If .flags.Privilegios And PlayerType.User Then
                        If UserList(Userindex).flags.Muerto = 1 Then
                            Call SendData(SendTarget.ToDeadArea, Userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                        Else
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, vbRed))
                        End If
                    Else
                        If Not (.flags.AdminInvisible = 1) Then
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, CHAT_COLOR_GM_YELL))
                        Else
                        End If
                    End If
                End If
            End If
        End If


        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "Whisper" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 15/07/2009
'28/05/2009: ZaMa - Now it doesn't appear any message when private talking to an invisible admin
'***************************************************
    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim chat As String
        Dim targetCharIndex As Integer
        Dim targetUserIndex As Integer
        Dim targetPriv As PlayerType

        targetCharIndex = buffer.ReadInteger()
        chat = buffer.ReadASCIIString()

        If CheckCRC(Userindex, buffer.ReadInteger - 42) Then
            targetUserIndex = CharIndexToUserIndex(targetCharIndex)
    
            If .flags.Muerto Then
                Call WriteConsoleMsg(Userindex, "¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. ", FontTypeNames.FONTTYPE_INFO)
            Else
                If targetUserIndex = INVALID_INDEX Then
                    Call WriteConsoleMsg(Userindex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                Else
                    targetPriv = UserList(targetUserIndex).flags.Privilegios
                    'A los dioses y admins no vale susurrarles si no sos uno vos mismo (así no pueden ver si están conectados o no)
                    If (targetPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 And (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Then
                        ' Controlamos que no este invisible
                        If UserList(targetUserIndex).flags.AdminInvisible <> 1 Then
                            Call WriteConsoleMsg(Userindex, "No puedes susurrarle a los Dioses y Admins.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        'A los Consejeros y SemiDioses no vale susurrarles si sos un PJ común.
                    ElseIf (.flags.Privilegios And PlayerType.User) <> 0 And (Not targetPriv And PlayerType.User) <> 0 Then
                        ' Controlamos que no este invisible
                        If UserList(targetUserIndex).flags.AdminInvisible <> 1 Then
                            Call WriteConsoleMsg(Userindex, "No puedes susurrarle a los GMs.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    ElseIf Not EstaPCarea(Userindex, targetUserIndex) Then
                        Call WriteConsoleMsg(Userindex, "Estas muy lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
    
                    Else
                        '[Consejeros & GMs]
                        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                            Call LogGM(.Name, "Le dijo a '" & UserList(targetUserIndex).Name & "' " & chat)
                        End If
    
                        If LenB(chat) <> 0 Then
                            'Analize chat...
                            Call Statistics.ParseChat(chat)
    
                            If Not (.flags.AdminInvisible = 1) Then
                                Call WriteChatOverHead(Userindex, chat, .Char.CharIndex, vbYellow)
                                Call WriteChatOverHead(targetUserIndex, chat, .Char.CharIndex, vbYellow)
                                Call FlushBuffer(targetUserIndex)
    
                                '[CDT 17-02-2004]
                                If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                                    Call SendData(SendTarget.ToAdminsAreaButConsejeros, Userindex, PrepareMessageChatOverHead("A " & UserList(targetUserIndex).Name & "> " & chat, .Char.CharIndex, vbYellow))
                                End If
                            Else
                                Call SendData(SendTarget.ToAdminsAreaButConsejeros, Userindex, PrepareMessageChatOverHead("A " & UserList(targetUserIndex).Name & "> " & chat, .Char.CharIndex, vbYellow))
                                'Call WriteConsoleMsg(UserIndex, "Susurraste> " & chat, FontTypeNames.FONTTYPE_GM)
                                'If UserIndex <> targetUserIndex Then Call WriteConsoleMsg(targetUserIndex, "Gm susurra> " & chat, FontTypeNames.FONTTYPE_GM)
    
                                'If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                                '    Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageConsoleMsg("Gm dijo a " & UserList(targetUserIndex).name & "> " & chat, FontTypeNames.FONTTYPE_GM))
                                'End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "Walk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010 (ZaMa)
'11/19/09 Pato - Now the class bandit can walk hidden.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    Dim dummy  As Long
    Dim TempTick As Long
    Dim Heading As eHeading

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Heading = .incomingData.ReadByte()
        If Not CheckCRC(Userindex, .incomingData.ReadInteger - 42) Then Exit Sub
        
        'Prevent SpeedHack
        If .flags.TimesWalk >= 30 Then
            TempTick = GetTickCount And &H7FFFFFFF
            dummy = (TempTick - .flags.StartWalk)

            ' 5800 is actually less than what would be needed in perfect conditions to take 30 steps
            '(it's about 193 ms per step against the over 200 needed in perfect conditions)
            If dummy < 5800 Then
                If TempTick - .flags.CountSH > 30000 Then
                    .flags.CountSH = 0
                End If

                If .flags.Montando Then
                    If TempTick - .flags.CountSH < 45000 Then
                        .flags.CountSH = 0
                    End If
                End If

                If Not .flags.CountSH = 0 Then
                    If dummy <> 0 Then _
                       dummy = 126000 \ dummy

                    Call LogHackAttemp("Tramposo SH: " & .Name & " , " & dummy)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha sido echado por el servidor por posible uso de SH.", FontTypeNames.FONTTYPE_SERVER))
                    Call CloseSocket(Userindex)

                    Exit Sub
                Else
                    .flags.CountSH = TempTick
                End If
            End If
            .flags.StartWalk = TempTick
            .flags.TimesWalk = 0
        End If

        .flags.TimesWalk = .flags.TimesWalk + 1

        'If exiting, cancel
        Call CancelExit(Userindex)

        'TODO: Debería decirle por consola que no puede?
        'Esta usando el /HOGAR, no se puede mover
        If .flags.Traveling = 1 Then Exit Sub

        If .flags.Paralizado = 0 Then
            If .flags.Meditando Then
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                .Char.FX = 0
                .Char.loops = 0

                Call WriteConsoleMsg(Userindex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)

                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            End If
            
            'Move user
            Call MoveUserChar(Userindex, Heading)

            'Stop resting if needed
            If .flags.Descansar Then
                .flags.Descansar = False

                Call WriteRestOK(Userindex)
                Call WriteConsoleMsg(Userindex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else    'paralized
            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1

                Call WriteConsoleMsg(Userindex, "No puedes moverte porque estás paralizado.", FontTypeNames.FONTTYPE_INFO)
            End If

            .flags.CountSH = 0
        End If

        'Can't move while hidden except he is a thief
        If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
            If .clase <> eClass.Thief Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0

                If .flags.Navegando = 1 Then
                    If .clase = eClass.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
                        Call ToggleBoatBody(Userindex)

                        Call WriteConsoleMsg(Userindex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.Heading, NingunArma, _
                                            NingunEscudo, NingunCasco)
                    End If
                Else
                    'If not under a spell effect, show char
                    If .flags.invisible = 0 Then
                        Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                        Call UsUaRiOs.SetInvisible(Userindex, .Char.CharIndex, False)
                    End If
                End If
            End If
        End If
    End With
End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    UserList(Userindex).incomingData.ReadByte

    Call WritePosUpdate(Userindex)
End Sub
Private Sub HandleAttack(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010
'Last Modified By: ZaMa
'10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo.
'13/11/2009: ZaMa - Se cancela el estado no atacable al atcar.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If Not CheckCRC(Userindex, .incomingData.ReadInteger - 42) Then Exit Sub
        
        'If dead, can't attack
        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'If user meditates, can't attack
        If .flags.Meditando Then
            Exit Sub
        End If

        If .flags.ModoCombate = False Then
            WriteConsoleMsg Userindex, "Necesitas estar en Modo Combate para atacar", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If

        'If equiped weapon is ranged, can't attack this way
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
                Call WriteConsoleMsg(Userindex, "No puedes usar así este arma.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If

        'If exiting, cancel
        Call CancelExit(Userindex)
        If (Mod_AntiCheat.PuedoPegar(Userindex) = False) Then Exit Sub
        'Attack!
        Call UsuarioAtaca(Userindex)

        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False

        'I see you...
        If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0

            If .flags.Navegando = 1 Then
                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(Userindex)
                    Call WriteConsoleMsg(Userindex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.Heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                End If
            Else
                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(Userindex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(Userindex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With
End Sub


Private Sub HandlePickUp(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/25/09
'02/26/2006: Marco - Agregué un checkeo por si el usuario trata de agarrar un item mientras comercia.
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'If dead, it can't pick up objects
        If .flags.Muerto = 1 Then Exit Sub

        'If user is trading items and attempts to pickup an item, he's cheating, so we kick him.
        If .flags.Comerciando Then Exit Sub

        'Lower rank administrators can't pick up items

        Call GetObj(Userindex)
    End With
End Sub
Private Sub HanldeCombatModeToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.ModoCombate Then
            Call WriteConsoleMsg(Userindex, "Has salido del modo combate.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(Userindex, "Has pasado al modo combate.", FontTypeNames.FONTTYPE_INFO)
        End If

        .flags.ModoCombate = Not .flags.ModoCombate
    End With
End Sub

''
' Handles the "SafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSafeToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Seguro Then
            Call WriteMultiMessage(Userindex, eMessages.SafeModeOff)    'Call WriteSafeModeOff(UserIndex)
        Else
            Call WriteMultiMessage(Userindex, eMessages.SafeModeOn)    'Call WriteSafeModeOn(UserIndex)
        End If

        .flags.Seguro = Not .flags.Seguro
    End With
End Sub

''
' Handles the "ResuscitationSafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResuscitationToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Rapsodius
'Creation Date: 10/10/07
'***************************************************
    With UserList(Userindex)
        Call .incomingData.ReadByte

        .flags.SeguroResu = Not .flags.SeguroResu

        If .flags.SeguroResu Then
            Call WriteMultiMessage(Userindex, eMessages.ResuscitationSafeOn)    'Call WriteResuscitationSafeOn(UserIndex)
        Else
            Call WriteMultiMessage(Userindex, eMessages.ResuscitationSafeOff)    'Call WriteResuscitationSafeOff(UserIndex)
        End If
    End With
End Sub

''
' Handles the "RequestGuildLeaderInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestGuildLeaderInfo(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    UserList(Userindex).incomingData.ReadByte

    Call modGuilds.SendGuildLeaderInfo(Userindex)
End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAtributes(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    Call WriteAttributes(Userindex)
End Sub

''
' Handles the "RequestFame" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestFame(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    Call EnviarFama(Userindex)
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    Call WriteSendSkills(Userindex)
End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMiniStats(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    Call WriteMiniStats(Userindex)
End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    'User quits commerce mode
    UserList(Userindex).flags.Comerciando = False
    Call WriteCommerceEnd(Userindex)
End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 11/03/2010
'11/03/2010: ZaMa - Le avisa por consola al que cencela que dejo de comerciar.
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Quits commerce mode with user
        If .ComUsu.DestUsu > 0 Then
            If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = Userindex Then
                Call WriteConsoleMsg(.ComUsu.DestUsu, .Name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_GUILD)
                Call FinComerciarUsu(.ComUsu.DestUsu)

                'Send data in the outgoing buffer of the other user
                Call FlushBuffer(.ComUsu.DestUsu)
            End If
        End If

        Call FinComerciarUsu(Userindex)
        Call WriteConsoleMsg(Userindex, "Has dejado de comerciar.", FontTypeNames.FONTTYPE_GUILD)
    End With

End Sub

''
' Handles the "UserCommerceConfirm" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUserCommerceConfirm(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************

'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    'Validate the commerce
    If PuedeSeguirComerciando(Userindex) Then
        'Tell the other user the confirmation of the offer
        Call WriteUserOfferConfirm(UserList(Userindex).ComUsu.DestUsu)
        UserList(Userindex).ComUsu.Confirmo = True
    End If

End Sub

Private Sub HandleCommerceChat(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim chat As String

        chat = buffer.ReadASCIIString()

        If LenB(chat) <> 0 Then
            If PuedeSeguirComerciando(Userindex) Then
                'Analize chat...
                Call Statistics.ParseChat(chat)

                chat = UserList(Userindex).Name & "> " & chat
                Call WriteCommerceChat(Userindex, chat, FontTypeNames.FONTTYPE_PARTY)
                Call WriteCommerceChat(UserList(Userindex).ComUsu.DestUsu, chat, FontTypeNames.FONTTYPE_PARTY)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub


''
' Handles the "BankEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'User exits banking mode
        .flags.Comerciando = False
        Call WriteBankEnd(Userindex)
    End With
End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOk(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

If UserList(Userindex).ComUsu.Confirmo = False Then Exit Sub

    'Trade accepted
    Call AceptarComercioUsu(Userindex)
End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim otherUser As Integer

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        otherUser = .ComUsu.DestUsu

        'Offer rejected
        If otherUser > 0 Then
            If UserList(otherUser).flags.UserLogged Then
                Call WriteConsoleMsg(otherUser, .Name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_GUILD)
                Call FinComerciarUsu(otherUser)

                'Send data in the outgoing buffer of the other user
                Call FlushBuffer(otherUser)
            End If
        End If

        Call WriteConsoleMsg(Userindex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_GUILD)
        Call FinComerciarUsu(Userindex)
    End With
End Sub

''
' Handles the "Drop" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/25/09
'07/25/09: Marco - Agregué un checkeo para patear a los usuarios que tiran items mientras comercian.
'***************************************************
    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    Dim Slot As Byte, Amount As Integer

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        If Not CheckCRC(Userindex, UserList(Userindex).incomingData.ReadInteger - 42) Then Exit Sub
        
        'low rank admins can't drop item. Neither can the dead nor those sailing.
        If .flags.Navegando = 1 Or _
           .flags.Montando = 1 Or _
           .flags.Muerto = 1 Then Exit Sub
        ' ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0)

        If Amount > 10000 Then Amount = 10000


        If Slot = FLAGORO Then
            If Amount <= 0 Or .Stats.Gld < Amount Then
                'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Anti Cheat > El usuario " & UserList(Userindex).Name & " está intentado dupear oro (Drop).", FontTypeNames.FONTTYPE_ADMIN))
                'Call LogAntiCheat(UserList(Userindex).Name & " intentó dupear oro.)")
                Exit Sub
            End If
        
        Else
            If Amount <= 0 Or Amount > UserList(Userindex).Invent.Object(Slot).Amount Then
                'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Anti Cheat > El usuario " & UserList(Userindex).Name & " está intentado tirar oro dupeado.", FontTypeNames.FONTTYPE_ADMIN))
               ' Call LogAntiCheat(UserList(Userindex).Name & " intentó dupear oro.)")
                Exit Sub
            End If
        End If

        'If the user is trading, he can't drop items => He's cheating, we kick him.
        If .flags.Comerciando Then Exit Sub

        'Are we dropping gold or other items??
        If Slot = FLAGORO Then
            If Amount > 10000 Then Exit Sub    'Don't drop too much gold
            'Call TirarOro(Amount, UserIndex)

            Call WriteUpdateGold(Userindex)
        Else
            'Only drop valid slots
            If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
                If .Invent.Object(Slot).objindex = 0 Then
                    Exit Sub
                End If
                Call DropObj(Userindex, Slot, Amount, .Pos.map, .Pos.X, .Pos.Y)
            End If
        End If

    End With
End Sub

''
' Handles the "CastSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCastSpell(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'13/11/2009: ZaMa - Ahora los npcs pueden atacar al usuario si quizo castear un hechizo
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim spell As Byte

        spell = .incomingData.ReadByte()

        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If .flags.MenuCliente <> 255 Then
            If .flags.MenuCliente <> 1 Then
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("ANTICHEAT > Vigilar a " & .Name, _
                                                                               FontTypeNames.FONTTYPE_EJECUCION))
                Exit Sub

            End If

        End If

        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False

        If spell < 1 Then
            .flags.Hechizo = 0
            Exit Sub
        ElseIf spell > MAXUSERHECHIZOS Then
            .flags.Hechizo = 0
            Exit Sub
        End If

        .flags.Hechizo = .Stats.UserHechizos(spell)
    End With
End Sub

''
' Handles the "LeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex).incomingData
        'Remove packet ID
        Call .ReadByte

        Dim X  As Byte
        Dim Y  As Byte

        X = .ReadByte()
        Y = .ReadByte()

        Call LookatTile(Userindex, UserList(Userindex).Pos.map, X, Y)
    End With
End Sub

''
' Handles the "DoubleClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDoubleClick(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex).incomingData
        'Remove packet ID
        Call .ReadByte

        Dim X  As Byte
        Dim Y  As Byte

        X = .ReadByte()
        Y = .ReadByte()

        Call Accion(Userindex, UserList(Userindex).Pos.map, X, Y)
    End With
End Sub

''
' Handles the "Work" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWork(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010 (ZaMa)
'13/01/2010: ZaMa - El pirata se puede ocultar en barca
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Skill As eSkill

        Skill = .incomingData.ReadByte()

        If UserList(Userindex).flags.Muerto = 1 Then Exit Sub

        'If exiting, cancel
        Call CancelExit(Userindex)

        Select Case Skill

        Case Robar, Magia, Domar
            Call WriteMultiMessage(Userindex, eMessages.WorkRequestTarget, Skill)

        Case Ocultarse

            ' Verifico si se peude ocultar en este mapa
            If MapInfo(.Pos.map).OcultarSinEfecto = 1 Then
                Call WriteConsoleMsg(Userindex, "¡Ocultarse no funciona aquí!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            If .flags.EnConsulta Then
                Call WriteConsoleMsg(Userindex, "No puedes ocultarte si estás en consulta.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If .flags.SlotReto > 0 Then
                WriteConsoleMsg Userindex, "No puedes ocultarte en reto.", FontTypeNames.FONTTYPE_INFO
                Exit Sub
            End If
            
            If .flags.SlotEvent > 0 Then
                WriteConsoleMsg Userindex, "No puedes ocultarte en evento.", FontTypeNames.FONTTYPE_INFO
                Exit Sub
            End If

            If .flags.Navegando = 1 Then
                If .clase <> eClass.Pirat Then
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 3 Then
                        Call WriteConsoleMsg(Userindex, "No puedes ocultarte si estás navegando.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 3
                    End If
                    '[/CDT]
                    Exit Sub
                End If
            End If


            If .flags.Montando = 1 Then
                '[CDT 17-02-2004]
                If Not .flags.UltimoMensaje = 3 Then
                    Call WriteConsoleMsg(Userindex, "No puedes ocultarte si estás montando.", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 3
                End If
                '[/CDT]
                Exit Sub
            End If

            If .flags.Oculto = 1 Then
                '[CDT 17-02-2004]
                If Not .flags.UltimoMensaje = 2 Then
                    Call WriteConsoleMsg(Userindex, "Ya estás oculto.", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 2
                End If
                '[/CDT]
                Exit Sub
            End If

            Call DoOcultarse(Userindex)

        End Select

    End With
End Sub


''
' Handles the "InitCrafting" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInitCrafting(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 29/01/2010
'
'***************************************************
    Dim TotalItems As Long
    Dim ItemsPorCiclo As Integer

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        TotalItems = .incomingData.ReadLong
        ItemsPorCiclo = .incomingData.ReadInteger

        If TotalItems > 0 Then

            .Construir.Cantidad = TotalItems
            .Construir.PorCiclo = MinimoInt(MaxItemsConstruibles(Userindex), ItemsPorCiclo)

        End If
    End With
End Sub

''
' Handles the "UseSpellMacro" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseSpellMacro(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Call SendData(SendTarget.ToAdmins, Userindex, PrepareMessageConsoleMsg(.Name & " fue expulsado por Anti-macro de hechizos.", FontTypeNames.FONTTYPE_VENENO))
        Call WriteErrorMsg(Userindex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
    End With
End Sub

''
' Handles the "UseItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Slot As Byte
        'Dim tipo As Byte

        Slot = .incomingData.ReadByte()
        If Not CheckCRC(Userindex, .incomingData.ReadInteger - 42) Then Exit Sub
        
        .incomingData.ReadInteger
        .incomingData.ReadByte
        
        If .flags.LastSlotClient <> 255 Then
            If Slot <> .flags.LastSlotClient Then
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("ANTICHEAT > VIGILAR ACTITUD MUY SOSPECHOSA a " & .Name & " Informacion confidencial. ", FontTypeNames.FONTTYPE_EJECUCION))
                Call LogAntiCheat(.Name & " Cambio de slot estando en la ventana de hechizos.")
                Exit Sub

            End If

        End If

        If Slot <= .CurrentInventorySlots And Slot > 0 Then
            If .Invent.Object(Slot).objindex = 0 Then Exit Sub
        End If

        If .flags.Meditando Then
            Exit Sub    'The error message should have been provided by the client.
        End If

        If ObjData(.Invent.Object(Slot).objindex).OBJType = otPociones Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("ANTICHEAT > MISERY DLL DETECTED BY LAUTARO:  " & .Name & ".", FontTypeNames.FONTTYPE_EJECUCION))
            Call LogAntiCheat("ANTICHEAT > MISERY DLL DETECTED BY LAUTARO:  " & .Name & ".")
        End If
        
        Call UseInvItem(Userindex, Slot)
        Call WriteUpdateFollow(Userindex)

    End With
End Sub

Private Sub HandleUsePotions(ByVal Userindex As Integer)

    With UserList(Userindex)

        Call .incomingData.ReadByte

        Dim Slot As Byte
        Dim byClick As Byte

        Slot = .incomingData.ReadByte()
        byClick = .incomingData.ReadByte()
        
        If Not CheckCRC(Userindex, .incomingData.ReadInteger - 42) Then Exit Sub
        
        .incomingData.ReadByte
        .incomingData.ReadByte
        
        If byClick = 0 Then
            If .flags.MenuCliente <> 255 Then
                If .flags.MenuCliente <> 2 Then
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("ANTICHEAT > Vigilar a " & .Name & _
                                                                                   " posible AutoPots.", FontTypeNames.FONTTYPE_EJECUCION))
                    Exit Sub

                End If

            End If

        End If

        If .flags.LastSlotClient <> 255 Then

            If Slot <> .flags.LastSlotClient Then
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg( _
                                                      "ANTICHEAT > VIGILAR ACTITUD MUY SOSPECHOSA a " & .Name & " Informacion confidencial. ", _
                              FontTypeNames.FONTTYPE_EJECUCION))
                Call LogAntiCheat(.Name & " Cambio de slot estando en la ventana de hechizos.")
                Exit Sub

            End If

        End If

        If Slot <= .CurrentInventorySlots And Slot > 0 Then
            If .Invent.Object(Slot).objindex = 0 Then Exit Sub

        End If

        If .flags.Meditando Then

            Exit Sub    'The error message should have been provided by the client.

        End If
        ' Esto de aca, como te dije no deberia, pero puede psar
        'If ((Mod_AntiCheat.PuedoUsar(Userindex, byClick)) = False) Then Exit Sub

        Call UseInvPotion(Userindex, Slot, byClick)
        Call WriteUpdateFollow(Userindex)

    End With

End Sub

Private Sub HandleSetMenu(ByVal Userindex As Integer)

    With UserList(Userindex)


        Call .incomingData.ReadByte

        '1 spell
        '2 inventario

        .flags.MenuCliente = .incomingData.ReadByte
        .flags.LastSlotClient = .incomingData.ReadByte

    End With

End Sub

''
' Handles the "CraftBlacksmith" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftBlacksmith(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex).incomingData
        'Remove packet ID
        Call .ReadByte

        Dim Item As Integer

        Item = .ReadInteger()

        If Item < 1 Then Exit Sub

        If ObjData(Item).SkHerreria = 0 Then Exit Sub

        Call HerreroConstruirItem(Userindex, Item)
    End With
End Sub

''
' Handles the "CraftCarpenter" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftCarpenter(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex).incomingData
        'Remove packet ID
        Call .ReadByte

        Dim Item As Integer

        Item = .ReadInteger()

        If Item < 1 Then Exit Sub

        If ObjData(Item).SkCarpinteria = 0 Then Exit Sub

        Call CarpinteroConstruirItem(Userindex, Item)
    End With
End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 14/01/2010 (ZaMa)
'16/11/2009: ZaMa - Agregada la posibilidad de extraer madera elfica.
'12/01/2010: ZaMa - Ahora se admiten armas arrojadizas (proyectiles sin municiones).
'14/01/2010: ZaMa - Ya no se pierden municiones al atacar npcs con dueño.
'***************************************************
    If UserList(Userindex).incomingData.length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim X  As Byte
        Dim Y  As Byte
        Dim Skill As eSkill
        Dim DummyInt As Integer
        Dim tU As Integer   'Target user
        Dim tN As Integer   'Target NPC

        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()

        Skill = .incomingData.ReadByte()

        
        If Not CheckCRC(Userindex, UserList(Userindex).incomingData.ReadInteger - 42) Then Exit Sub

        .incomingData.ReadByte
        .incomingData.ReadInteger
        
        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando _
           Or Not InMapBounds(.Pos.map, X, Y) Then Exit Sub

        If Not InRangoVision(Userindex, X, Y) Then
            Call WritePosUpdate(Userindex)
            Exit Sub
        End If

        'If exiting, cancel
        Call CancelExit(Userindex)

        Select Case Skill
        Case eSkill.Proyectiles

            'Check attack interval
            If Not IntervaloPermiteAtacar(Userindex, False) Then Exit Sub
            'Check Magic interval
            If Not IntervaloPermiteLanzarSpell(Userindex, False) Then Exit Sub
            'Check bow's interval
            If Not IntervaloPermiteUsarArcos(Userindex) Then Exit Sub

            Dim Atacked As Boolean
            Atacked = True

            'Make sure the item is valid and there is ammo equipped.
            With .Invent
                ' Tiene arma equipada?
                If .WeaponEqpObjIndex = 0 Then
                    DummyInt = 1
                    ' En un slot válido?
                ElseIf .WeaponEqpSlot < 1 Or .WeaponEqpSlot > UserList(Userindex).CurrentInventorySlots Then
                    DummyInt = 1
                    ' Usa munición? (Si no la usa, puede ser un arma arrojadiza)
                ElseIf ObjData(.WeaponEqpObjIndex).Municion = 1 Then
                    ' La municion esta equipada en un slot valido?
                    If .MunicionEqpSlot < 1 Or .MunicionEqpSlot > UserList(Userindex).CurrentInventorySlots Then
                        DummyInt = 1
                        ' Tiene munición?
                    ElseIf .MunicionEqpObjIndex = 0 Then
                        DummyInt = 1
                        ' Son flechas?
                    ElseIf ObjData(.MunicionEqpObjIndex).OBJType <> eOBJType.otFlechas Then
                        DummyInt = 1
                        ' Tiene suficientes?
                    ElseIf .Object(.MunicionEqpSlot).Amount < 1 Then
                        DummyInt = 1
                    End If
                    ' Es un arma de proyectiles?
                ElseIf ObjData(.WeaponEqpObjIndex).proyectil <> 1 Then
                    DummyInt = 2
                End If

                If DummyInt <> 0 Then
                    If DummyInt = 1 Then
                        Call WriteConsoleMsg(Userindex, "No tienes municiones.", FontTypeNames.FONTTYPE_INFO)

                        Call Desequipar(Userindex, .WeaponEqpSlot)
                    End If

                    Call Desequipar(Userindex, .MunicionEqpSlot)
                    Exit Sub
                End If
            End With

            'Quitamos stamina
            If .Stats.MinSta >= 10 Then
                Call QuitarSta(Userindex, RandomNumber(1, 10))
            Else
                If .Genero = eGenero.Hombre Then
                    Call WriteConsoleMsg(Userindex, "Estás muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, "Estás muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)
                End If
                Exit Sub
            End If

            SendData SendTarget.ToPCArea, Userindex, PrepareMessageMovimientSW(.Char.CharIndex, 1)

            Call LookatTile(Userindex, .Pos.map, X, Y)

            tU = .flags.TargetUser
            tN = .flags.TargetNPC

            'Validate target
            If tU > 0 Then
                'Only allow to atack if the other one can retaliate (can see us)
                If Abs(UserList(tU).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
                    Call WriteConsoleMsg(Userindex, "Estás demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub
                End If

                'Prevent from hitting self
                If tU = Userindex Then
                    Call WriteConsoleMsg(Userindex, "¡No puedes atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                'Attack!
                Atacked = UsuarioAtacaUsuario(Userindex, tU)


            ElseIf tN > 0 Then
                'Only allow to atack if the other one can retaliate (can see us)
                If Abs(Npclist(tN).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(Npclist(tN).Pos.X - .Pos.X) > RANGO_VISION_X Then
                    Call WriteConsoleMsg(Userindex, "Estás demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub
                End If

                'Is it attackable???
                If Npclist(tN).Attackable <> 0 Then

                    'Attack!
                    Atacked = UsuarioAtacaNpc(Userindex, tN)
                End If
            End If

            ' Solo pierde la munición si pudo atacar al target, o tiro al aire
            If Atacked Then
                With .Invent
                    ' Tiene equipado arco y flecha?
                    If ObjData(.WeaponEqpObjIndex).Municion = 1 Then
                        DummyInt = .MunicionEqpSlot


                        'Take 1 arrow away - we do it AFTER hitting, since if Ammo Slot is 0 it gives a rt9 and kicks players
                        Call QuitarUserInvItem(Userindex, DummyInt, 1)

                        If .Object(DummyInt).Amount > 0 Then
                            'QuitarUserInvItem unequips the ammo, so we equip it again
                            .MunicionEqpSlot = DummyInt
                            .MunicionEqpObjIndex = .Object(DummyInt).objindex
                            .Object(DummyInt).Equipped = 1
                        Else
                            .MunicionEqpSlot = 0
                            .MunicionEqpObjIndex = 0
                        End If
                        ' Tiene equipado un arma arrojadiza
                    Else
                        DummyInt = .WeaponEqpSlot

                        'Take 1 knife away
                        Call QuitarUserInvItem(Userindex, DummyInt, 1)

                        If .Object(DummyInt).Amount > 0 Then
                            'QuitarUserInvItem unequips the weapon, so we equip it again
                            .WeaponEqpSlot = DummyInt
                            .WeaponEqpObjIndex = .Object(DummyInt).objindex
                            .Object(DummyInt).Equipped = 1
                        Else
                            .WeaponEqpSlot = 0
                            .WeaponEqpObjIndex = 0
                        End If

                    End If

                    Call UpdateUserInv(False, Userindex, DummyInt)
                End With
            End If

        Case eSkill.Magia
            'Check the map allows spells to be casted.
            If MapInfo(.Pos.map).MagiaSinEfecto > 0 Then
                Call WriteConsoleMsg(Userindex, "Una fuerza oscura te impide canalizar tu energía.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If

            'Target whatever is in that tile
            Call LookatTile(Userindex, .Pos.map, X, Y)

            'If it's outside range log it and exit
            If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                Call LogCheating("Ataque fuera de rango de " & .Name & "(" & .Pos.map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .ip & " a la posición (" & .Pos.map & "/" & X & "/" & Y & ")")
                Exit Sub
            End If

            'Check bow's interval
            If Not IntervaloPermiteUsarArcos(Userindex, False) Then Exit Sub

            If (Mod_AntiCheat.PuedoCasteoHechizo(Userindex) = False) Then Exit Sub
            
            'Check Spell-Hit interval
            If Not IntervaloPermiteGolpeMagia(Userindex) Then
                'Check Magic interval
                If Not IntervaloPermiteLanzarSpell(Userindex) Then
                    Exit Sub
                End If
            End If


            'Check intervals and cast
            If .flags.Hechizo > 0 Then
                Call LanzarHechizo(.flags.Hechizo, Userindex)
                .flags.Hechizo = 0
            Else
                Call WriteConsoleMsg(Userindex, "¡Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)
            End If

        Case eSkill.Pesca
            DummyInt = .Invent.WeaponEqpObjIndex
            If DummyInt = 0 Then Exit Sub

            'Check interval
            If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub

            'Basado en la idea de Barrin
            'Comentario por Barrin: jah, "basado", caradura ! ^^
            If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 1 Then
                Call WriteConsoleMsg(Userindex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            If HayAgua(.Pos.map, X, Y) Then
                Select Case DummyInt
                Case CAÑA_PESCA
                    Call DoPescar(Userindex)

                Case RED_PESCA
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(Userindex, "Estás demasiado lejos para pescar.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    Call DoPescarRed(Userindex)

                Case Else
                    Exit Sub    'Invalid item!
                End Select

                'Play sound!
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
            Else
                Call WriteConsoleMsg(Userindex, "No hay agua donde pescar. Busca un lago, río o mar.", FontTypeNames.FONTTYPE_INFO)
            End If

        Case eSkill.Robar
            'Does the map allow us to steal here?
            If MapInfo(.Pos.map).Pk Then

                'Check interval
                If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub

                'Target whatever is in that tile
                Call LookatTile(Userindex, UserList(Userindex).Pos.map, X, Y)

                tU = .flags.TargetUser

                If tU > 0 And tU <> Userindex Then
                    'Can't steal administrative players
                    If UserList(tU).flags.Privilegios And PlayerType.User Then
                        If UserList(tU).flags.Muerto = 0 Then
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 1 Then
                                'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
                                Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If

                            '17/09/02
                            'Check the trigger
                            If MapData(UserList(tU).Pos.map, X, Y).trigger = eTrigger.ZONASEGURA Then
                                Call WriteConsoleMsg(Userindex, "No puedes robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                Exit Sub
                            End If

                            If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                Call WriteConsoleMsg(Userindex, "No puedes robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                Exit Sub
                            End If

                            Call DoRobar(Userindex, tU)
                        End If
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "¡No hay a quien robarle!", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(Userindex, "¡No puedes robar en zonas seguras!", FontTypeNames.FONTTYPE_INFO)
            End If

        Case eSkill.talar
            'Check interval
            If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub

            If .Invent.WeaponEqpObjIndex = 0 Then
                Call WriteConsoleMsg(Userindex, "Deberías equiparte el hacha.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            If .Invent.WeaponEqpObjIndex <> HACHA_LEÑADOR And _
               .Invent.WeaponEqpObjIndex <> HACHA_DORADA Then
                ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                Exit Sub
            End If

            DummyInt = MapData(.Pos.map, X, Y).ObjInfo.objindex

            If DummyInt > 0 Then
                If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                    'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
                    Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                'Barrin 29/9/03
                If .Pos.X = X And .Pos.Y = Y Then
                    Call WriteConsoleMsg(Userindex, "No puedes talar desde allí.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                ArbT = DummyInt
                '¿Hay un arbol donde clickeo?
                If ObjData(DummyInt).OBJType = eOBJType.otarboles Then
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
                    Call DoTalar(Userindex)
                ElseIf ObjData(DummyInt).OBJType = 38 Then
                    SendData SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y)

                    DoTalar Userindex
                End If
            Else
                Call WriteConsoleMsg(Userindex, "No hay ningún árbol ahí.", FontTypeNames.FONTTYPE_INFO)
            End If

        Case eSkill.Mineria
            If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub

            If .Invent.WeaponEqpObjIndex = 0 Then Exit Sub

            If Not ((.Invent.WeaponEqpObjIndex <> PIQUETE_MINERO) Or (.Invent.WeaponEqpObjIndex <> PIQUETE_ORO)) Then
                ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                Exit Sub
            End If

            'Target whatever is in the tile
            Call LookatTile(Userindex, .Pos.map, X, Y)

            DummyInt = MapData(.Pos.map, X, Y).ObjInfo.objindex

            If DummyInt > 0 Then
                'Check distance
                If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                    'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
                    Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                DummyInt = MapData(.Pos.map, X, Y).ObjInfo.objindex    'CHECK
                '¿Hay un yacimiento donde clickeo?
                If ObjData(DummyInt).OBJType = eOBJType.otYacimiento Then
                    Call DoMineria(Userindex)
                Else
                    Call WriteConsoleMsg(Userindex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(Userindex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
            End If

        Case eSkill.Domar
            'Modificado 25/11/02
            'Optimizado y solucionado el bug de la doma de
            'criaturas hostiles.

            'Target whatever is that tile
            Call LookatTile(Userindex, .Pos.map, X, Y)
            tN = .flags.TargetNPC

            If tN > 0 Then
                If Npclist(tN).flags.Domable > 0 Then
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
                        Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    If LenB(Npclist(tN).flags.AttackedBy) <> 0 Then
                        Call WriteConsoleMsg(Userindex, "No puedes domar una criatura que está luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    Call DoDomar(Userindex, tN)
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(Userindex, "¡No hay ninguna criatura allí!", FontTypeNames.FONTTYPE_INFO)
            End If

        Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
            'Check interval
            If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub

            'Check there is a proper item there
            If .flags.TargetObj > 0 Then
                If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then
                    'Validate other items
                    If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > .CurrentInventorySlots Then
                        Exit Sub
                    End If

                    ''chequeamos que no se zarpe duplicando oro
                    If .Invent.Object(.flags.TargetObjInvSlot).objindex <> .flags.TargetObjInvIndex Then
                        If .Invent.Object(.flags.TargetObjInvSlot).objindex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).Amount = 0 Then
                            Call WriteConsoleMsg(Userindex, "No tienes más minerales.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If

                        ''FUISTE
                        Call WriteErrorMsg(Userindex, "Has sido expulsado por el sistema anti cheats.")
                        Call FlushBuffer(Userindex)
                        Call CloseSocket(Userindex)
                        Exit Sub
                    End If
                    Call FundirMineral(Userindex)
                Else
                    Call WriteConsoleMsg(Userindex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(Userindex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
            End If

        Case eSkill.herreria
            'Target wehatever is in that tile
            Call LookatTile(Userindex, .Pos.map, X, Y)

            If .flags.TargetObj > 0 Then
                If ObjData(.flags.TargetObj).OBJType = eOBJType.otYunque Then
                    Call EnivarArmasConstruibles(Userindex)
                    Call EnivarArmadurasConstruibles(Userindex)
                    Call WriteShowBlacksmithForm(Userindex)
                Else
                    Call WriteConsoleMsg(Userindex, "Ahí no hay ningún yunque.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(Userindex, "Ahí no hay ningún yunque.", FontTypeNames.FONTTYPE_INFO)
            End If
        End Select
    End With
End Sub

''
' Handles the "CreateNewGuild" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateNewGuild(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/11/09
'05/11/09: Pato - Ahora se quitan los espacios del principio y del fin del nombre del clan
'***************************************************
    If UserList(Userindex).incomingData.length < 9 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim desc As String
        Dim GuildName As String
        Dim site As String
        Dim codex() As String
        Dim errorStr As String

        desc = buffer.ReadASCIIString()
        GuildName = Trim$(buffer.ReadASCIIString())
        site = buffer.ReadASCIIString()
        codex = Split(buffer.ReadASCIIString(), SEPARATOR)

        If modGuilds.CrearNuevoClan(Userindex, desc, GuildName, site, codex, .FundandoGuildAlineacion, errorStr) Then
            Call SendData(SendTarget.ToAll, Userindex, PrepareMessageConsoleMsg(.Name & " fundó el clan " & GuildName & " de alineación " & modGuilds.GuildAlignment(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
            .Stats.Gld = .Stats.Gld - 25000000
            WriteUpdateGold Userindex
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))


            'Update tag
            Call RefreshCharStatus(Userindex)
        Else
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub
''
' Handles the "SpellInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpellInfo(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim spellSlot As Byte
        Dim spell As Integer

        spellSlot = .incomingData.ReadByte()

        'Validate slot
        If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(Userindex, "¡Primero selecciona el hechizo!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Validate spell in the slot
        spell = .Stats.UserHechizos(spellSlot)
        If spell > 0 And spell < NumeroHechizos + 1 Then
            With Hechizos(spell)
                'Send information
                Call WriteConsoleMsg(Userindex, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & vbCrLf _
                                                & "Nombre:" & .Nombre & vbCrLf _
                                                & "Descripción:" & .desc & vbCrLf _
                                                & "Skill requerido: " & .MinSkill & " de magia." & vbCrLf _
                                                & "Maná necesario: " & .ManaRequerido & vbCrLf _
                                                & "Energía necesaria: " & .StaRequerido & vbCrLf _
                                                & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%", FontTypeNames.FONTTYPE_INFO)
            End With
        End If
    End With


End Sub

''
' Handles the "EquipItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim itemSlot As Byte

        itemSlot = .incomingData.ReadByte()

        'Dead users can't equip items
        If .flags.Muerto = 1 Then Exit Sub

        'Validate item slot
        If itemSlot > .CurrentInventorySlots Or itemSlot < 1 Then Exit Sub

        If .Invent.Object(itemSlot).objindex = 0 Then Exit Sub

        Call EquiparInvItem(Userindex, itemSlot)
    End With
End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 06/28/2008
'Last Modified By: NicoNZ
' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
' 06/28/2008: NicoNZ - Sólo se puede cambiar si está inmovilizado.
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Heading As eHeading
        Dim posX As Integer
        Dim posY As Integer

        Heading = .incomingData.ReadByte()

        If .flags.Paralizado = 1 And .flags.Inmovilizado = 0 Then
            Select Case Heading
            Case eHeading.NORTH
                posY = -1
            Case eHeading.EAST
                posX = 1
            Case eHeading.SOUTH
                posY = 1
            Case eHeading.WEST
                posX = -1
            End Select

            If LegalPos(.Pos.map, .Pos.X + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then
                Exit Sub
            End If
        End If

        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)

        If Heading > 0 And Heading < 5 Then
            .Char.Heading = Heading

            SendData SendTarget.ToPCArea, Userindex, PrepareMessageChangeHeading(.Char.CharIndex, Heading)
        End If
    End With
End Sub

''
' Handles the "ModifySkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleModifySkills(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 11/19/09
'11/19/09: Pato - Adapting to new skills system.
'***************************************************
    If UserList(Userindex).incomingData.length < 1 + NUMSKILLS Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim i  As Long
        Dim Count As Integer
        Dim Points(1 To NUMSKILLS) As Byte

        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 1 To NUMSKILLS
            Points(i) = .incomingData.ReadByte()

            If Points(i) < 0 Then
                Call LogHackAttemp(.Name & " IP:" & .ip & " trató de hackear los skills.")
                .Stats.SkillPts = 0
                Call CloseSocket(Userindex)
                Exit Sub
            End If

            Count = Count + Points(i)
        Next i

        If Count > .Stats.SkillPts Then
            Call LogHackAttemp(.Name & " IP:" & .ip & " trató de hackear los skills.")
            Call CloseSocket(Userindex)
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

                    Call CheckEluSkill(Userindex, i, True)
                End If
            Next i
        End With
    End With
End Sub

''
' Handles the "Train" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim SpawnedNpc As Integer
        Dim PetIndex As Byte

        PetIndex = .incomingData.ReadByte()

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
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
        End If
    End With
End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Slot As Byte
        Dim Amount As Integer

        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()

        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub

        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If

        'Only if in commerce mode....
        If Not .flags.Comerciando Then
            Call WriteConsoleMsg(Userindex, "No estás comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'User compra el item
        Call Comercio(eModoComercio.Compra, Userindex, .flags.TargetNPC, Slot, Amount)
    End With
End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Slot As Byte
        Dim Amount As Integer

        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()

        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub

        '¿Es el banquero?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub
        End If

        'User retira el item del slot
        Call UserRetiraItem(Userindex, Slot, Amount)
    End With
End Sub

''
' Handles the "CommerceSell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Slot As Byte
        Dim Amount As Integer

        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()

        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub

        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If

        'User compra el item del slot
        Call Comercio(eModoComercio.Venta, Userindex, .flags.TargetNPC, Slot, Amount)
    End With
End Sub

''
' Handles the "BankDeposit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Slot As Byte
        Dim Amount As Integer

        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()

        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub

        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub
        End If

        'User deposita el item del slot rdata
        Call UserDepositaItem(Userindex, Slot, Amount)
    End With
End Sub

''
' Handles the "ForumPost" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForumPost(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 02/01/2010
'02/01/2010: ZaMa - Implemento nuevo sistema de foros
'***************************************************
    If UserList(Userindex).incomingData.length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim ForumMsgType As eForumMsgType

        Dim File As String
        Dim Title As String
        Dim Post As String
        Dim ForumIndex As Integer
        Dim postFile As String
        Dim ForumType As Byte

        ForumMsgType = buffer.ReadByte()

        Title = buffer.ReadASCIIString()
        Post = buffer.ReadASCIIString()

        If .flags.TargetObj > 0 Then
            ForumType = ForumAlignment(ForumMsgType)

            Select Case ForumType

            Case eForumType.ieGeneral
                ForumIndex = GetForumIndex(ObjData(.flags.TargetObj).ForoID)

            Case eForumType.ieREAL
                ForumIndex = GetForumIndex(FORO_REAL_ID)

            Case eForumType.ieCAOS
                ForumIndex = GetForumIndex(FORO_CAOS_ID)

            End Select

            Call AddPost(ForumIndex, Post, .Name, Title, EsAnuncio(ForumMsgType))
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "MoveSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex).incomingData
        'Remove packet ID
        Call .ReadByte

        Dim dir As Integer

        If .ReadBoolean() Then
            dir = 1
        Else
            dir = -1
        End If

        Call DesplazarHechizo(Userindex, dir, .ReadByte())
    End With
End Sub

''
' Handles the "MoveBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveBank(ByVal Userindex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 06/14/09
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex).incomingData
        'Remove packet ID
        Call .ReadByte

        Dim dir As Integer
        Dim Slot As Byte
        Dim TempItem As Obj

        If .ReadBoolean() Then
            dir = 1
        Else
            dir = -1
        End If

        Slot = .ReadByte()
    End With

    With UserList(Userindex)
        TempItem.objindex = .BancoInvent.Object(Slot).objindex
        TempItem.Amount = .BancoInvent.Object(Slot).Amount

        If dir = 1 Then    'Mover arriba
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot - 1)
            .BancoInvent.Object(Slot - 1).objindex = TempItem.objindex
            .BancoInvent.Object(Slot - 1).Amount = TempItem.Amount
        Else    'mover abajo
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot + 1)
            .BancoInvent.Object(Slot + 1).objindex = TempItem.objindex
            .BancoInvent.Object(Slot + 1).Amount = TempItem.Amount
        End If
    End With

    Call UpdateBanUserInv(True, Userindex, 0)
    Call UpdateVentanaBanco(Userindex)

End Sub

''
' Handles the "ClanCodexUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleClanCodexUpdate(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim desc As String
        Dim codex() As String

        desc = buffer.ReadASCIIString()
        codex = Split(buffer.ReadASCIIString(), SEPARATOR)

        Call modGuilds.ChangeCodexAndDesc(desc, codex, .GuildIndex)

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 24/11/2009
'24/11/2009: ZaMa - Nuevo sistema de comercio
'***************************************************
    If UserList(Userindex).incomingData.length < 7 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        Dim Slot As Byte
        Dim tUser As Integer
        Dim OfferSlot As Byte
        Dim objindex As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadLong()
        OfferSlot = .incomingData.ReadByte()
        
        'Get the other player
        tUser = .ComUsu.DestUsu
        
        ' If he's already confirmed his offer, but now tries to change it, then he's cheating
        If UserList(Userindex).ComUsu.Confirmo = True Then
            
            ' Finish the trade
            Call FinComerciarUsu(Userindex)
        
            If tUser <= 0 Or tUser > MaxUsers Then
                Call FinComerciarUsu(tUser)
                Call Protocol.FlushBuffer(tUser)
            End If
        
            Exit Sub
        End If
        
        'If slot is invalid and it's not gold or it's not 0 (Substracting), then ignore it.
        If ((Slot < 0 Or Slot > UserList(Userindex).CurrentInventorySlots) And Slot <> FLAGORO) Then Exit Sub
        
        'If OfferSlot is invalid, then ignore it.
        If OfferSlot < 1 Or OfferSlot > MAX_OFFER_SLOTS + 1 Then Exit Sub
        
        ' Can be negative if substracted from the offer, but never 0.
        If Amount = 0 Then Exit Sub
        
        'Has he got enough??
        If Slot = FLAGORO Then
            ' Can't offer more than he has
            If Amount > .Stats.Gld - .ComUsu.GoldAmount Then
                Call WriteCommerceChat(Userindex, "No tienes esa cantidad de oro para agregar a la oferta.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
            
            If Amount < 0 Then
                If Abs(Amount) > .ComUsu.GoldAmount Then
                    Amount = .ComUsu.GoldAmount * (-1)
                End If
            End If
        Else
            'If modifing a filled offerSlot, we already got the objIndex, then we don't need to know it
            If Slot <> 0 Then objindex = .Invent.Object(Slot).objindex
            ' Can't offer more than he has
            If Not HasEnoughItems(Userindex, objindex, _
                TotalOfferItems(objindex, Userindex) + Amount) Then
                
                Call WriteCommerceChat(Userindex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
            
            If Amount < 0 Then
                If Abs(Amount) > .ComUsu.cant(OfferSlot) Then
                    Amount = .ComUsu.cant(OfferSlot) * (-1)
                End If
            End If
        
            If ItemNewbie(objindex) Then
                Call WriteCancelOfferItem(Userindex, OfferSlot)
                Exit Sub
            End If
            
            'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
            If .flags.Navegando = 1 Then
                If .Invent.BarcoSlot = Slot Then
                    Call WriteCommerceChat(Userindex, "No puedes vender tu barco mientras lo estés usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            
            If .Invent.MochilaEqpSlot > 0 Then
                If .Invent.MochilaEqpSlot = Slot Then
                    Call WriteCommerceChat(Userindex, "No puedes vender tu mochila mientras la estés usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
        End If
        
        Call AgregarOferta(Userindex, OfferSlot, objindex, Amount, Slot = FLAGORO)
        Call EnviarOferta(tUser, OfferSlot)
    End With
End Sub

''
' Handles the "GuildAcceptPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptPeace(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String

        guild = buffer.ReadASCIIString()

        otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(Userindex, guild, errorStr)

        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildRejectAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectAlliance(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String

        guild = buffer.ReadASCIIString()

        otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(Userindex, guild, errorStr)

        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de alianza de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", FontTypeNames.FONTTYPE_GUILD))
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildRejectPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectPeace(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String

        guild = buffer.ReadASCIIString()

        otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(Userindex, guild, errorStr)

        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de paz de " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", FontTypeNames.FONTTYPE_GUILD))
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildAcceptAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptAlliance(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String

        guild = buffer.ReadASCIIString()

        otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(Userindex, guild, errorStr)

        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la alianza con " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildOfferPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferPeace(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim guild As String
        Dim proposal As String
        Dim errorStr As String

        guild = buffer.ReadASCIIString()
        proposal = buffer.ReadASCIIString()

        If modGuilds.r_ClanGeneraPropuesta(Userindex, guild, RELACIONES_GUILD.PAZ, proposal, errorStr) Then
            Call WriteConsoleMsg(Userindex, "Propuesta de paz enviada.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildOfferAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferAlliance(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim guild As String
        Dim proposal As String
        Dim errorStr As String

        guild = buffer.ReadASCIIString()
        proposal = buffer.ReadASCIIString()

        If modGuilds.r_ClanGeneraPropuesta(Userindex, guild, RELACIONES_GUILD.ALIADOS, proposal, errorStr) Then
            Call WriteConsoleMsg(Userindex, "Propuesta de alianza enviada.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildAllianceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAllianceDetails(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim guild As String
        Dim errorStr As String
        Dim details As String

        guild = buffer.ReadASCIIString()

        details = modGuilds.r_VerPropuesta(Userindex, guild, RELACIONES_GUILD.ALIADOS, errorStr)

        If LenB(details) = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(Userindex, details)
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildPeaceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeaceDetails(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim guild As String
        Dim errorStr As String
        Dim details As String

        guild = buffer.ReadASCIIString()

        details = modGuilds.r_VerPropuesta(Userindex, guild, RELACIONES_GUILD.PAZ, errorStr)

        If LenB(details) = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(Userindex, details)
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildRequestJoinerInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestJoinerInfo(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim User As String
        Dim details As String

        User = buffer.ReadASCIIString()

        details = modGuilds.a_DetallesAspirante(Userindex, User)

        If LenB(details) = 0 Then
            Call WriteConsoleMsg(Userindex, "El personaje no ha mandado solicitud, o no estás habilitado para verla.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteShowUserRequest(Userindex, details)
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildAlliancePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAlliancePropList(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    Call WriteAlianceProposalsList(Userindex, r_ListaDePropuestas(Userindex, RELACIONES_GUILD.ALIADOS))
End Sub

''
' Handles the "GuildPeacePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeacePropList(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    Call WritePeaceProposalsList(Userindex, r_ListaDePropuestas(Userindex, RELACIONES_GUILD.PAZ))
End Sub

''
' Handles the "GuildDeclareWar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildDeclareWar(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim guild As String
        Dim errorStr As String
        Dim otherGuildIndex As Integer

        guild = buffer.ReadASCIIString()

        otherGuildIndex = modGuilds.r_DeclararGuerra(Userindex, guild, errorStr)

        If otherGuildIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            'WAR shall be!
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("TU CLAN HA ENTRADO EN GUERRA CON " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " LE DECLARA LA GUERRA A TU CLAN.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildNewWebsite" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildNewWebsite(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Call modGuilds.ActualizarWebSite(Userindex, buffer.ReadASCIIString())

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildAcceptNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptNewMember(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim errorStr As String
        Dim UserName As String
        Dim tUser As Integer

        UserName = buffer.ReadASCIIString()

        If Not modGuilds.a_AceptarAspirante(Userindex, UserName, errorStr) Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                Call modGuilds.m_ConectarMiembroAClan(tUser, .GuildIndex)
                Call RefreshCharStatus(tUser)
            End If

            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido aceptado como miembro del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildRejectNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectNewMember(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'
'***************************************************
    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim errorStr As String
        Dim UserName As String
        Dim Reason As String
        Dim tUser As Integer

        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()

        If Not modGuilds.a_RechazarAspirante(Userindex, UserName, errorStr) Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)

            If tUser > 0 Then
                Call WriteConsoleMsg(tUser, errorStr & " : " & Reason, FontTypeNames.FONTTYPE_GUILD)
            Else
                'hay que grabar en el char su rechazo
                Call modGuilds.a_RechazarAspiranteChar(UserName, .GuildIndex, Reason)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildKickMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildKickMember(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim GuildIndex As Integer

        UserName = buffer.ReadASCIIString()

        GuildIndex = modGuilds.m_EcharMiembroDeClan(Userindex, UserName)

        If GuildIndex > 0 Then
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " fue expulsado del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        Else
            Call WriteConsoleMsg(Userindex, "No puedes expulsar ese personaje del clan.", FontTypeNames.FONTTYPE_GUILD)
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildUpdateNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildUpdateNews(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Call modGuilds.ActualizarNoticias(Userindex, buffer.ReadASCIIString())

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildMemberInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberInfo(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Call modGuilds.SendDetallesPersonaje(Userindex, buffer.ReadASCIIString())

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildOpenElections" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOpenElections(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim error As String

        If Not modGuilds.v_AbrirElecciones(Userindex, error) Then
            Call WriteConsoleMsg(Userindex, error, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("¡Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & .Name, FontTypeNames.FONTTYPE_GUILD))
        End If
    End With
End Sub

''
' Handles the "GuildRequestMembership" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestMembership(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim guild As String
        Dim application As String
        Dim errorStr As String

        guild = buffer.ReadASCIIString()
        application = buffer.ReadASCIIString()

        If Not modGuilds.a_NuevoAspirante(Userindex, guild, application, errorStr) Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(Userindex, "Tu solicitud ha sido enviada. Espera prontas noticias del líder de " & guild & ".", FontTypeNames.FONTTYPE_GUILD)
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildRequestDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestDetails(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Call modGuilds.SendGuildDetails(Userindex, buffer.ReadASCIIString())

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub
Private Sub HandleOnline(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 27/01/2010 (JoaCo)
'mandamos la lista entera de nombres
'***************************************************
    Dim i As Long
    Dim Count As Long
     Dim list As String
   
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
       
         For i = 1 To LastUser
            If LenB(UserList(i).Name) <> 0 Then
                If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then _
                    Count = Count + 1
            End If
        Next i
       
        If Count > 0 Then
            WriteConsoleMsg Userindex, "Personajes jugando " & CStr(Count) & ". Record " & Declaraciones.recordusuarios, FontTypeNames.FONTTYPE_INFO
        Else
            Call WriteConsoleMsg(Userindex, "No hay usuarios Online.", FontTypeNames.FONTTYPE_INFO)
        End If
       
    End With
End Sub

''
' Handles the "Quit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ)
'If user is invisible, it automatically becomes
'visible before doing the countdown to exit
'04/15/2008 - No se reseteaban lso contadores de invi ni de ocultar. (NicoNZ)
'***************************************************
    Dim tUser  As Integer
    Dim isNotVisible As Boolean

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.automatico = True Then
            Call WriteConsoleMsg(Userindex, "No puedes salir estando en un evento.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If

        If .flags.Plantico = True Then
            Call WriteConsoleMsg(Userindex, "No puedes salir estando en un evento.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If

        If .flags.Paralizado = 1 Then
            Call WriteConsoleMsg(Userindex, "No puedes salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If

        If .flags.Montando = 1 Then
            Call WriteConsoleMsg(Userindex, "No puedes salir mientras te encuentres montando.", FontTypeNames.FONTTYPE_CONSEJOVesA)
            Exit Sub
        End If

        'exit secure commerce
        If .ComUsu.DestUsu > 0 Then
            tUser = .ComUsu.DestUsu

            If UserList(tUser).flags.UserLogged Then
                If UserList(tUser).ComUsu.DestUsu = Userindex Then
                    Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_GUILD)
                    Call FinComerciarUsu(tUser)
                End If
            End If

            Call WriteConsoleMsg(Userindex, "Comercio cancelado.", FontTypeNames.FONTTYPE_GUILD)
            Call FinComerciarUsu(Userindex)
        End If

        Call Cerrar_Usuario(Userindex)
    End With
End Sub

''
' Handles the "GuildLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildLeave(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim GuildIndex As Integer

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'obtengo el guildindex
        GuildIndex = m_EcharMiembroDeClan(Userindex, .Name)

        If GuildIndex > 0 Then
            Call WriteConsoleMsg(Userindex, "Dejas el clan.", FontTypeNames.FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(.Name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
        Else
            Call WriteConsoleMsg(Userindex, "Tú no puedes salir de este clan.", FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub

''
' Handles the "RequestAccountState" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAccountState(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim earnings As Integer
    Dim Percentage As Integer

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead people can't check their accounts
        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
            Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Select Case Npclist(.flags.TargetNPC).NPCtype
        Case eNPCType.Banquero
            Call WriteChatOverHead(Userindex, "Tienes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

        Case eNPCType.Timbero
            If Not .flags.Privilegios And PlayerType.User Then
                earnings = Apuestas.Ganancias - Apuestas.Perdidas

                If earnings >= 0 And Apuestas.Ganancias <> 0 Then
                    Percentage = Int(earnings * 100 / Apuestas.Ganancias)
                End If

                If earnings < 0 And Apuestas.Perdidas <> 0 Then
                    Percentage = Int(earnings * 100 / Apuestas.Perdidas)
                End If

                Call WriteConsoleMsg(Userindex, "Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & earnings & " (" & Percentage & "%) Jugadas: " & Apuestas.Jugadas, FontTypeNames.FONTTYPE_INFO)
            End If
        End Select
    End With
End Sub

''
' Handles the "PetStand" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetStand(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead people can't use pets
        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
            Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Make sure it's his pet
        If Npclist(.flags.TargetNPC).MaestroUser <> Userindex Then Exit Sub

        'Do it!
        Npclist(.flags.TargetNPC).Movement = TipoAI.ESTATICO

        Call Expresar(.flags.TargetNPC, Userindex)
    End With
End Sub

''
' Handles the "PetFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetFollow(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
            Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> Userindex Then Exit Sub

        'Do it
        Call FollowAmo(.flags.TargetNPC)

        Call Expresar(.flags.TargetNPC, Userindex)
    End With
End Sub


''
' Handles the "ReleasePet" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReleasePet(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 18/11/2009
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
            Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> Userindex Then Exit Sub

        'Do it
        Call QuitarPet(Userindex, .flags.TargetNPC)

    End With
End Sub

''
' Handles the "TrainList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrainList(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
            Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Make sure it's the trainer
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub

        Call WriteTrainerCreatureList(Userindex, .flags.TargetNPC)
    End With
End Sub

''
' Handles the "Rest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRest(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!! Solo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If HayOBJarea(.Pos, FOGATA) Then
            Call WriteRestOK(Userindex)

            If Not .flags.Descansar Then
                Call WriteConsoleMsg(Userindex, "Te acomodás junto a la fogata y comienzas a descansar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
            End If

            .flags.Descansar = Not .flags.Descansar
        Else
            If .flags.Descansar Then
                Call WriteRestOK(Userindex)
                Call WriteConsoleMsg(Userindex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)

                .flags.Descansar = False
                Exit Sub
            End If

            Call WriteConsoleMsg(Userindex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub
Private Sub HandleFianzah(ByVal Userindex As Integer)
'***************************************************
'Author: Matías Ezequiel
'Last Modification: 16/03/2016 by DS
'Sistema de fianzas TDS.
'***************************************************
    Dim Fianza As Long

    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        Call .incomingData.ReadByte
        Fianza = .incomingData.ReadLong


        If Not UserList(Userindex).Pos.map = 1 Then
            Call WriteConsoleMsg(Userindex, "No puedes pagar la fianza si no estás en Ullathorpe.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        ' @@ Rezniaq bronza
        If .flags.Muerto Then Call WriteConsoleMsg(Userindex, "Estás muerto.", FontTypeNames.FONTTYPE_INFO): Exit Sub

        If Not criminal(Userindex) Then Call WriteConsoleMsg(Userindex, "Ya eres ciudadano, no podrás realizar la fianza.", FontTypeNames.FONTTYPE_INFO): Exit Sub



        If Fianza <= 0 Then
            '  Call WriteConsoleMsg(UserIndex, "El minimo de fianza es 1", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        ElseIf (Fianza * 25) > .Stats.Gld Then
            Call WriteConsoleMsg(Userindex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        .Reputacion.NobleRep = .Reputacion.NobleRep + Fianza
        .Stats.Gld = .Stats.Gld - Fianza * 25

        Call WriteConsoleMsg(Userindex, "Has ganado " & Fianza & " puntos de noble.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(Userindex, "Se te han descontado " & Fianza * 25 & " monedas de oro", FontTypeNames.FONTTYPE_INFO)
    End With
    Call WriteUpdateGold(Userindex)
    Call RefreshCharStatus(Userindex)
End Sub

''
' Handles the "Meditate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMeditate(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/08 (NicoNZ)
'Arreglé un bug que mandaba un index de la meditacion diferente
'al que decia el server.
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!! Sólo puedes meditar cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Can he meditate?
        If .Stats.MaxMAN = 0 Then
            Call WriteConsoleMsg(Userindex, "Sólo las clases mágicas conocen el arte de la meditación.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Admins don't have to wait :D
        If Not .flags.Privilegios And PlayerType.User Then
            .Stats.MinMAN = .Stats.MaxMAN
            Call WriteUpdateMana(Userindex)
            Call WriteUpdateFollow(Userindex)
            Exit Sub
        End If

        Call WriteMeditateToggle(Userindex)

        If .flags.Meditando Then _
           Call WriteConsoleMsg(Userindex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)

        .flags.Meditando = Not .flags.Meditando

        'Barrin 3/10/03 Tiempo de inicio al meditar
        If .flags.Meditando Then

            .Char.loops = INFINITE_LOOPS

            'Show proper FX according to level
            If .Stats.ELV < 15 Then
                .Char.FX = FXIDs.FXMEDITARCHICO

                'ElseIf .Stats.ELV < 25 Then
                '     .Char.FX = FXIDs.FXMEDITARMEDIANO

            ElseIf .Stats.ELV < 30 Then
                .Char.FX = FXIDs.FXMEDITARMEDIANO

            ElseIf .Stats.ELV < 45 Then
                .Char.FX = FXIDs.FXMEDITARGRANDE

            Else
                .Char.FX = FXIDs.FXMEDITARXXGRANDE
            End If
            
            If .flags.IsDios Then
                .Char.FX = 29
            End If
            
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS))
        Else
            '.Counters.bPuedeMeditar = False

            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
        End If
    End With
End Sub

''
' Handles the "Resucitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Validate NPC and make sure player is dead
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor _
            And (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(Userindex))) _
            Or .flags.Muerto = 0 Then Exit Sub

        'Make sure it's close enough
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "El sacerdote no puede resucitarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Call RevivirUsuario(Userindex)
        Call WriteConsoleMsg(Userindex, "¡¡Has sido resucitado!!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Consultation" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleConsultation(ByVal Userindex As String)
'***************************************************
'Author: ZaMa
'Last Modification: 01/05/2010
'Habilita/Deshabilita el modo consulta.
'01/05/2010: ZaMa - Agrego validaciones.
'***************************************************

    Dim UserConsulta As Integer

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        ' Comando exclusivo para gms
        If Not EsGM(Userindex) Then Exit Sub

        UserConsulta = .flags.TargetUser

        'Se asegura que el target es un usuario
        If UserConsulta = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un usuario, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        ' No podes ponerte a vos mismo en modo consulta.
        If UserConsulta = Userindex Then Exit Sub

        ' No podes estra en consulta con otro gm
        If EsGM(UserConsulta) Then
            Call WriteConsoleMsg(Userindex, "No puedes iniciar el modo consulta con otro administrador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Dim UserName As String
        UserName = UserList(UserConsulta).Name

        ' Si ya estaba en consulta, termina la consulta
        If UserList(UserConsulta).flags.EnConsulta Then
            Call WriteConsoleMsg(Userindex, "Has terminado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has terminado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.Name, "Termino consulta con " & UserName)

            UserList(UserConsulta).flags.EnConsulta = False

            ' Sino la inicia
        Else
            Call WriteConsoleMsg(Userindex, "Has iniciado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has iniciado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.Name, "Inicio consulta con " & UserName)

            With UserList(UserConsulta)
                .flags.EnConsulta = True

                ' Pierde invi u ocu
                If .flags.invisible = 1 Or .flags.Oculto = 1 Then
                    .flags.Oculto = 0
                    .flags.invisible = 0
                    .Counters.TiempoOculto = 0
                    .Counters.Invisibilidad = 0

                    Call UsUaRiOs.SetInvisible(UserConsulta, UserList(UserConsulta).Char.CharIndex, False)
                End If
            End With
        End If

        Call UsUaRiOs.SetConsulatMode(UserConsulta)
    End With

End Sub

''
' Handles the "Heal" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHeal(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor _
            And Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie) _
            Or .flags.Muerto <> 0 Then Exit Sub

        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            'Call WriteConsoleMsg(Userindex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Call WriteShortMsj(Userindex, 7, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        .Stats.MinHp = .Stats.MaxHp

        Call WriteUpdateHP(Userindex)
        Call WriteUpdateFollow(Userindex)

        Call WriteConsoleMsg(Userindex, "¡¡Has sido curado!!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "RequestStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestStats(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    Call SendUserStatsTxt(Userindex, Userindex)
End Sub

''
' Handles the "Help" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHelp(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    Call SendHelp(Userindex)
End Sub

''
' Handles the "CommerceStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i      As Integer
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Not UserList(Userindex).Pos.map = 200 Then
            WriteConsoleMsg Userindex, "Para comerciar debes encontrarte en la Zona de Comercio.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If

        If .flags.Envenenado = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás envenenado!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If MapInfo(.Pos.map).Pk = True Then
            Call WriteConsoleMsg(Userindex, "¡Para poder comerciar debes estar en una ciudad segura!", FONTTYPE_INFO)
            Exit Sub
        End If

        'Is it already in commerce mode??
        If .flags.Comerciando Then
            Call WriteConsoleMsg(Userindex, "Ya estás comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            'Does the NPC want to trade??
            If Npclist(.flags.TargetNPC).Comercia = 0 Then
                If LenB(Npclist(.flags.TargetNPC).desc) <> 0 Then
                    Call WriteChatOverHead(Userindex, "No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                End If

                Exit Sub
            End If

            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            'Start commerce....
            Call IniciarComercioNPC(Userindex)
            '[Alejo]
        ElseIf .flags.TargetUser > 0 Then
            'User commerce...
            'Can he commerce??
            If .flags.Privilegios And PlayerType.Consejero Then
                Call WriteConsoleMsg(Userindex, "No puedes vender ítems.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If

            'Is the other one dead??
            If UserList(.flags.TargetUser).flags.Muerto = 1 Then
                Call WriteConsoleMsg(Userindex, "¡¡No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            'Is it me??
            If .flags.TargetUser = Userindex Then
                Call WriteConsoleMsg(Userindex, "¡¡No puedes comerciar con vos mismo!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            'Check distance
            If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            'Is he already trading?? is it with me or someone else??
            If UserList(.flags.TargetUser).flags.Comerciando = True And _
               UserList(.flags.TargetUser).ComUsu.DestUsu <> Userindex Then
                Call WriteConsoleMsg(Userindex, "No puedes comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
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

            .ComUsu.Acepto = False
            .ComUsu.Confirmo = False

            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(Userindex, .flags.TargetUser)
        Else
            Call WriteConsoleMsg(Userindex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "BankStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankStart(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If .flags.Comerciando Then
            Call WriteConsoleMsg(Userindex, "Ya estás comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            'If it's the banker....
            If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                Call IniciarDeposito(Userindex)
            End If
        Else
            Call WriteConsoleMsg(Userindex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Enlist" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnlist(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
           Or .flags.Muerto <> 0 Then Exit Sub

        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(Userindex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            Call EnlistarArmadaReal(Userindex)
        Else
            Call EnlistarCaos(Userindex)
        End If
    End With
End Sub

''
' Handles the "Information" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInformation(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim Matados As Integer
    Dim NextRecom As Integer
    Dim Diferencia As Integer

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
           Or .flags.Muerto <> 0 Then Exit Sub

        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
            Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If


        NextRecom = .Faccion.NextRecompensa

        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            If .Faccion.ArmadaReal = 0 Then
                Call WriteChatOverHead(Userindex, "¡¡No perteneces a las tropas reales!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub
            End If

            Matados = .Faccion.CriminalesMatados
            Diferencia = NextRecom - Matados

            If Diferencia > 0 Then
                Call WriteChatOverHead(Userindex, "Tu deber es combatir criminales, mata " & Diferencia & " criminales más y te daré una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else
                Call WriteChatOverHead(Userindex, "Tu deber es combatir criminales, y ya has matado los suficientes como para merecerte una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            End If
        Else
            If .Faccion.FuerzasCaos = 0 Then
                Call WriteChatOverHead(Userindex, "¡¡No perteneces a la legión oscura!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub
            End If

            Matados = .Faccion.CiudadanosMatados
            Diferencia = NextRecom - Matados

            If Diferencia > 0 Then
                Call WriteChatOverHead(Userindex, "Tu deber es sembrar el caos y la desesperanza, mata " & Diferencia & " ciudadanos más y te daré una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else
                Call WriteChatOverHead(Userindex, "Tu deber es sembrar el caos y la desesperanza, y creo que estás en condiciones de merecer una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            End If
        End If
    End With
End Sub

''
' Handles the "Reward" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReward(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
           Or .flags.Muerto <> 0 Then Exit Sub

        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            If .Faccion.ArmadaReal = 0 Then
                Call WriteChatOverHead(Userindex, "¡¡No perteneces a las tropas reales!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub
            End If
            Call RecompensaArmadaReal(Userindex)
        Else
            If .Faccion.FuerzasCaos = 0 Then
                Call WriteChatOverHead(Userindex, "¡¡No perteneces a la legión oscura!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub
            End If
            Call RecompensaCaos(Userindex)
        End If
    End With
End Sub

''
' Handles the "UpTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUpTime(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/10/08
'01/10/2008 - Marcos Martinez (ByVal) - Automatic restart removed from the server along with all their assignments and varibles
'***************************************************
'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    Dim time   As Long
    Dim UpTimeStr As String

    'Get total time in seconds
    time = ((GetTickCount() And &H7FFFFFFF) - tInicioServer) \ 1000

    'Get times in dd:hh:mm:ss format
    UpTimeStr = (time Mod 60) & " segundos."
    time = time \ 60

    UpTimeStr = (time Mod 60) & " minutos, " & UpTimeStr
    time = time \ 60

    UpTimeStr = (time Mod 24) & " horas, " & UpTimeStr
    time = time \ 24

    If time = 1 Then
        UpTimeStr = time & " día, " & UpTimeStr
    Else
        UpTimeStr = time & " días, " & UpTimeStr
    End If

    Call WriteConsoleMsg(Userindex, "Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)
End Sub

''
' Handles the "PartyLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyLeave(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    Call mdParty.SalirDeParty(Userindex)
End Sub

''
' Handles the "PartyCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyCreate(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    If Not mdParty.PuedeCrearParty(Userindex) Then Exit Sub

    Call mdParty.CrearParty(Userindex)
End Sub

''
' Handles the "PartyJoin" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyJoin(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    Call mdParty.SolicitarIngresoAParty(Userindex)
    NickPjIngreso = UserList(Userindex).Name
End Sub

''
' Handles the "ShareNpc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShareNpc(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 15/04/2010
'Shares owned npcs with other user
'***************************************************

    Dim targetUserIndex As Integer
    Dim SharingUserIndex As Integer

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        ' Didn't target any user
        targetUserIndex = .flags.TargetUser
        If targetUserIndex = 0 Then Exit Sub

        ' Can't share with admins
        If EsGM(targetUserIndex) Then
            Call WriteConsoleMsg(Userindex, "No puedes compartir npcs con administradores!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        ' Pk or Caos?
        If criminal(Userindex) Then
            ' Caos can only share with other caos
            If esCaos(Userindex) Then
                If Not esCaos(targetUserIndex) Then
                    Call WriteConsoleMsg(Userindex, "Solo puedes compartir npcs con miembros de tu misma facción!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                ' Pks don't need to share with anyone
            Else
                Exit Sub
            End If

            ' Ciuda or Army?
        Else
            ' Can't share
            If criminal(targetUserIndex) Then
                Call WriteConsoleMsg(Userindex, "No puedes compartir npcs con criminales!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If

        ' Already sharing with target
        SharingUserIndex = .flags.ShareNpcWith
        If SharingUserIndex = targetUserIndex Then Exit Sub

        ' Aviso al usuario anterior que dejo de compartir
        If SharingUserIndex <> 0 Then
            Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(Userindex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
        End If

        .flags.ShareNpcWith = targetUserIndex

        Call WriteConsoleMsg(targetUserIndex, .Name & " ahora comparte sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(Userindex, "Ahora compartes tus npcs con " & UserList(targetUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "StopSharingNpc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleStopSharingNpc(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 15/04/2010
'Stop Sharing owned npcs with other user
'***************************************************

    Dim SharingUserIndex As Integer

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        SharingUserIndex = .flags.ShareNpcWith

        If SharingUserIndex <> 0 Then

            ' Aviso al que compartia y al que le compartia.
            Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SharingUserIndex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)

            .flags.ShareNpcWith = 0
        End If

    End With

End Sub

''
' Handles the "Inquiry" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiry(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    ConsultaPopular.SendInfoEncuesta (Userindex)
End Sub

''
' Handles the "GuildMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 15/07/2009
'02/03/2009: ZaMa - Arreglado un indice mal pasado a la funcion de cartel de clanes overhead.
'15/07/2009: ZaMa - Now invisible admins only speak by console
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim chat As String
        Dim CanTalk As Boolean
        
        chat = buffer.ReadASCIIString()

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
        
        If LenB(chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(chat)
            
            CanTalk = True
            If .flags.SlotEvent > 0 Then
                If Events(.flags.SlotEvent).Modality = DeathMatch Then
                    CanTalk = False
                End If
            End If
            
            If CanTalk Then
                If .GuildIndex > 0 Then
                    Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & "> " & chat))
    
                    'If Not (.flags.AdminInvisible = 1) Then _
                     '    Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageChatOverHead("< " & chat & " >", .Char.CharIndex, vbYellow))
                End If
            End If
        End If

    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub
Private Sub HandlePartyMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim chat As String

        chat = buffer.ReadASCIIString()

        If LenB(chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(chat)

            Call mdParty.BroadCastParty(Userindex, chat)
            'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
            'Call SendData(SendTarget.ToPartyArea, UserIndex, UserList(UserIndex).Pos.map, "||" & vbYellow & "°< " & mid$(rData, 7) & " >°" & CStr(UserList(UserIndex).Char.CharIndex))
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "CentinelReport" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCentinelReport(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Call CentinelaCheckClave(Userindex, .incomingData.ReadInteger())
    End With
End Sub

''
' Handles the "GuildOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnline(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim onlinelist As String

        onlinelist = modGuilds.m_ListaDeMiembrosOnline(Userindex, .GuildIndex)

        If .GuildIndex <> 0 Then
            Call WriteConsoleMsg(Userindex, "Compañeros de tu clan conectados: " & onlinelist, FontTypeNames.FONTTYPE_GUILDMSG)
        Else
            Call WriteConsoleMsg(Userindex, "No pertences a ningún clan.", FontTypeNames.FONTTYPE_GUILDMSG)
        End If
    End With
End Sub

''
' Handles the "PartyOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyOnline(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    Call mdParty.OnlineParty(Userindex)
End Sub

''
' Handles the "CouncilMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler

    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim chat As String

        chat = buffer.ReadASCIIString()

        If LenB(chat) <> 0 Then
        
            'Analize chat...
            Call Statistics.ParseChat(chat)

            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                Call SendData(SendTarget.ToConsejo, Userindex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJO))
            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                Call SendData(SendTarget.ToConsejoCaos, Userindex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "RoleMasterRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoleMasterRequest(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim request As String

        request = buffer.ReadASCIIString()

        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(Userindex, "Su solicitud ha sido enviada.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.Name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GMRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMRequest(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If Not Ayuda.Existe(.Name) Then
            Call WriteConsoleMsg(Userindex, "El mensaje ha sido entregado, ahora sólo debes esperar que se desocupe algún GM.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(LCase$(.Name) & " GM: " & "Un usuario mando /GM por favor usa el comando /SHOW SOS", FontTypeNames.FONTTYPE_ADMIN))
            Call Ayuda.Push(.Name)
        Else
            Call Ayuda.Quitar(.Name)
            Call Ayuda.Push(.Name)
            Call WriteConsoleMsg(Userindex, "Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim Description As String, TmpStr As String, p As String

        Description = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
     
            If .flags.Muerto = 1 Then
                Call WriteConsoleMsg(Userindex, "No puedes cambiar la descripción estando muerto.", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not AsciiValidos(Description) Then
                    Call WriteConsoleMsg(Userindex, "La descripción tiene caracteres inválidos.", FontTypeNames.FONTTYPE_INFO)
                Else
                    .desc = Trim$(Description)
                    Call WriteConsoleMsg(Userindex, "La descripción ha cambiado.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If

    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildVote(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim vote As String
        Dim errorStr As String

        vote = buffer.ReadASCIIString()

        If Not modGuilds.v_UsuarioVota(Userindex, vote, errorStr) Then
            Call WriteConsoleMsg(Userindex, "Voto NO contabilizado: " & errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(Userindex, "Voto contabilizado.", FontTypeNames.FONTTYPE_GUILD)
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "ShowGuildNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowGuildNews(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMA
'Last Modification: 05/17/06
'
'***************************************************

    With UserList(Userindex)

        'Remove packet ID
        Call .incomingData.ReadByte

        Call modGuilds.SendGuildNews(Userindex)
    End With
End Sub

''
' Handles the "Punishments" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePunishments(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 25/08/2009
'25/08/2009: ZaMa - Now only admins can see other admins' punishment list
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim Name As String
        Dim Count As Integer

        Name = buffer.ReadASCIIString()

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


            If FileExist(CharPath & Name & ".chr", vbNormal) Then
                Count = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
                If Count = 0 Then
                    Call WriteConsoleMsg(Userindex, "No tienes penas.", FontTypeNames.FONTTYPE_INFO)
                Else
                    While Count > 0
                        Call WriteConsoleMsg(Userindex, Count & " - " & GetVar(CharPath & Name & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
                        Count = Count - 1
                    Wend
                End If


            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "ChangePassword" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangePassword(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Creation Date: 10/10/07
'Last Modified By: Rapsodius
'***************************************************
#If SeguridadAlkon Then
    If UserList(Userindex).incomingData.length < 65 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#Else
    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Dim oldPass As String
        Dim newPass As String
        Dim oldPass2 As String
        
        'Remove packet ID
        Call buffer.ReadByte
        
#If SeguridadAlkon Then
        oldPass = UCase$(buffer.ReadASCIIStringFixed(32))
        newPass = UCase$(buffer.ReadASCIIStringFixed(32))
#Else
        oldPass = UCase$(buffer.ReadASCIIString())
        newPass = UCase$(buffer.ReadASCIIString())
#End If
        
        If LenB(newPass) = 0 Then
            Call WriteConsoleMsg(Userindex, "Debes especificar una contraseña nueva, inténtalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
        Else
            oldPass2 = UCase$(GetVar(CharPath & UserList(Userindex).Name & ".chr", "INIT", "Password"))
            
            If oldPass2 <> oldPass Then
                Call WriteConsoleMsg(Userindex, "La contraseña actual proporcionada no es correcta. La contraseña no ha sido cambiada, inténtalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteVar(CharPath & UserList(Userindex).Name & ".chr", "INIT", "Password", newPass)
                Call WriteConsoleMsg(Userindex, "La contraseña fue cambiada con éxito.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub
Private Sub HandleChangePin(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Creation Date: 10/10/07
'Last Modified By: Rapsodius
'***************************************************
    #If SeguridadAlkon Then
        If UserList(Userindex).incomingData.length < 65 Then
            Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub
        End If
    #Else
        If UserList(Userindex).incomingData.length < 5 Then
            Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub
        End If
    #End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        Dim oldPin As String
        Dim newPin As String
        Dim oldPin2 As String

        'Remove packet ID
        Call buffer.ReadByte

        #If SeguridadAlkon Then
            oldPin = UCase$(buffer.ReadASCIIStringFixed(32))
            newPin = UCase$(buffer.ReadASCIIStringFixed(32))
        #Else
            oldPin = UCase$(buffer.ReadASCIIString())
            newPin = UCase$(buffer.ReadASCIIString())
        #End If

        If LenB(newPin) = 0 Then
            Call WriteConsoleMsg(Userindex, "Debes especificar una nueva clave PIN, inténtalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
        Else
            oldPin2 = UCase$(GetVar(CharPath & UserList(Userindex).Name & ".chr", "INIT", "PIN"))

            If oldPin2 <> oldPin Then
                Call WriteConsoleMsg(Userindex, "La clave Pin proporcionada no es correcta. La clave Pin no ha sido cambiada, inténtalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteVar(CharPath & UserList(Userindex).Name & ".chr", "INIT", "PIN", newPin)
                Call WriteConsoleMsg(Userindex, "La clave Pin fue cambiada con éxito.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub



''
' Handles the "Gamble" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGamble(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Amount As Integer

        Amount = .incomingData.ReadInteger()

        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
        ElseIf .flags.TargetNPC = 0 Then
            'Validate target NPC
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
        ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
            Call WriteChatOverHead(Userindex, "No tengo ningún interés en apostar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf Amount < 1 Then
            Call WriteChatOverHead(Userindex, "El mínimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf Amount > 5000 Then
            Call WriteChatOverHead(Userindex, "El máximo de apuesta es 5000 monedas.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf .Stats.Gld < Amount Then
            Call WriteChatOverHead(Userindex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            If RandomNumber(1, 100) <= 47 Then
                .Stats.Gld = .Stats.Gld + Amount
                Call WriteChatOverHead(Userindex, "¡Felicidades! Has ganado " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

                Apuestas.Perdidas = Apuestas.Perdidas + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                .Stats.Gld = .Stats.Gld - Amount
                Call WriteChatOverHead(Userindex, "Lo siento, has perdido " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

                Apuestas.Ganancias = Apuestas.Ganancias + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
            End If

            Apuestas.Jugadas = Apuestas.Jugadas + 1

            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))

            Call WriteUpdateGold(Userindex)
        End If
    End With
End Sub

''
' Handles the "InquiryVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiryVote(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim opt As Byte

        opt = .incomingData.ReadByte()

        Call WriteConsoleMsg(Userindex, ConsultaPopular.doVotar(Userindex, opt), FontTypeNames.FONTTYPE_GUILD)
    End With
End Sub
Private Sub HandleBankExtractGold(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Amount As Long

        Amount = .incomingData.ReadLong()

        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub

        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Dim Monea As Long
        Monea = .Stats.Banco

        If Amount > 0 And Amount <= .Stats.Banco Then
            .Stats.Banco = .Stats.Banco - Amount
            .Stats.Gld = .Stats.Gld + Amount
            Call WriteChatOverHead(Userindex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            .Stats.Gld = .Stats.Gld + .Stats.Banco
            .Stats.Banco = 0
            Call WriteChatOverHead(Userindex, "Has retirado " & Monea & " monedas de oro de tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If

        Call WriteUpdateGold(Userindex)
        Call WriteUpdateBankGold(Userindex)
    End With
End Sub

''
' Handles the "LeaveFaction" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeaveFaction(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim TalkToKing As Boolean
    Dim TalkToDemon As Boolean
    Dim NpcIndex As Integer

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        ' Chequea si habla con el rey o el demonio. Puede salir sin hacerlo, pero si lo hace le reponden los npcs
        NpcIndex = .flags.TargetNPC
        If NpcIndex <> 0 Then
            ' Es rey o domonio?
            If Npclist(NpcIndex).NPCtype = eNPCType.Noble Then
                'Rey?
                If Npclist(NpcIndex).flags.Faccion = 0 Then
                    TalkToKing = True
                    ' Demonio
                Else
                    TalkToDemon = True
                End If
            End If
        End If

        'Quit the Royal Army?
        If .Faccion.ArmadaReal = 1 Then
            ' Si le pidio al demonio salir de la armada, este le responde.
            If TalkToDemon Then
                Call WriteChatOverHead(Userindex, "¡¡¡Sal de aquí bufón!!!", _
                                       Npclist(NpcIndex).Char.CharIndex, vbWhite)

            Else
                ' Si le pidio al rey salir de la armada, le responde.
                If TalkToKing Then
                    Call WriteChatOverHead(Userindex, "Serás bienvenido a las fuerzas imperiales si deseas regresar.", _
                                           Npclist(NpcIndex).Char.CharIndex, vbWhite)
                End If

                Call ExpulsarFaccionReal(Userindex, False)

            End If

            'Quit the Chaos Legion?
        ElseIf .Faccion.FuerzasCaos = 1 Then
            ' Si le pidio al rey salir del caos, le responde.
            If TalkToKing Then
                Call WriteChatOverHead(Userindex, "¡¡¡Sal de aquí maldito criminal!!!", _
                                       Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Else
                ' Si le pidio al demonio salir del caos, este le responde.
                If TalkToDemon Then
                    Call WriteChatOverHead(Userindex, "Ya volverás arrastrandote.", _
                                           Npclist(NpcIndex).Char.CharIndex, vbWhite)
                End If

                Call ExpulsarFaccionCaos(Userindex, False)
            End If
            ' No es faccionario
        Else

            ' Si le hablaba al rey o demonio, le repsonden ellos
            If NpcIndex > 0 Then
                Call WriteChatOverHead(Userindex, "¡No perteneces a ninguna facción!", _
                                       Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Else
                Call WriteConsoleMsg(Userindex, "¡No perteneces a ninguna facción!", FontTypeNames.FONTTYPE_FIGHT)
            End If

        End If

    End With

End Sub
Private Sub HandleBankDepositGold(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Amount As Long

        Amount = .incomingData.ReadLong()

        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub

        If Amount > 0 And Amount <= .Stats.Gld Then
            .Stats.Banco = .Stats.Banco + Amount
            .Stats.Gld = .Stats.Gld - Amount
            Call WriteChatOverHead(Userindex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

            Call WriteUpdateGold(Userindex)
            Call WriteUpdateBankGold(Userindex)
        Else
            Call WriteChatOverHead(Userindex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
    End With
End Sub

''
' Handles the "Denounce" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim Text As String

        Text = buffer.ReadASCIIString()

        If .flags.Silenciado = 0 And .Counters.Denuncia = 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Text)

            If UCase$(Left$(Text, 15)) = "[FOTODENUNCIAS]" Then
                SendData SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(Text & ". Hecha por: " & .Name, FontTypeNames.FONTTYPE_INFO)
                .Counters.Denuncia = 20
            Else
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(LCase$(.Name) & " DENUNCIA: " & Text, FontTypeNames.fonttype_dios))
                Call WriteConsoleMsg(Userindex, "Recuerda que sólo puedes enviar una denuncia cada 30 segundos.", FontTypeNames.fonttype_dios)
                Call WriteConsoleMsg(Userindex, "Denuncia enviada, pronto será atendido por un Game Master.", FontTypeNames.FONTTYPE_INFO)
                .Counters.Denuncia = 30
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildFundate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundate(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    If UserList(Userindex).incomingData.length < 1 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        Call .incomingData.ReadByte

        If HasFound(.Name) Then
            Call WriteConsoleMsg(Userindex, "¡Ya has fundado un clan, no puedes fundar otro!", FontTypeNames.FONTTYPE_INFOBOLD)
            Exit Sub
        End If

        Call WriteShowGuildAlign(Userindex)
    End With
End Sub

''
' Handles the "GuildFundation" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundation(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim clanType As eClanType
        Dim error As String

        clanType = .incomingData.ReadByte()

        If HasFound(.Name) Then
            Call WriteConsoleMsg(Userindex, "¡Ya has fundado un clan, no puedes fundar otro!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogCheating("El usuario " & .Name & " ha intentado fundar un clan ya habiendo fundado otro desde la IP " & .ip)
            Exit Sub
        End If

        Select Case UCase$(Trim(clanType))
        Case eClanType.ct_RoyalArmy
            .FundandoGuildAlineacion = ALINEACION_ARMADA
        Case eClanType.ct_Evil
            .FundandoGuildAlineacion = ALINEACION_LEGION
        Case eClanType.ct_Neutral
            .FundandoGuildAlineacion = ALINEACION_NEUTRO
        Case eClanType.ct_GM
            .FundandoGuildAlineacion = ALINEACION_MASTER
        Case eClanType.ct_Legal
            .FundandoGuildAlineacion = ALINEACION_CIUDA
        Case eClanType.ct_Criminal
            .FundandoGuildAlineacion = ALINEACION_CRIMINAL
        Case Else
            Call WriteConsoleMsg(Userindex, "Alineación inválida.", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End Select

        If modGuilds.PuedeFundarUnClan(Userindex, .FundandoGuildAlineacion, error) Then
            Call WriteShowGuildFundationForm(Userindex)
        Else
            .FundandoGuildAlineacion = 0
            Call WriteConsoleMsg(Userindex, error, FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub

''
' Handles the "PartyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyKick(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/05/09
'Last Modification by: Marco Vanotti (Marco)
'- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer

        UserName = buffer.ReadASCIIString()

        If UserPuedeEjecutarComandos(Userindex) Then
            tUser = NameIndex(UserName)

            If tUser > 0 Then
                Call mdParty.ExpulsarDeParty(Userindex, tUser)
            Else
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                End If

                Call WriteConsoleMsg(Userindex, LCase(UserName) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "PartySetLeader" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartySetLeader(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/05/09
'Last Modification by: Marco Vanotti (MarKoxX)
'- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    'On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer
        Dim Rank As Integer
        Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

        UserName = buffer.ReadASCIIString()
        If UserPuedeEjecutarComandos(Userindex) Then
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                'Don't allow users to spoof online GMs
                If (UserDarPrivilegioLevel(UserName) And Rank) <= (.flags.Privilegios And Rank) Then
                    Call mdParty.TransformarEnLider(Userindex, tUser)
                Else
                    Call WriteConsoleMsg(Userindex, LCase(UserList(tUser).Name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
                End If

            Else
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                End If
                Call WriteConsoleMsg(Userindex, LCase(UserName) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "PartyAcceptMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyAcceptMember(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/05/09
'Last Modification by: Marco Vanotti (Marco)
'- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer
        Dim Rank As Integer
        Dim bUserVivo As Boolean

        Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

        UserName = buffer.ReadASCIIString()
        If UserList(Userindex).flags.Muerto Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_PARTY)
        Else
            bUserVivo = True
        End If

        If mdParty.UserPuedeEjecutarComandos(Userindex) And bUserVivo Then
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                'Validate administrative ranks - don't allow users to spoof online GMs
                If (UserList(tUser).flags.Privilegios And Rank) <= (.flags.Privilegios And Rank) Then
                    Call mdParty.AprobarIngresoAParty(Userindex, tUser)
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes incorporar a tu party a personajes de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                End If

                'Don't allow users to spoof online GMs
                If (UserDarPrivilegioLevel(UserName) And Rank) <= (.flags.Privilegios And Rank) Then
                    Call WriteConsoleMsg(Userindex, LCase(UserName) & " no ha solicitado ingresar a tu party.", FontTypeNames.FONTTYPE_PARTY)
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes incorporar a tu party a personajes de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GuildMemberList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberList(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim guild As String
        Dim memberCount As Integer
        Dim i  As Long
        Dim UserName As String

        guild = buffer.ReadASCIIString()

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If (InStrB(guild, "\") <> 0) Then
                guild = Replace(guild, "\", "")
            End If
            If (InStrB(guild, "/") <> 0) Then
                guild = Replace(guild, "/", "")
            End If

            If Not FileExist(App.Path & "\guilds\" & guild & "-members.mem") Then
                Call WriteConsoleMsg(Userindex, "No existe el clan: " & guild, FontTypeNames.FONTTYPE_INFO)
            Else
                memberCount = val(GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "INIT", "NroMembers"))

                For i = 1 To memberCount
                    UserName = GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "Members", "Member" & i)

                    Call WriteConsoleMsg(Userindex, UserName & "<" & guild & ">", FontTypeNames.FONTTYPE_INFO)
                Next i
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GMMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim message As String

        message = buffer.ReadASCIIString()

        Call .incomingData.CopyBuffer(buffer)

        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Mensaje a Gms:" & message)

            If LenB(message) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(message)

                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & "> " & message, FontTypeNames.FONTTYPE_GMMSG))
            End If
        End If


    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "ShowName" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            .showName = Not .showName    'Show / Hide the name

            Call RefreshCharStatus(Userindex)
        End If
    End With
End Sub

''
' Handles the "OnlineRoyalArmy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineRoyalArmy(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 28/05/2010
'28/05/2010: ZaMa - Ahora solo dioses pueden ver otros dioses online.
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        .incomingData.ReadByte

        If .flags.Privilegios And PlayerType.User Then Exit Sub

        Dim i  As Long
        Dim list As String
        Dim priv As PlayerType

        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoyalCouncil
        
        ' Solo dioses pueden ver otros dioses online
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = priv Or PlayerType.Dios Or PlayerType.Admin
        End If

        For i = 1 To LastUser
            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.ArmadaReal = 1 Then
                    If UserList(i).flags.Privilegios And priv Then
                        list = list & UserList(i).Name & ", "
                    End If
                End If
            End If
        Next i
    End With

    If Len(list) > 0 Then
        Call WriteConsoleMsg(Userindex, "Reales conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(Userindex, "No hay reales conectados.", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineChaosLegion(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 28/05/2010
'28/05/2010: ZaMa - Ahora solo dioses pueden ver otros dioses online.
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        .incomingData.ReadByte

        If .flags.Privilegios And PlayerType.User Then Exit Sub

        Dim i  As Long
        Dim list As String
        Dim priv As PlayerType

        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.ChaosCouncil

        ' Solo dioses pueden ver otros dioses online
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = priv Or PlayerType.Dios Or PlayerType.Admin
        End If

        For i = 1 To LastUser
            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.FuerzasCaos = 1 Then
                    If UserList(i).flags.Privilegios And priv Then
                        list = list & UserList(i).Name & ", "
                    End If
                End If
            End If
        Next i
    End With

    If Len(list) > 0 Then
        Call WriteConsoleMsg(Userindex, "Caos conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(Userindex, "No hay Caos conectados.", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

''
' Handles the "GoNearby" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoNearby(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/10/07
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String

        UserName = buffer.ReadASCIIString()

        Dim tIndex As Integer
        Dim X  As Long
        Dim Y  As Long
        Dim i  As Long
        Dim Found As Boolean

        tIndex = NameIndex(UserName)

        'Check the user has enough powers
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            If Not StrComp(UCase$(UserName), "THYRAH") = 0 Then
                'Si es dios o Admins no podemos salvo que nosotros también lo seamos
                If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
                    If tIndex <= 0 Then    'existe el usuario destino?
                        Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        For i = 2 To 5    'esto for sirve ir cambiando la distancia destino
                            For X = UserList(tIndex).Pos.X - i To UserList(tIndex).Pos.X + i
                                For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i
                                    If MapData(UserList(tIndex).Pos.map, X, Y).Userindex = 0 Then
                                        If LegalPos(UserList(tIndex).Pos.map, X, Y, True, True) Then
                                            Call WarpUserChar(Userindex, UserList(tIndex).Pos.map, X, Y, True)
                                            Call LogGM(.Name, "/IRCERCA " & UserName & " Mapa:" & UserList(tIndex).Pos.map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y)
                                            Found = True
                                            Exit For
                                        End If
                                    End If
                                Next Y
    
                                If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                            Next X
    
                            If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                        Next i
    
                        'No space found??
                        If Not Found Then
                            Call WriteConsoleMsg(Userindex, "Todos los lugares están ocupados.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub
Private Sub Elmasbuscado(ByVal Userindex As String)

    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte
        Dim UserName As String
        UserName = buffer.ReadASCIIString()
        Dim tIndex As String
        tIndex = NameIndex(UserName)

        If Not EsGM(Userindex) Then Exit Sub

        If tIndex <= 0 Then    'usuario Offline
            Call WriteConsoleMsg(Userindex, "Usuario Offline.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(tIndex).flags.Muerto = 1 Then    'tu enemigo esta muerto
            Call WriteConsoleMsg(Userindex, "El usuario que queres que sea buscado esta muerto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(tIndex).Pos.map = 201 Then
            Call WriteConsoleMsg(Userindex, "Esta ocupado en un reto.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Atencion!!: Se Busca el usuario " & UserList(tIndex).Name & ", el que lo asesine tendra su recompensa.", FontTypeNames.FONTTYPE_GUILD))
            Call WriteConsoleMsg(tIndex, "Tu eres el usuario más buscado, ten cuidado!!.", FontTypeNames.FONTTYPE_GUILD)
            ElmasbuscadoFusion = UserList(tIndex).Name
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error

    Exit Sub
End Sub


''
' Handles the "Comment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleComment(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim comment As String
        comment = buffer.ReadASCIIString()

        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Comentario: " & comment)
            Call WriteConsoleMsg(Userindex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "ServerTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And PlayerType.User Then Exit Sub

        Call LogGM(.Name, "Hora.")
    End With

    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & time & " " & Date, FontTypeNames.FONTTYPE_INFO))
End Sub

''
' Handles the "Where" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 18/11/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'18/11/2010: ZaMa - Obtengo los privs del charfile antes de mostrar la posicion de un usuario offline.
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer
        Dim miPos As String

        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
            

        If EsGM(Userindex) Then
            If tUser <= 0 Then
                If Not StrComp(UCase$(UserName), "THYRAH") = 0 Then
                    If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                        miPos = GetVar(CharPath & UserName & ".chr", "INIT", "POSITION")
                        Call WriteConsoleMsg(Userindex, "Ubicación  " & UserName & " (Offline): " & ReadField(1, miPos, 45) & ", " & ReadField(2, miPos, 45) & ", " & ReadField(3, miPos, 45) & ".", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Else
                    If Not StrComp(UCase$(UserList(tUser).Name), "THYRAH") = 0 Then
                        Call WriteConsoleMsg(Userindex, "Ubicación  " & UserName & ": " & UserList(tUser).Pos.map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub
''
' Handles the "CreaturesInMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 30/07/06
'Pablo (ToxicWaste): modificaciones generales para simplificar la visualización.
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim map As Integer
        Dim i, j As Long
        Dim NPCcount1, NPCcount2 As Integer
        Dim NPCcant1() As Integer
        Dim NPCcant2() As Integer
        Dim List1() As String
        Dim List2() As String

        map = .incomingData.ReadInteger()

        If .flags.Privilegios And PlayerType.User Then Exit Sub

        If MapaValido(map) Then
            For i = 1 To LastNPC
                'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                If Npclist(i).Pos.map = map Then
                    '¿esta vivo?
                    If Npclist(i).flags.NPCActive And Npclist(i).Hostile = 1 And Npclist(i).Stats.Alineacion = 2 Then
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

            Call WriteConsoleMsg(Userindex, "Npcs Hostiles en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            If NPCcount1 = 0 Then
                Call WriteConsoleMsg(Userindex, "No hay NPCS Hostiles.", FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To NPCcount1 - 1
                    Call WriteConsoleMsg(Userindex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
                Next j
            End If
            Call WriteConsoleMsg(Userindex, "Otros Npcs en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            If NPCcount2 = 0 Then
                Call WriteConsoleMsg(Userindex, "No hay más NPCS.", FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To NPCcount2 - 1
                    Call WriteConsoleMsg(Userindex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
                Next j
            End If
            Call LogGM(.Name, "Numero enemigos en mapa " & map)
        End If
    End With
End Sub

''
' Handles the "WarpMeToTarget" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpMeToTarget(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/09
'26/03/06: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim X  As Integer
        Dim Y  As Integer

        If .flags.Privilegios And PlayerType.User Then Exit Sub

        X = .flags.TargetX
        Y = .flags.TargetY

        Call FindLegalPos(Userindex, .flags.TargetMap, X, Y)
        Call WarpUserChar(Userindex, .flags.TargetMap, X, Y, True)
        Call LogGM(.Name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.map)
    End With
End Sub

''
' Handles the "WarpChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
'***************************************************
    If UserList(Userindex).incomingData.length < 7 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim map As Integer
        Dim X  As Integer
        Dim Y  As Integer
        Dim tUser As Integer

        UserName = buffer.ReadASCIIString()
        map = buffer.ReadInteger()
        X = buffer.ReadByte()
        Y = buffer.ReadByte()

        If Not .flags.Privilegios And PlayerType.User Then
            If MapaValido(map) And LenB(UserName) <> 0 Then
                If UCase$(UserName) <> "YO" Then
                    If Not .flags.Privilegios And PlayerType.Consejero Then
                        tUser = NameIndex(UserName)
                    End If
                Else
                    tUser = Userindex
                End If

                If tUser <= 0 Then
                    Call WriteVar(CharPath & UserName & ".chr", "INIT", "Position", map & "-" & X & "-" & Y)
                    Call WriteConsoleMsg(Userindex, "Charfile modificado", FontTypeNames.FONTTYPE_GM)
                ElseIf InMapBounds(map, X, Y) Then
                    Call FindLegalPos(tUser, map, X, Y)
                    Call WarpUserChar(tUser, map, X, Y, True, True)
                    If tUser <> Userindex Then Call WriteConsoleMsg(Userindex, UserList(tUser).Name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "Transportó a " & UserList(tUser).Name & " hacia " & "Mapa" & map & " X:" & X & " Y:" & Y)
                    
                    
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "Silence" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSilence(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer

        UserName = buffer.ReadASCIIString()

        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                If UserList(tUser).flags.Silenciado = 0 Then
                    UserList(tUser).flags.Silenciado = 1
                    Call WriteConsoleMsg(Userindex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteShowMessageBox(tUser, "Estimado usuario, ud. ha sido silenciado por los administradores. Sus denuncias serán ignoradas por el servidor de aquí en más. Utilice /GM para contactar un administrador.")
                    Call LogGM(.Name, "/silenciar " & UserList(tUser).Name)

                    'Flush the other user's buffer
                    Call FlushBuffer(tUser)
                Else
                    UserList(tUser).flags.Silenciado = 0
                    Call WriteConsoleMsg(Userindex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "/DESsilenciar " & UserList(tUser).Name)
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "SOSShowList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSShowList(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowSOSForm(Userindex)
    End With
End Sub

''
' Handles the "RequestPartyForm" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyForm(ByVal Userindex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        If .PartyIndex > 0 Then
            Call WriteShowPartyForm(Userindex)

        Else
            Call WriteConsoleMsg(Userindex, "No perteneces a ningún grupo!", FontTypeNames.FONTTYPE_INFOBOLD)
        End If
    End With
End Sub

''
' Handles the "ItemUpgrade" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleItemUpgrade(ByVal Userindex As Integer)
'***************************************************
'Author: Torres Patricio
'Last Modification: 12/09/09
'
'***************************************************
    With UserList(Userindex)
        Dim ItemIndex As Integer

        'Remove packet ID
        Call .incomingData.ReadByte

        ItemIndex = .incomingData.ReadInteger()

        If ItemIndex <= 0 Then Exit Sub
        If Not TieneObjetos(ItemIndex, 1, Userindex) Then Exit Sub

        Call DoUpgrade(Userindex, ItemIndex)
    End With
End Sub

''
' Handles the "SOSRemove" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSRemove(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        UserName = buffer.ReadASCIIString()

        If Not .flags.Privilegios And PlayerType.User Then _
           Call Ayuda.Quitar(UserName)

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "GoToChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer
        Dim X  As Integer
        Dim Y  As Integer

        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)

        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios) Then
            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                If Not StrComp(UCase$(UserName), "THYRAH") = 0 Or StrComp(UCase$(.Name), "LAUTARO") = 0 Then
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        If Not UserList(tUser).Pos.map = 290 Then
                            X = UserList(tUser).Pos.X
                            Y = UserList(tUser).Pos.Y + 1
                            Call FindLegalPos(Userindex, UserList(tUser).Pos.map, X, Y)
    
                            Call WarpUserChar(Userindex, UserList(tUser).Pos.map, X, Y, True)
    
                            If .flags.AdminInvisible = 0 Then
                                'Call WriteConsoleMsg(tUser, " sientes una presencia cerca de ti.", FontTypeNames.FONTTYPE_INFO)
                                Call FlushBuffer(tUser)
                            End If
    
                            Call LogGM(.Name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y)
                        End If
                    End If
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub
''
' Handles the "Invisible" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And PlayerType.User Then Exit Sub

        Call DoAdminInvisible(Userindex)
        Call LogGM(.Name, "/INVISIBLE")
    End With
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And PlayerType.User Then Exit Sub

        Call WriteShowGMPanelForm(Userindex)
    End With
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'Last modified by: Lucas Tavolaro Ortiz (Tavo)
'I haven`t found a solution to split, so i make an array of names
'***************************************************
    Dim i      As Long
    Dim names() As String
    Dim Count  As Long

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub

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

        If Count > 1 Then Call WriteUserNameList(Userindex, names(), Count - 1)
    End With
End Sub

''
' Handles the "Working" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorking(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i      As Long
    Dim Users  As String

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub

        For i = 1 To LastUser
            If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
                Users = Users & ", " & UserList(i).Name

                ' Display the user being checked by the centinel
                If modCentinela.Centinela.RevisandoUserIndex = i Then _
                   Users = Users & " (*)"
            End If
        Next i

        If LenB(Users) <> 0 Then
            Users = Right$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(Userindex, "Usuarios trabajando: " & Users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(Userindex, "No hay usuarios trabajando.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Hiding" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHiding(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i      As Long
    Dim Users  As String

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub

        For i = 1 To LastUser
            If (LenB(UserList(i).Name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
                Users = Users & UserList(i).Name & ", "
            End If
        Next i

        If LenB(Users) <> 0 Then
            Users = Left$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(Userindex, "Usuarios ocultandose: " & Users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(Userindex, "No hay usuarios ocultandose.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Jail" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleJail(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim Reason As String
        Dim jailTime As Byte
        Dim Count As Byte
        Dim tUser As Integer

        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        jailTime = buffer.ReadByte()

        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")
        End If

        '/carcel nick@motivo@<tiempo>
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(Userindex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(UserName)

                If tUser <= 0 Then
                    Call WriteConsoleMsg(Userindex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                        Call WriteConsoleMsg(Userindex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf jailTime > 60 Then
                        Call WriteConsoleMsg(Userindex, "No puedés encarcelar por más de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
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
                            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(Reason) & " " & Date & " " & time)
                        End If

                        Call Encarcelar(tUser, jailTime, .Name)
                        Call LogGM(.Name, " encarceló a " & UserName)
                    End If
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "KillNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/22/08 (NicoNZ)
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And PlayerType.User Then Exit Sub

        Dim tNpc As Integer
        Dim auxNPC As Npc

        'Los consejeros no pueden RMATAr a nada en el mapa pretoriano
        If .flags.Privilegios And PlayerType.Consejero Then
            If .Pos.map = MAPA_PRETORIANO Then
                Call WriteConsoleMsg(Userindex, "Los consejeros no pueden usar este comando en el mapa pretoriano.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If

        tNpc = .flags.TargetNPC

        If tNpc > 0 Then
            Call WriteConsoleMsg(Userindex, "RMatas (con posible respawn) a: " & Npclist(tNpc).Name, FontTypeNames.FONTTYPE_INFO)

            auxNPC = Npclist(tNpc)
            Call QuitarNPC(tNpc)
            Call ReSpawnNpc(auxNPC)

            .flags.TargetNPC = 0
        Else
            Call WriteConsoleMsg(Userindex, "Antes debes hacer click sobre el NPC.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "WarnUser" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim Reason As String
        Dim Privs As PlayerType
        Dim Count As Byte

        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And _
            (Not .flags.Privilegios And PlayerType.User) <> 0 Or _
             (.flags.Privilegios And PlayerType.ChaosCouncil Or _
             .flags.Privilegios And PlayerType.RoyalCouncil) Then
            
            If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(Userindex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Else
                Privs = UserDarPrivilegioLevel(UserName)

                If Not Privs And PlayerType.User Then
                    Call WriteConsoleMsg(Userindex, "No puedes advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
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
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": ADVERTENCIA por: " & LCase$(Reason) & " " & Date & " " & time)

                        Call WriteConsoleMsg(Userindex, "Has advertido a " & UCase$(UserName) & ".", FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.Name, " advirtio a " & UserName)
                    End If
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub


''
' Handles the "RequestCharInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInfo(ByVal Userindex As Integer)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Last Modification by: (liquid).. alto bug zapallo..
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim TargetName As String
        Dim TargetIndex As Integer

        TargetName = Replace$(buffer.ReadASCIIString(), "+", " ")
        TargetIndex = NameIndex(TargetName)


        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            'is the player offline?
            If TargetIndex <= 0 Then
                'don't allow to retrieve administrator's info
                If Not (EsDios(TargetName) Or EsAdmin(TargetName)) Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline, buscando en charfile.", FontTypeNames.FONTTYPE_INFO)
                    Call SendUserStatsTxtOFF(Userindex, TargetName)
                End If
            Else
                'don't allow to retrieve administrator's info
                If UserList(TargetIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                    Call SendUserStatsTxt(Userindex, TargetIndex)
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "RequestCharStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharStats(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer
        UserName = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call LogGM(.Name, "/STAT " & UserName)

            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_INFO)

                Call SendUserMiniStatsTxtFromChar(Userindex, UserName)
            Else
                Call SendUserMiniStatsTxt(Userindex, tUser)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "RequestCharGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharGold(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer

        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/BAL " & UserName)

            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)

                Call SendUserOROTxtFromChar(Userindex, UserName)
            Else
                Call WriteConsoleMsg(Userindex, "El usuario " & UserName & " tiene " & UserList(tUser).Stats.Banco & " en el banco.", FontTypeNames.FONTTYPE_TALK)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "RequestCharInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInventory(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer

        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)


        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/INV " & UserName)

            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline. Leyendo del charfile...", FontTypeNames.FONTTYPE_TALK)

                Call SendUserInvTxtFromChar(Userindex, UserName)
            Else
                Call SendUserInvTxt(Userindex, tUser)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "RequestCharBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharBank(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer

        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)


        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/BOV " & UserName)

            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)

                Call SendUserBovedaTxtFromChar(Userindex, UserName)
            Else
                Call SendUserBovedaTxt(Userindex, tUser)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "RequestCharSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharSkills(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Long
        Dim message As String

        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)


        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/STATS " & UserName)

            If tUser <= 0 Then
                If (InStrB(UserName, "\") <> 0) Then
                    UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                    UserName = Replace(UserName, "/", "")
                End If

                For LoopC = 1 To NUMSKILLS
                    message = message & "CHAR>" & SkillsNames(LoopC) & " = " & GetVar(CharPath & UserName & ".chr", "SKILLS", "SK" & LoopC) & vbCrLf
                Next LoopC

                Call WriteConsoleMsg(Userindex, message & "CHAR> Libres:" & GetVar(CharPath & UserName & ".chr", "STATS", "SKILLPTSLIBRES"), FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendUserSkillsTxt(Userindex, tUser)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "ReviveChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReviveChar(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 11/03/2010
'11/03/2010: ZaMa - Al revivir con el comando, si esta navegando le da cuerpo e barca.
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte

        UserName = buffer.ReadASCIIString()


        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
            Else
                tUser = Userindex
            End If

            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                With UserList(tUser)
                    'If dead, show him alive (naked).
                    If .flags.Muerto = 1 Then
                        .flags.Muerto = 0

                        If .flags.Navegando = 1 Then
                            Call ToggleBoatBody(tUser)
                        Else
                            Call DarCuerpoDesnudo(tUser)
                        End If

                        If .flags.Traveling = 1 Then
                            .flags.Traveling = 0
                            .Counters.goHome = 0
                            Call WriteMultiMessage(tUser, eMessages.CancelHome)
                        End If

                        Call ChangeUserChar(tUser, .Char.body, .OrigChar.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        Call WriteConsoleMsg(tUser, UserList(Userindex).Name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(tUser, UserList(Userindex).Name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)
                    End If

                    .Stats.MinHp = .Stats.MaxHp

                    If .flags.Traveling = 1 Then
                        .Counters.goHome = 0
                        .flags.Traveling = 0
                        Call WriteMultiMessage(tUser, eMessages.CancelHome)
                    End If

                End With

                Call WriteUpdateHP(tUser)
                Call WriteUpdateFollow(tUser)

                Call FlushBuffer(tUser)

                Call LogGM(.Name, "Resucito a " & UserName)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "OnlineGM" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineGM(ByVal Userindex As Integer)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 12/28/06
'
'***************************************************
    Dim i      As Long
    Dim list   As String
    Dim priv   As PlayerType

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        priv = PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv Or PlayerType.Dios Or PlayerType.Admin
        
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged Then
                If UserList(i).flags.Privilegios And priv And _
                    Not StrComp(UCase$(UserList(i).Name), "THYRAH") = 0 Then
                   list = list & UserList(i).Name & ", "
                
                End If
            End If
        Next i

        If LenB(list) <> 0 Then
            list = Left$(list, Len(list) - 2)
            Call WriteConsoleMsg(Userindex, list & ".", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(Userindex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "OnlineMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineMap(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 23/03/2009
'23/03/2009: ZaMa - Ahora no requiere estar en el mapa, sino que por defecto se toma en el que esta, pero se puede especificar otro
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim map As Integer
        map = .incomingData.ReadInteger

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        Dim LoopC As Long
        Dim list As String
        Dim priv As PlayerType

        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv + (PlayerType.Dios Or PlayerType.Admin)

        For LoopC = 1 To LastUser
            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).Pos.map = map Then
                If UserList(LoopC).flags.Privilegios And priv Then _
                   list = list & UserList(LoopC).Name & ", "
            End If
        Next LoopC

        If Len(list) > 2 Then list = Left$(list, Len(list) - 2)

        Call WriteConsoleMsg(Userindex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Forgive" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForgive(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer

        UserName = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)

            If tUser > 0 Then
                If EsNewbie(tUser) Then
                    Call VolverCiudadano(tUser)
                Else
                    Call LogGM(.Name, "Intento perdonar un personaje de nivel avanzado.")
                    Call WriteConsoleMsg(Userindex, "Sólo se permite perdonar newbies.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "Kick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKick(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer
        Dim Rank As Integer

        Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

        UserName = buffer.ReadASCIIString()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (UserList(tUser).flags.Privilegios And Rank) > (.flags.Privilegios And Rank) Then
                    Call WriteConsoleMsg(Userindex, "No puedes echar a alguien con jerarquía mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " echó a " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
                    Call CloseSocket(tUser)
                    Call LogGM(.Name, "Echó a " & UserName)
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "Execute" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExecute(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer

        UserName = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)

            If tUser > 0 Then
                If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                    Call WriteConsoleMsg(Userindex, "¿¿Estás loco?? ¿¿Cómo vas a piñatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call UserDie(tUser)
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " ha ejecutado a " & UserName & ".", FontTypeNames.FONTTYPE_EJECUCION))
                    Call LogGM(.Name, " ejecuto a " & UserName)
                End If
            Else
                Call WriteConsoleMsg(Userindex, "No está online.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "BanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanChar(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim Reason As String

        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanCharacter(Userindex, UserName, Reason)
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "UnbanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanChar(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim cantPenas As Byte
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            
            If Not FileExist(CharPath & UserName & ".chr", vbNormal) Then
                Call WriteConsoleMsg(Userindex, "Charfile inexistente (no use +).", FontTypeNames.FONTTYPE_INFO)
            Else
                If (val(GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban")) = 1) Then
                    Call UnBan(UserName)
                
                    'penas
                    cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": UNBAN. " & Date & " " & time)
                
                    Call LogGM(.Name, "/UNBAN a " & UserName)
                    Call WriteConsoleMsg(Userindex, UserName & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, UserName & " no está baneado. Imposible unbanear.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "NPCFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNPCFollow(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        If .flags.TargetNPC > 0 Then
            Call DoFollow(.flags.TargetNPC, .Name)
            Npclist(.flags.TargetNPC).flags.Inmovilizado = 0
            Npclist(.flags.TargetNPC).flags.Paralizado = 0
            Npclist(.flags.TargetNPC).Contadores.Paralisis = 0
        End If
    End With
End Sub

''
' Handles the "SummonChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSummonChar(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer
        Dim X  As Integer
        Dim Y  As Integer

        UserName = buffer.ReadASCIIString()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) Then
            tUser = NameIndex(UserName)

            If Not StrComp(UCase$(UserName), "THYRAH") = 0 Then
                If tUser <= 0 Then
                    Call WriteConsoleMsg(Userindex, "El jugador no está online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or _
                       (UserList(tUser).flags.Privilegios And (PlayerType.Consejero Or PlayerType.User)) <> 0 Then
                        
                        If UserList(tUser).flags.SlotEvent > 0 Or UserList(tUser).flags.SlotReto > 0 Then
                            Call WriteConsoleMsg(Userindex, "El personaje esta evento. Tene mayor cuidado para la proxima que me vas a buguear el evento " & .Name & ".", FontTypeNames.FONTTYPE_ADMIN)
                        Else
                            If Not UserList(tUser).Counters.Pena >= 1 Then
                                Call WriteConsoleMsg(tUser, .Name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
                                X = .Pos.X
                                Y = .Pos.Y + 1
                                Call FindLegalPos(tUser, .Pos.map, X, Y)
                                Call WarpUserChar(tUser, .Pos.map, X, Y, True, True)
                                Call LogGM(.Name, "/SUM " & UserName & " Map:" & .Pos.map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                            Else
                                Call WriteConsoleMsg(Userindex, "Está en la carcel", FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "SpawnListRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnListRequest(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        Call EnviarSpawnList(Userindex)
    End With
End Sub

''
' Handles the "SpawnCreature" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnCreature(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Npc As Integer
        Npc = .incomingData.ReadInteger()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If Npc > 0 And Npc <= UBound(Declaraciones.SpawnList()) Then _
               Call SpawnNpc(Declaraciones.SpawnList(Npc).NpcIndex, .Pos, True, False)

            Call LogGM(.Name, "Sumoneo " & Declaraciones.SpawnList(Npc).NpcName)
        End If
    End With
End Sub

''
' Handles the "ResetNPCInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResetNPCInventory(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        If .flags.TargetNPC = 0 Then Exit Sub

        Call ResetNpcInv(.flags.TargetNPC)
        Call LogGM(.Name, "/RESETINV " & Npclist(.flags.TargetNPC).Name)
    End With
End Sub

''
' Handles the "CleanWorld" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCleanWorld(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub

        CountDownLimpieza = 10
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Limpieza del mundo en 10 segundos. Recojan sus objetos para no perderlos.", FontTypeNames.FONTTYPE_SERVER)
        
        'Call LimpiarMundo
    End With
End Sub

''
' Handles the "ServerMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 28/05/2010
'28/05/2010: ZaMa - Ahora no dice el nombre del gm que lo dice.
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim message As String
        message = buffer.ReadASCIIString()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) Then
            If LenB(message) <> 0 Then
                Call LogGM(.Name, "Mensaje Broadcast:" & message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(message, FontTypeNames.FONTTYPE_GUILD))
                ''''''''''''''''SOLO PARA EL TESTEO'''''''
                ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
                'frmMain.txtChat.Text = frmMain.txtChat.Text & vbNewLine & UserList(UserIndex).name & " > " & message
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub
Private Sub HandleRolMensaje(ByVal Userindex As Integer)
'***************************************************
'Author: Tomás (Nibf ~) Para Gs zone y Servers Argentum
'Last Modification: 20/09/13
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim message As String
        message = buffer.ReadASCIIString()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(message) <> 0 Then
                Call LogGM(.Name, "Mensaje Broadcast:" & message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("" & .Name & "> " & message, FontTypeNames.FONTTYPE_GUILD))
                ''''''''''''''''SOLO PARA EL TESTEO'''''''
                ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
                frmMain.txtChat.Text = frmMain.txtChat.Text & vbNewLine & message
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "NickToIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNickToIP(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/06/2010
'Pablo (ToxicWaste): Agrego para que el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer
        Dim priv As PlayerType
        Dim IsAdmin As Boolean

        UserName = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            Call LogGM(.Name, "NICK2IP Solicito la IP de " & UserName)

            IsAdmin = (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0
            If IsAdmin Then
                priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            Else
                priv = PlayerType.User
            End If

            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And priv Then
                    Call WriteConsoleMsg(Userindex, "El ip de " & UserName & " es " & UserList(tUser).ip, FontTypeNames.FONTTYPE_INFO)
                    Dim ip As String
                    Dim lista As String
                    Dim LoopC As Long
                    ip = UserList(tUser).ip
                    For LoopC = 1 To LastUser
                        If UserList(LoopC).ip = ip Then
                            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                                If UserList(LoopC).flags.Privilegios And priv Then
                                    lista = lista & UserList(LoopC).Name & ", "
                                End If
                            End If
                        End If
                    Next LoopC
                    If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                    Call WriteConsoleMsg(Userindex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If Not (EsDios(UserName) Or EsAdmin(UserName)) Or IsAdmin Then
                    Call WriteConsoleMsg(Userindex, "No hay ningún personaje con ese nick.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub
''
' Handles the "IPToNick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleIPToNick(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim ip As String
        Dim LoopC As Long
        Dim lista As String
        Dim priv As PlayerType

        ip = .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, "IP2NICK Solicito los Nicks de IP " & ip)

        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
        Else
            priv = PlayerType.User
        End If

        For LoopC = 1 To LastUser
            If UserList(LoopC).ip = ip Then
                If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).flags.Privilegios And priv Then
                        lista = lista & UserList(LoopC).Name & ", "
                    End If
                End If
            End If
        Next LoopC

        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        Call WriteConsoleMsg(Userindex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "GuildOnlineMembers" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnlineMembers(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim GuildName As String
        Dim tGuild As Integer

        GuildName = buffer.ReadASCIIString()

        If (InStrB(GuildName, "+") <> 0) Then
            GuildName = Replace(GuildName, "+", " ")
        End If

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tGuild = GuildIndex(GuildName)

            If tGuild > 0 Then
                Call WriteConsoleMsg(Userindex, "Clan " & UCase(GuildName) & ": " & _
                                                modGuilds.m_ListaDeMiembrosOnline(Userindex, tGuild), FontTypeNames.FONTTYPE_GUILDMSG)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "TeleportCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportCreate(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 22/03/2010
'15/11/2009: ZaMa - Ahora se crea un teleport con un radio especificado.
'22/03/2010: ZaMa - Harcodeo los teleps y radios en el dat, para evitar mapas bugueados.
'***************************************************
    If UserList(Userindex).incomingData.length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Mapa As Integer
        Dim X  As Byte
        Dim Y  As Byte
        Dim Radio As Byte

        Mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        Radio = .incomingData.ReadByte()

        Radio = MinimoInt(Radio, 6)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        Call LogGM(.Name, "/CT " & Mapa & "," & X & "," & Y & "," & Radio)

        If Not MapaValido(Mapa) Or Not InMapBounds(Mapa, X, Y) Then _
           Exit Sub

        If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).ObjInfo.objindex > 0 Then _
           Exit Sub

        If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).TileExit.map > 0 Then _
           Exit Sub

        If MapData(Mapa, X, Y).ObjInfo.objindex > 0 Then
            Call WriteConsoleMsg(Userindex, "Hay un objeto en el piso en ese lugar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If MapData(Mapa, X, Y).TileExit.map > 0 Then
            Call WriteConsoleMsg(Userindex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Dim ET As Obj
        ET.Amount = 1
        ' Es el numero en el dat. El indice es el comienzo + el radio, todo harcodeado :(.
        ET.objindex = 378

        With MapData(.Pos.map, .Pos.X, .Pos.Y - 1)
            .TileExit.map = Mapa
            .TileExit.X = X
            .TileExit.Y = Y
        End With

        Call MakeObj(ET, .Pos.map, .Pos.X, .Pos.Y - 1)
    End With
End Sub

''
' Handles the "TeleportDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportDestroy(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(Userindex)
        Dim Mapa As Integer
        Dim X  As Byte
        Dim Y  As Byte

        'Remove packet ID
        Call .incomingData.ReadByte

        '/dt
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        Mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY

        If Not InMapBounds(Mapa, X, Y) Then Exit Sub

        With MapData(Mapa, X, Y)
            If .ObjInfo.objindex = 0 Then Exit Sub

            If ObjData(.ObjInfo.objindex).OBJType = eOBJType.otTeleport And .TileExit.map > 0 Then
                Call LogGM(UserList(Userindex).Name, "/DT: " & Mapa & "," & X & "," & Y)

                Call EraseObj(.ObjInfo.Amount, Mapa, X, Y)

                If MapData(.TileExit.map, .TileExit.X, .TileExit.Y).ObjInfo.objindex = 651 Then
                    Call EraseObj(1, .TileExit.map, .TileExit.X, .TileExit.Y)
                End If

                .TileExit.map = 0
                .TileExit.X = 0
                .TileExit.Y = 0
            End If
        End With
    End With
End Sub

''
' Handles the "RainToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRainToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        Call LogGM(.Name, "/LLUVIA")
        Lloviendo = Not Lloviendo

        'Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
    End With
End Sub

''
' Handles the "SetCharDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetCharDescription(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim tUser As Integer
        Dim desc As String

        desc = buffer.ReadASCIIString()

        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (.flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
            tUser = .flags.TargetUser
            If tUser > 0 Then
                UserList(tUser).DescRM = desc
            Else
                Call WriteConsoleMsg(Userindex, "Haz click sobre un personaje antes.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "ForceMIDIToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HanldeForceMIDIToMap(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim midiID As Byte
        Dim Mapa As Integer

        midiID = .incomingData.ReadByte
        Mapa = .incomingData.ReadInteger

        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(Mapa, 50, 50) Then
                Mapa = .Pos.map
            End If

            If midiID = 0 Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(MapInfo(.Pos.map).Music))
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(midiID))
            End If
        End If
    End With
End Sub

''
' Handles the "ForceWAVEToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEToMap(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim waveID As Byte
        Dim Mapa As Integer
        Dim X  As Byte
        Dim Y  As Byte

        waveID = .incomingData.ReadByte()
        Mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()

        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(Mapa, X, Y) Then
                Mapa = .Pos.map
                X = .Pos.X
                Y = .Pos.Y
            End If

            'Ponemos el pedido por el GM
            Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayWave(waveID, X, Y))
        End If
    End With
End Sub

''
' Handles the "RoyalArmyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim message As String
        message = buffer.ReadASCIIString()

        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster Or PlayerType.RoyalCouncil) Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Consejo de Banderbill] " & .Name & "> " & message, FontTypeNames.FONTTYPE_CONSEJOVesA))
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "ChaosLegionMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim message As String
        message = buffer.ReadASCIIString()

        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster Or PlayerType.ChaosCouncil) Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Concilio de las Sombras] " & .Name & "> " & message, FontTypeNames.FONTTYPE_EJECUCION))
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "CitizenMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCitizenMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim message As String
        message = buffer.ReadASCIIString()

        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCiudadanosYRMs, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & message, FontTypeNames.FONTTYPE_TALK))
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "CriminalMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCriminalMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim message As String
        message = buffer.ReadASCIIString()

        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCriminalesYRMs, 0, PrepareMessageConsoleMsg("CRIMINALES> " & message, FontTypeNames.FONTTYPE_TALK))
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "TalkAsNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalkAsNPC(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim message As String
        message = buffer.ReadASCIIString()

        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(message, Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Else
                Call WriteConsoleMsg(Userindex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "DestroyAllItemsInArea" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyAllItemsInArea(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Dim X  As Long
        Dim Y  As Long
        Dim bIsExit As Boolean

        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.map, X, Y).ObjInfo.objindex > 0 Then
                        bIsExit = MapData(.Pos.map, X, Y).TileExit.map > 0
                        If ItemNoEsDeMapa(MapData(.Pos.map, X, Y).ObjInfo.objindex, bIsExit) Then
                            Call EraseObj(MAX_INVENTORY_OBJS, .Pos.map, X, Y)
                        End If
                    End If
                End If
            Next X
        Next Y

        Call LogGM(UserList(Userindex).Name, "/MASSDEST")
    End With
End Sub

''
' Handles the "AcceptRoyalCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptRoyalCouncilMember(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte

        UserName = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                    If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil

                    Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False)
                End With
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "ChaosCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptChaosCouncilMember(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte

        UserName = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))

                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                    If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil

                    Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False)
                End With
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "ItemsInTheFloor" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleItemsInTheFloor(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Dim tobj As Integer
        Dim lista As String
        Dim X  As Long
        Dim Y  As Long

        For X = 5 To 95
            For Y = 5 To 95
                tobj = MapData(.Pos.map, X, Y).ObjInfo.objindex
                If tobj > 0 Then
                    If ObjData(tobj).OBJType <> eOBJType.otarboles Then
                        Call WriteConsoleMsg(Userindex, "(" & X & "," & Y & ") " & ObjData(tobj).Name, FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Next Y
        Next X
    End With
End Sub

''
' Handles the "MakeDumb" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumb(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer

        UserName = buffer.ReadASCIIString()

        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "MakeDumbNoMore" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumbNoMore(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer

        UserName = buffer.ReadASCIIString()

        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser)
                Call FlushBuffer(tUser)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "DumpIPTables" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDumpIPTables(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Call SecurityIp.DumpTables
    End With
End Sub

''
' Handles the "CouncilKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilKick(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer

        UserName = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline, echando de los consejos.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECE", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECECAOS", 0)
                Else
                    Call WriteConsoleMsg(Userindex, "No se encuentra el charfile " & CharPath & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Banderbill.", FontTypeNames.FONTTYPE_GUILD)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil

                        Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill.", FontTypeNames.FONTTYPE_INFO))
                    End If

                    If .flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_GUILD)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil

                        Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_INFO))
                    End If
                End With
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "SetTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetTrigger(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim tTrigger As Byte
        Dim tLog As String

        tTrigger = .incomingData.ReadByte()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        If tTrigger >= 0 Then
            MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = tTrigger
            tLog = "Trigger " & tTrigger & " en mapa " & .Pos.map & " " & .Pos.X & "," & .Pos.Y

            Call LogGM(.Name, tLog)
            Call WriteConsoleMsg(Userindex, tLog, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "AskTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAskTrigger(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 04/13/07
'
'***************************************************
    Dim tTrigger As Byte

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        tTrigger = MapData(.Pos.map, .Pos.X, .Pos.Y).trigger

        Call LogGM(.Name, "Miro el trigger en " & .Pos.map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)

        Call WriteConsoleMsg(Userindex, _
                             "Trigger " & tTrigger & " en mapa " & .Pos.map & " " & .Pos.X & ", " & .Pos.Y _
                             , FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "BannedIPList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Dim lista As String
        Dim LoopC As Long

        Call LogGM(.Name, "/BANIPLIST")

        For LoopC = 1 To BanIps.Count
            lista = lista & BanIps.Item(LoopC) & ", "
        Next LoopC

        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)

        Call WriteConsoleMsg(Userindex, lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call BanIpGuardar
        Call BanIpCargar
    End With
End Sub

''
' Handles the "GuildBan" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildBan(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim GuildName As String
        Dim cantMembers As Integer
        Dim LoopC As Long
        Dim member As String
        Dim Count As Byte
        Dim tIndex As Integer
        Dim tFile As String

        GuildName = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tFile = App.Path & "\guilds\" & GuildName & "-members.mem"

            If Not FileExist(tFile) Then
                Call WriteConsoleMsg(Userindex, "No existe el clan: " & GuildName, FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " baneó al clan " & UCase$(GuildName), FontTypeNames.FONTTYPE_FIGHT))

                'baneamos a los miembros
                Call LogGM(.Name, "BANCLAN a " & UCase$(GuildName))

                cantMembers = val(GetVar(tFile, "INIT", "NroMembers"))

                For LoopC = 1 To cantMembers
                    member = GetVar(tFile, "Members", "Member" & LoopC)
                    'member es la victima
                    Call Ban(member, "Administracion del servidor", "Clan Banned")

                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("   " & member & "<" & GuildName & "> ha sido expulsado del servidor.", FontTypeNames.FONTTYPE_FIGHT))

                    tIndex = NameIndex(member)
                    If tIndex > 0 Then
                        'esta online
                        UserList(tIndex).flags.Ban = 1
                        Call CloseSocket(tIndex)
                    End If

                    'ponemos el flag de ban a 1
                    Call WriteVar(CharPath & member & ".chr", "FLAGS", "Ban", "1")
                    'ponemos la pena
                    Count = val(GetVar(CharPath & member & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & member & ".chr", "PENAS", "Cant", Count + 1)
                    Call WriteVar(CharPath & member & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": BAN AL CLAN: " & GuildName & " " & Date & " " & time)
                Next LoopC
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub
''
' Handles the "CheckHD" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleCheckHD(ByVal Userindex As Integer)
'***************************************************
'Author: ArzenaTh
'Last Modification: 01/09/10
'Verifica el HD del usuario.
'***************************************************

    If UserList(Userindex).incomingData.length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler


    With UserList(Userindex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte

        Dim Usuario As Integer
        Dim nickUsuario As String
        nickUsuario = buffer.ReadASCIIString()
        Usuario = NameIndex(nickUsuario)

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then


            If Usuario = 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "El disco del usuario " & UserList(Usuario).Name & " es " & UserList(Usuario).HD, FONTTYPE_INFOBOLD)
            End If
        End If

        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub
''
' Handles the "BanHD" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleBanHD(ByVal Userindex As Integer)
'***************************************************
'Author: ArzenaTh
'Last Modification: 02/09/10
'Maneja el baneo del serial del HD de un usuario.
'***************************************************

    If UserList(Userindex).incomingData.length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte

        Dim Usuario As Integer
        Usuario = NameIndex(buffer.ReadASCIIString())
        Dim bannedHD As String
        If Usuario > 0 Then bannedHD = UserList(Usuario).HD
        Dim i  As Long    'El mandamás dijo Long, Long será.
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Dios) Then
            If LenB(bannedHD) > 0 Then
                If BuscarRegistroHD(bannedHD) > 0 Then
                    Call WriteConsoleMsg(Userindex, "El usuario ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call AgregarRegistroHD(bannedHD)
                    Call WriteConsoleMsg(Userindex, "Has baneado el root " & bannedHD & " del usuario " & UserList(Usuario).Name, FontTypeNames.FONTTYPE_INFO)
                    'Call CloseSocket(Usuario)
                    For i = 1 To LastUser
                        If UserList(i).ConnIDValida Then
                            If UserList(i).HD = bannedHD Then
                                Call BanCharacter(Userindex, UserList(i).Name, "t0 en el servidor")
                            End If
                        End If
                    Next i
                End If
            ElseIf Usuario <= 0 Then
                Call WriteConsoleMsg(Userindex, "El usuario no se encuentra online.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    Set buffer = Nothing

    If error <> 0 Then Err.Raise error
End Sub

''
' Handles the "UnBanHD" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUnbanHD(ByVal Userindex As Integer)
'***************************************************
'Author: ArzenaTh
'Last Modification: 02/09/10
'Maneja el unbaneo del serial del HD de un usuario.
'***************************************************

    If UserList(Userindex).incomingData.length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)

       If .flags.Privilegios And (PlayerType.User Or PlayerType.SemiDios Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
       
            Dim buffer As New clsByteQueue
            Call buffer.CopyBuffer(.incomingData)
            Call buffer.ReadByte

            Dim HD As String
            HD = buffer.ReadASCIIString()

            If (RemoverRegistroHD(HD)) Then
                Call WriteConsoleMsg(Userindex, "El root n°" & HD & " se ha quitado de la lista de baneados.", FONTTYPE_INFOBOLD)
            Else
                Call WriteConsoleMsg(Userindex, "El root n°" & HD & " no se encuentra en la lista de baneados.", FONTTYPE_INFO)
            End If
        
        Call .incomingData.CopyBuffer(buffer)
    End With
Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub


''
' Handles the "BanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/02/09
'Agregado un CopyBuffer porque se producia un bucle
'inifito al intentar banear una ip ya baneada. (NicoNZ)
'07/02/09 Pato - Ahora no es posible saber si un gm está o no online.
'***************************************************
    If UserList(Userindex).incomingData.length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim bannedIP As String
        Dim tUser As Integer
        Dim Reason As String
        Dim i  As Long

        ' Is it by ip??
        If buffer.ReadBoolean() Then
            bannedIP = buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte()
        Else
            tUser = NameIndex(buffer.ReadASCIIString())

            If tUser > 0 Then bannedIP = UserList(tUser).ip
        End If

        Reason = buffer.ReadASCIIString()


        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If LenB(bannedIP) > 0 Then
                Call LogGM(.Name, "/BanIP " & bannedIP & " por " & Reason)

                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(Userindex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call BanIpAgrega(bannedIP)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " baneó la IP " & bannedIP & " por " & Reason, FontTypeNames.FONTTYPE_FIGHT))

                    'Find every player with that ip and ban him!
                    For i = 1 To LastUser
                        If UserList(i).ConnIDValida Then
                            If UserList(i).ip = bannedIP Then
                                Call BanCharacter(Userindex, UserList(i).Name, "IP POR " & Reason)
                            End If
                        End If
                    Next i
                End If
            ElseIf tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "El personaje no está online.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

Private Sub HandleUnbanIP(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim bannedIP As String
        
        bannedIP = .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        If BanIpQuita(bannedIP) Then
            Call WriteConsoleMsg(Userindex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(Userindex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "CreateItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateItem(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 11/02/2011
'maTih.- : Ahora se puede elegir, la cantidad a crear.
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim tobj As Integer
        Dim Cuantos As Integer
        Dim tStr As String
        tobj = .incomingData.ReadInteger()
        Cuantos = .incomingData.ReadInteger()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios) Then Exit Sub

        Call LogGM(.Name, "/CI: " & tobj & " Cantidad : " & Cuantos)

        If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).ObjInfo.objindex > 0 Then _
           Exit Sub

        If Cuantos > 9999 Then Call WriteConsoleMsg(Userindex, "Demasiados, máximo para crear : 10.000", FontTypeNames.FONTTYPE_TALK): Exit Sub

        If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).TileExit.map > 0 Then _
           Exit Sub

        If tobj < 1 Or tobj > NumObjDatas Then _
           Exit Sub

        'Is the object not null?
        If LenB(ObjData(tobj).Name) = 0 Then Exit Sub

        Dim Objeto As Obj
        Call WriteConsoleMsg(Userindex, "¡¡ATENCIÓN: FUERON CREADOS ***" & Cuantos & "*** ÍTEMS, TIRE Y /DEST LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)

        Objeto.Amount = Cuantos
        Objeto.objindex = tobj
        Call MakeObj(Objeto, .Pos.map, .Pos.X, .Pos.Y - 1)

        'Agrega a la lista.
        Dim tmpPos As WorldPos

        tmpPos = .Pos
        tmpPos.Y = .Pos.X - 1


    End With
End Sub

''
' Handles the "DestroyItems" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyItems(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        If MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.objindex = 0 Then Exit Sub

        Call LogGM(.Name, "/DEST")

        If ObjData(MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.objindex).OBJType = eOBJType.otTeleport And _
           MapData(.Pos.map, .Pos.X, .Pos.Y).TileExit.map > 0 Then

            Call WriteConsoleMsg(Userindex, "No puede destruir teleports así. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Call EraseObj(10000, .Pos.map, .Pos.X, .Pos.Y)
    End With
End Sub

''
' Handles the "ChaosLegionKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionKick(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer

        UserName = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And _
            (.flags.Privilegios And PlayerType.ChaosCouncil) Or _
            (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            tUser = NameIndex(UserName)

            Call LogGM(.Name, "ECHO DEL CAOS A: " & UserName)

            If tUser > 0 Then
                Call ExpulsarFaccionCaos(tUser, True)
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(Userindex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas del caos.", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(tUser)
            Else
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoCaos", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas", 200)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & .Name)
                    Call WriteConsoleMsg(Userindex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "RoyalArmyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyKick(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer
        Dim Privs As PlayerType
        
        UserName = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And _
             .flags.Privilegios And PlayerType.RoyalCouncil Or _
             (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            tUser = NameIndex(UserName)

            Call LogGM(.Name, "ECHÓ DE LA REAL A: " & UserName)

            If tUser > 0 Then
                Call ExpulsarFaccionReal(tUser, True)
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(Userindex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas reales.", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(tUser)
            Else
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoReal", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas", 200)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & .Name)
                    Call WriteConsoleMsg(Userindex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "ForceMIDIAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceMIDIAll(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim midiID As Byte
        midiID = .incomingData.ReadByte()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " broadcast música: " & midiID, FontTypeNames.FONTTYPE_SERVER))

        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMidi(midiID))
    End With
End Sub

''
' Handles the "ForceWAVEAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEAll(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim waveID As Byte
        waveID = .incomingData.ReadByte()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))
    End With
End Sub

''
' Handles the "RemovePunishment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRemovePunishment(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 1/05/07
'Pablo (ToxicWaste): 1/05/07, You can now edit the punishment.
'***************************************************
    If UserList(Userindex).incomingData.length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim punishment As Byte
        Dim NewText As String

        UserName = buffer.ReadASCIIString()
        punishment = buffer.ReadByte
        NewText = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Then
                Call WriteConsoleMsg(Userindex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", FontTypeNames.FONTTYPE_INFO)
            Else
                If (InStrB(UserName, "\") <> 0) Then
                    UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                    UserName = Replace(UserName, "/", "")
                End If

                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    Call LogGM(.Name, " borro la pena: " & punishment & "-" & _
                                      GetVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment) _
                                      & " de " & UserName & " y la cambió por: " & NewText)

                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment, LCase$(.Name) & ": <" & NewText & "> " & Date & " " & time)

                    Call WriteConsoleMsg(Userindex, "Pena modificada.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "TileBlockedToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTileBlockedToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        Call LogGM(.Name, "/BLOQ")

        If MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked = 0 Then
            MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked = 1
        Else
            MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked = 0
        End If

        Call Bloquear(True, .Pos.map, .Pos.X, .Pos.Y, MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked)
    End With
End Sub

''
' Handles the "KillNPCNoRespawn" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPCNoRespawn(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        If .flags.TargetNPC = 0 Then Exit Sub

        Call QuitarNPC(.flags.TargetNPC)
        Call LogGM(.Name, "/MATA " & Npclist(.flags.TargetNPC).Name)
    End With
End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Dim X  As Long
        Dim Y  As Long

        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(.Pos.map, X, Y).NpcIndex)
                End If
            Next X
        Next Y
        Call LogGM(.Name, "/MASSKILL")
    End With
End Sub

''
' Handles the "LastIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLastIP(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim lista As String
        Dim LoopC As Byte
        Dim priv As Integer
        Dim validCheck As Boolean

        priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        UserName = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
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
                validCheck = (UserList(NameIndex(UserName)).flags.Privilegios And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            Else
                validCheck = (UserDarPrivilegioLevel(UserName) And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            End If

            If validCheck Then
                Call LogGM(.Name, "/LASTIP " & UserName)

                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    lista = "Las ultimas IPs con las que " & UserName & " se conectó son:"
                    For LoopC = 1 To 5
                        lista = lista & vbCrLf & LoopC & " - " & GetVar(CharPath & UserName & ".chr", "INIT", "LastIP" & LoopC)
                    Next LoopC
                    Call WriteConsoleMsg(Userindex, lista, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, "Charfile """ & UserName & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(Userindex, UserName & " es de mayor jerarquía que vos.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "ChatColor" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleChatColor(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the user`s chat color
'***************************************************
    If UserList(Userindex).incomingData.length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim color As Long

        color = RGB(.incomingData.ReadByte(), .incomingData.ReadByte(), .incomingData.ReadByte())

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            .flags.ChatColor = color
        End If
    End With
End Sub

''
' Handles the "Ignored" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIgnored(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Ignore the user
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            .flags.AdminPerseguible = Not .flags.AdminPerseguible
        End If
    End With
End Sub
Public Sub HandleUserOro(ByVal Userindex As Integer)

    With UserList(Userindex)
        Call .incomingData.ReadByte


        'Lo hace vip
        If .flags.Oro = 0 Then
            .flags.Oro = 1
        End If

        'Le da el Random de vida entre 5 y 13 cambiar a su gusto
        If .Stats.MaxHp + RandomNumber(1, 3) Then
        End If
    End With
End Sub
Public Sub HandleUserPlata(ByVal Userindex As Integer)

    With UserList(Userindex)
        Call .incomingData.ReadByte


        'Lo hace vip
        If .flags.Plata = 0 Then
            .flags.Plata = 1
        End If

        'Le da el Random de vida entre 5 y 13 cambiar a su gusto
        If .Stats.MaxHp + RandomNumber(1, 2) Then
        End If
    End With
End Sub
Public Sub HandleUserBronce(ByVal Userindex As Integer)

    With UserList(Userindex)
        Call .incomingData.ReadByte


        'Lo hace vip
        If .flags.Bronce = 0 Then
            .flags.Bronce = 1
        End If

        'Le da el Random de vida entre 5 y 13 cambiar a su gusto
        If .Stats.MaxHp + RandomNumber(1, 1) Then
        End If
    End With
End Sub

''
' Handles the "CheckSlot" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleCheckSlot(ByVal Userindex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 09/09/2008 (NicoNZ)
'Check one Users Slot in Particular from Inventory
'***************************************************
    If UserList(Userindex).incomingData.length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim Slot As Byte
        Dim tIndex As Integer

        UserName = buffer.ReadASCIIString()    'Que UserName?
        Slot = buffer.ReadByte()    'Que Slot?

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Dios) Then
            tIndex = NameIndex(UserName)  'Que user index?

            Call LogGM(.Name, .Name & " Checkeó el slot " & Slot & " de " & UserName)

            If tIndex > 0 Then
                If Slot > 0 And Slot <= UserList(tIndex).CurrentInventorySlots Then
                    If UserList(tIndex).Invent.Object(Slot).objindex > 0 Then
                        Call WriteConsoleMsg(Userindex, " Objeto " & Slot & ") " & ObjData(UserList(tIndex).Invent.Object(Slot).objindex).Name & " Cantidad:" & UserList(tIndex).Invent.Object(Slot).Amount, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(Userindex, "No hay ningún objeto en slot seleccionado.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "Slot Inválido.", FontTypeNames.FONTTYPE_TALK)
                End If
            Else
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handles the "ResetAutoUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleResetAutoUpdate(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reset the AutoUpdate
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If StrComp(UCase$(.Name), "THYRAH") <> 0 Then Exit Sub

        Call WriteConsoleMsg(Userindex, "TID: " & CStr(ReiniciarAutoUpdate()), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Restart" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleRestart(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Restart the game
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If StrComp(UCase$(.Name), "THYRAH") <> 0 Then Exit Sub
        
        'time and Time BUG!
        Call LogGM(.Name, .Name & " reinició el mundo.")

        Call ReiniciarServidor(True)
    End With
End Sub

''
' Handles the "ReloadObjects" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadObjects(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the objects
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha recargado los objetos.")

        Call LoadOBJData
    End With
End Sub

''
' Handles the "ReloadSpells" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadSpells(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the spells
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha recargado los hechizos.")

        Call CargarHechizos
    End With
End Sub

''
' Handle the "ReloadServerIni" message.
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadServerIni(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s INI
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha recargado los INITs.")

        Call LoadSini
    End With
End Sub

''
' Handle the "ReloadNPCs" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadNPCs(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s NPC
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha recargado los NPCs.")

        Call CargaNpcsDat

        Call WriteConsoleMsg(Userindex, "Npcs.dat recargado.", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "KickAllChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleKickAllChars(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Kick all the chars that are online
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha echado a todos los personajes.")

        Call EcharPjsNoPrivilegiados
    End With
End Sub

''
' Handle the "Night" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNight(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If StrComp(UCase$(.Name), "THYRAH") <> 0 Then Exit Sub

        DeNoche = Not DeNoche

        Dim i  As Long

        For i = 1 To NumUsers
            If UserList(i).flags.UserLogged And UserList(i).ConnID > -1 Then
                Call EnviarNoche(i)
            End If
        Next i
    End With
End Sub

''
' Handle the "ShowServerForm" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowServerForm(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Show the server form
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha solicitado mostrar el formulario del servidor.")
        Call frmMain.mnuMostrar_Click
    End With
End Sub

''
' Handle the "CleanSOS" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCleanSOS(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Clean the SOS
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha borrado los SOS.")

        Call Ayuda.Reset
    End With
End Sub

''
' Handle the "SaveChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Save the characters
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha guardado todos los chars.")

        Call mdParty.ActualizaExperiencias
        Call GuardarUsuarios
    End With
End Sub

''
' Handle the "ChangeMapInfoBackup" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the backup`s info of the map
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        Dim doTheBackUp As Boolean

        doTheBackUp = .incomingData.ReadBoolean()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub

        Call LogGM(.Name, .Name & " ha cambiado la información sobre el BackUp.")

        'Change the boolean to byte in a fast way
        If doTheBackUp Then
            MapInfo(.Pos.map).BackUp = 1
        Else
            MapInfo(.Pos.map).BackUp = 0
        End If

        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.map & ".dat", "Mapa" & .Pos.map, "backup", MapInfo(.Pos.map).BackUp)

        Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.map & " Backup: " & MapInfo(.Pos.map).BackUp, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ChangeMapInfoPK" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoPK(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the pk`s info of the  map
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        Dim isMapPk As Boolean

        isMapPk = .incomingData.ReadBoolean()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub

        Call LogGM(.Name, .Name & " ha cambiado la información sobre si es PK el mapa.")

        MapInfo(.Pos.map).Pk = isMapPk

        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.map & ".dat", "Mapa" & .Pos.map, "Pk", IIf(isMapPk, "1", "0"))

        Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.map & " PK: " & MapInfo(.Pos.map).Pk, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ChangeMapInfoRestricted" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal Userindex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS", "FACCION".
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    Dim tStr   As String

    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove Packet ID
        Call buffer.ReadByte

        tStr = buffer.ReadASCIIString()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Or tStr = "QUINCE" Or tStr = "VEINTE" Or tStr = "VEINTICINCO" Or tStr = "CUARENTA" Or tStr = "SEIS" Or tStr = "SIETE" Or tStr = "OCHO" Or tStr = "NUEVE" Or tStr = "CINCO" Or tStr = "MENOSCINCO" Or tStr = "MENOSCUATRO" Or tStr = "NOESUM" Or tStr = "VIPP" Or tStr = "VIP" Then
                Call LogGM(.Name, .Name & " ha cambiado la información sobre si es restringido el mapa.")
                MapInfo(UserList(Userindex).Pos.map).Restringir = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(Userindex).Pos.map & ".dat", "Mapa" & UserList(Userindex).Pos.map, "Restringir", tStr)
                Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.map & " Restringido: " & MapInfo(.Pos.map).Restringir, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION', 'QUINCE',  'VENTE', 'VEINTICINCO', 'CUARENTA', 'CINCO', 'SEIS', 'SIETE', 'OCHO', 'NUEVE', 'MENOSCINCO', 'MENOSCUATRO', 'NOESUM', 'VIPP', 'VIPP'.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handle the "ChangeMapInfoNoMagic" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoMagic(ByVal Userindex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'MagiaSinEfecto -> Options: "1" , "0".
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    Dim nomagic As Boolean

    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        nomagic = .incomingData.ReadBoolean

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido usar la magia el mapa.")
            MapInfo(UserList(Userindex).Pos.map).MagiaSinEfecto = nomagic
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(Userindex).Pos.map & ".dat", "Mapa" & UserList(Userindex).Pos.map, "MagiaSinEfecto", nomagic)
            Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.map & " MagiaSinEfecto: " & MapInfo(.Pos.map).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoNoInvi" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvi(ByVal Userindex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'InviSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    Dim noinvi As Boolean

    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        noinvi = .incomingData.ReadBoolean()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido usar la invisibilidad en el mapa.")
            MapInfo(UserList(Userindex).Pos.map).InviSinEfecto = noinvi
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(Userindex).Pos.map & ".dat", "Mapa" & UserList(Userindex).Pos.map, "InviSinEfecto", noinvi)
            Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.map & " InviSinEfecto: " & MapInfo(.Pos.map).InviSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoNoResu" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoResu(ByVal Userindex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'ResuSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    Dim noresu As Boolean

    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        noresu = .incomingData.ReadBoolean()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido usar el resucitar en el mapa.")
            MapInfo(UserList(Userindex).Pos.map).ResuSinEfecto = noresu
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(Userindex).Pos.map & ".dat", "Mapa" & UserList(Userindex).Pos.map, "ResuSinEfecto", noresu)
            Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.map & " ResuSinEfecto: " & MapInfo(.Pos.map).ResuSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoLand" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoLand(ByVal Userindex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    Dim tStr   As String

    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove Packet ID
        Call buffer.ReadByte

        tStr = buffer.ReadASCIIString()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la información del terreno del mapa.")
                MapInfo(UserList(Userindex).Pos.map).Terreno = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(Userindex).Pos.map & ".dat", "Mapa" & UserList(Userindex).Pos.map, "Terreno", tStr)
                Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.map & " Terreno: " & MapInfo(.Pos.map).Terreno, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(Userindex, "Igualmente, el único útil es 'NIEVE' ya que al ingresarlo, la gente muere de frío en el mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handle the "ChangeMapInfoZone" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoZone(ByVal Userindex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    Dim tStr   As String

    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove Packet ID
        Call buffer.ReadByte

        tStr = buffer.ReadASCIIString()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la información de la zona del mapa.")
                MapInfo(UserList(Userindex).Pos.map).Zona = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(Userindex).Pos.map & ".dat", "Mapa" & UserList(Userindex).Pos.map, "Zona", tStr)
                Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.map & " Zona: " & MapInfo(.Pos.map).Zona, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(Userindex, "Igualmente, el único útil es 'DUNGEON' ya que al ingresarlo, NO se sentirá el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub
''
' Handle the "ChangeMapInfoStealNp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoStealNpc(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 25/07/2010
'RoboNpcsPermitido -> Options: "1", "0"
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    Dim RoboNpc As Byte

    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        RoboNpc = val(IIf(.incomingData.ReadBoolean(), 1, 0))

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido robar npcs en el mapa.")

            MapInfo(UserList(Userindex).Pos.map).RoboNpcsPermitido = RoboNpc

            Call WriteVar(App.Path & MapPath & "mapa" & UserList(Userindex).Pos.map & ".dat", "Mapa" & UserList(Userindex).Pos.map, "RoboNpcsPermitido", RoboNpc)
            Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.map & " RoboNpcsPermitido: " & MapInfo(.Pos.map).RoboNpcsPermitido, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoNoOcultar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoOcultar(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 18/09/2010
'OcultarSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    Dim NoOcultar As Byte
    Dim Mapa   As Integer

    With UserList(Userindex)

        'Remove Packet ID
        Call .incomingData.ReadByte

        NoOcultar = val(IIf(.incomingData.ReadBoolean(), 1, 0))

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then

            Mapa = .Pos.map

            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido ocultarse en el mapa " & Mapa & ".")

            MapInfo(Mapa).OcultarSinEfecto = NoOcultar

            Call WriteVar(App.Path & MapPath & "mapa" & Mapa & ".dat", "Mapa" & Mapa, "OcultarSinEfecto", NoOcultar)
            Call WriteConsoleMsg(Userindex, "Mapa " & Mapa & " OcultarSinEfecto: " & NoOcultar, FontTypeNames.FONTTYPE_INFO)
        End If

    End With

End Sub

''
' Handle the "ChangeMapInfoNoInvocar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvocar(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 18/09/2010
'InvocarSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    Dim NoInvocar As Byte
    Dim Mapa   As Integer

    With UserList(Userindex)

        'Remove Packet ID
        Call .incomingData.ReadByte

        NoInvocar = val(IIf(.incomingData.ReadBoolean(), 1, 0))

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then

            Mapa = .Pos.map

            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido invocar en el mapa " & Mapa & ".")

            MapInfo(Mapa).InvocarSinEfecto = NoInvocar

            Call WriteVar(App.Path & MapPath & "mapa" & Mapa & ".dat", "Mapa" & Mapa, "InvocarSinEfecto", NoInvocar)
            Call WriteConsoleMsg(Userindex, "Mapa " & Mapa & " InvocarSinEfecto: " & NoInvocar, FontTypeNames.FONTTYPE_INFO)
        End If

    End With

End Sub

''
' Handle the "SaveMap" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Saves the map
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha guardado el mapa " & CStr(.Pos.map))

        Call GrabarMapa(.Pos.map, App.Path & "\WorldBackUp\Mapa" & .Pos.map)

        Call WriteConsoleMsg(Userindex, "Mapa Guardado.", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ShowGuildMessages" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowGuildMessages(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Allows admins to read guild messages
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim guild As String

        guild = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call modGuilds.GMEscuchaClan(Userindex, guild)
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handle the "DoBackUp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Show guilds messages
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha hecho un backup.")

        Call ES.DoBackUp    'Sino lo confunde con la id del paquete
    End With
End Sub

''
' Handle the "ToggleCentinelActivated" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleToggleCentinelActivated(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/26/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Activate or desactivate the Centinel
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        centinelaActivado = Not centinelaActivado

        With Centinela
            .RevisandoUserIndex = 0
            .clave = 0
            .TiempoRestante = 0
        End With

        If CentinelaNPCIndex Then
            Call QuitarNPC(CentinelaNPCIndex)
            CentinelaNPCIndex = 0
        End If

        If centinelaActivado Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido activado.", FontTypeNames.FONTTYPE_SERVER))
        Else
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido desactivado.", FontTypeNames.FONTTYPE_SERVER))
        End If
    End With
End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterName(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user name
'***************************************************
    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        'Reads the userName and newUser Packets
        Dim UserName As String
        Dim newName As String
        Dim changeNameUI As Integer
        Dim GuildIndex As Integer

        UserName = buffer.ReadASCIIString()
        newName = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin)) Then
            If LenB(UserName) = 0 Or LenB(newName) = 0 Then
                Call WriteConsoleMsg(Userindex, "Usar: /ANAME origen@destino", FontTypeNames.FONTTYPE_INFO)
            Else
                changeNameUI = NameIndex(UserName)

                If changeNameUI > 0 Then
                    Call WriteConsoleMsg(Userindex, "El Pj está online, debe salir para hacer el cambio.", FontTypeNames.FONTTYPE_WARNING)
                Else
                    If Not FileExist(CharPath & UserName & ".chr") Then
                        Call WriteConsoleMsg(Userindex, "El pj " & UserName & " es inexistente.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        GuildIndex = val(GetVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX"))

                        If GuildIndex > 0 Then
                            Call WriteConsoleMsg(Userindex, "El pj " & UserName & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If Not FileExist(CharPath & newName & ".chr") Then
                                Call FileCopy(CharPath & UserName & ".chr", CharPath & UCase$(newName) & ".chr")

                                Call WriteConsoleMsg(Userindex, "Transferencia exitosa.", FontTypeNames.FONTTYPE_INFO)

                                Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")

                                Dim cantPenas As Byte

                                cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))

                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", CStr(cantPenas + 1))

                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & CStr(cantPenas + 1), LCase$(.Name) & ": BAN POR Cambio de nick a " & UCase$(newName) & " " & Date & " " & time)

                                Call LogGM(.Name, "Ha cambiado de nombre al usuario " & UserName & ". Ahora se llama " & newName)
                            Else
                                Call WriteConsoleMsg(Userindex, "El nick solicitado ya existe.", FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterMail(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user password
'***************************************************
    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim newMail As String

        UserName = buffer.ReadASCIIString()
        newMail = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin)) Then
            If LenB(UserName) = 0 Or LenB(newMail) = 0 Then
                Call WriteConsoleMsg(Userindex, "usar /AEMAIL <pj>-<nuevomail>", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not FileExist(CharPath & UserName & ".chr") Then
                    Call WriteConsoleMsg(Userindex, "No existe el charfile " & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteVar(CharPath & UserName & ".chr", "CONTACTO", "Email", newMail)
                    Call WriteConsoleMsg(Userindex, "Email de " & UserName & " cambiado a: " & newMail, FontTypeNames.FONTTYPE_INFO)
                End If

                Call LogGM(.Name, "Le ha cambiado el mail a " & UserName)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handle the "AlterPassword" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterPassword(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user password
'***************************************************
    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim copyFrom As String
        Dim Password As String
        
        UserName = Replace(buffer.ReadASCIIString(), "+", " ")
        copyFrom = Replace(buffer.ReadASCIIString(), "+", " ")
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin)) Then
            Call LogGM(.Name, "Ha alterado la contraseña de " & UserName)
            
            If LenB(UserName) = 0 Or LenB(copyFrom) = 0 Then
                Call WriteConsoleMsg(Userindex, "usar /APASS <pjsinpass>@<pjconpass>", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not FileExist(CharPath & UserName & ".chr") Or Not FileExist(CharPath & copyFrom & ".chr") Then
                    Call WriteConsoleMsg(Userindex, "Alguno de los PJs no existe " & UserName & "@" & copyFrom, FontTypeNames.FONTTYPE_INFO)
                Else
                    Password = GetVar(CharPath & copyFrom & ".chr", "INIT", "Password")
                    Call WriteVar(CharPath & UserName & ".chr", "INIT", "Password", Password)
                    
                    Call WriteConsoleMsg(Userindex, "Password de " & UserName & " ha cambiado por la de " & copyFrom, FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "HandleCreateNPC" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        Dim NpcIndex As Integer

        NpcIndex = .incomingData.ReadInteger()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)

        If NpcIndex <> 0 Then
            Call LogGM(.Name, "Sumoneó a " & Npclist(NpcIndex).Name & " en mapa " & .Pos.map)
        End If
    End With
End Sub


''
' Handle the "CreateNPCWithRespawn" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPCWithRespawn(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        Dim NpcIndex As Integer

        NpcIndex = .incomingData.ReadInteger()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True)

        If NpcIndex <> 0 Then
            Call LogGM(.Name, "Sumoneó con respawn " & Npclist(NpcIndex).Name & " en mapa " & .Pos.map)
        End If
    End With
End Sub

''
' Handle the "ImperialArmour" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleImperialArmour(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        Dim index As Byte
        Dim objindex As Integer

        index = .incomingData.ReadByte()
        objindex = .incomingData.ReadInteger()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Select Case index
        Case 1
            ArmaduraImperial1 = objindex

        Case 2
            ArmaduraImperial2 = objindex

        Case 3
            ArmaduraImperial3 = objindex

        Case 4
            TunicaMagoImperial = objindex
        End Select
    End With
End Sub

''
' Handle the "ChaosArmour" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChaosArmour(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        Dim index As Byte
        Dim objindex As Integer

        index = .incomingData.ReadByte()
        objindex = .incomingData.ReadInteger()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Select Case index
        Case 1
            ArmaduraCaos1 = objindex

        Case 2
            ArmaduraCaos2 = objindex

        Case 3
            ArmaduraCaos3 = objindex

        Case 4
            TunicaMagoCaos = objindex
        End Select
    End With
End Sub

''
' Handle the "NavigateToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNavigateToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/12/07
'
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        If .flags.Navegando = 1 Then
            .flags.Navegando = 0
        Else
            .flags.Navegando = 1
        End If

        'Tell the client that we are navigating.
        Call WriteNavigateToggle(Userindex)
    End With
End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        If ServerSoloGMs > 0 Then
            Call WriteConsoleMsg(Userindex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 0
        Else
            Call WriteConsoleMsg(Userindex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 1
        End If
    End With
End Sub

''
' Handle the "TurnOffServer" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnOffServer(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'Turns off the server
'***************************************************
    Dim handle As Integer

    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, "/APAGAR")
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡¡¡" & .Name & " VA A APAGAR EL SERVIDOR!!!", FontTypeNames.FONTTYPE_FIGHT))

        'Log
        handle = FreeFile
        Open App.Path & "\logs\Main.log" For Append Shared As #handle

        Print #handle, Date & " " & time & " server apagado por " & .Name & ". "

        Close #handle

        Unload frmMain
    End With
End Sub

''
' Handle the "TurnCriminal" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnCriminal(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer

        UserName = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/CONDEN " & UserName)

            tUser = NameIndex(UserName)
            If tUser > 0 Then _
               Call VolverCriminal(tUser)
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handle the "ResetFactions" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleResetFactionCaos(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 06/09/09
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer
        Dim Char As String

        UserName = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And _
            (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.ChaosCouncil)) Then
            Call LogGM(.Name, "/PERDONARCAOS " & UserName)

            tUser = NameIndex(UserName)

            If tUser > 0 Then
                Call ResetFaccionCaos(tUser)
            Else
                Char = CharPath & UserName & ".chr"

                If FileExist(Char, vbNormal) Then
                    Call WriteVar(Char, "FACCIONES", "EjercitoReal", 0)
                    Call WriteVar(Char, "FACCIONES", "CrimMatados", 0)
                    Call WriteVar(Char, "FACCIONES", "EjercitoCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "FechaIngreso", "No ingresó a ninguna Facción")
                    Call WriteVar(Char, "FACCIONES", "rArCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "rArReal", 0)
                    Call WriteVar(Char, "FACCIONES", "rExCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "rExReal", 0)
                    Call WriteVar(Char, "FACCIONES", "recCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "recReal", 0)
                    Call WriteVar(Char, "FACCIONES", "Reenlistadas", 0)
                    Call WriteVar(Char, "FACCIONES", "NivelIngreso", 0)
                    Call WriteVar(Char, "FACCIONES", "MatadosIngreso", 0)
                    Call WriteVar(Char, "FACCIONES", "NextRecompensa", 0)
                Else
                    Call WriteConsoleMsg(Userindex, "El personaje " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub
''
' Handle the "ResetFactions" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleResetFactionReal(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 06/09/09
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim tUser As Integer
        Dim Char As String

        UserName = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoyalCouncil)) Then
            Call LogGM(.Name, "/PERDONARREAL " & UserName)

            tUser = NameIndex(UserName)

            If tUser > 0 Then
                Call ResetFaccionReal(tUser)
            Else
                Char = CharPath & UserName & ".chr"

                If FileExist(Char, vbNormal) Then
                    Call WriteVar(Char, "FACCIONES", "EjercitoReal", 0)
                    Call WriteVar(Char, "FACCIONES", "CiudMatados", 0)
                    Call WriteVar(Char, "FACCIONES", "EjercitoCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "FechaIngreso", "No ingresó a ninguna Facción")
                    Call WriteVar(Char, "FACCIONES", "rArCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "rArReal", 0)
                    Call WriteVar(Char, "FACCIONES", "rExCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "rExReal", 0)
                    Call WriteVar(Char, "FACCIONES", "recCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "recReal", 0)
                    Call WriteVar(Char, "FACCIONES", "Reenlistadas", 0)
                    Call WriteVar(Char, "FACCIONES", "NivelIngreso", 0)
                    Call WriteVar(Char, "FACCIONES", "MatadosIngreso", 0)
                    Call WriteVar(Char, "FACCIONES", "NextRecompensa", 0)
                Else
                    Call WriteConsoleMsg(Userindex, "El personaje " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handle the "RemoveCharFromGuild" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRemoveCharFromGuild(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim GuildIndex As Integer

        UserName = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/RAJARCLAN " & UserName)

            GuildIndex = modGuilds.m_EcharMiembroDeClan(Userindex, UserName)

            If GuildIndex = 0 Then
                Call WriteConsoleMsg(Userindex, "No pertenece a ningún clan o es fundador.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Expulsado.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido expulsado del clan por los administradores del servidor.", FontTypeNames.FONTTYPE_GUILD))
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handle the "RequestCharMail" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRequestCharMail(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Request user mail
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String
        Dim mail As String

        UserName = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If FileExist(CharPath & UserName & ".chr") Then
                mail = GetVar(CharPath & UserName & ".chr", "CONTACTO", "email")

                Call WriteConsoleMsg(Userindex, "Last email de " & UserName & ":" & mail, FontTypeNames.FONTTYPE_INFO)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handle the "SystemMessage" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/29/06
'Send a message to all the users
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim message As String
        message = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Mensaje de sistema:" & message)

            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(message))
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

''
' Handle the "Ping" message
'
' @param userIndex The index of the user sending the message

Public Sub HandlePing(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Show guilds messages
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        Call WritePong(Userindex)
    End With
End Sub

''
' Handle the "SetIniVar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetIniVar(ByVal Userindex As Integer)
'***************************************************
'Author: Brian Chaia (BrianPr)
'Last Modification: 01/23/10 (Marco)
'Modify server.ini
'***************************************************
    If UserList(Userindex).incomingData.length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo Errhandler

    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim sLlave As String
        Dim sClave As String
        Dim sValor As String

        'Obtengo los parámetros
        sLlave = buffer.ReadASCIIString()
        sClave = buffer.ReadASCIIString()
        sValor = buffer.ReadASCIIString()

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            Dim sTmp As String

            'No podemos modificar [INIT]Dioses ni [Dioses]*
            If (UCase$(sLlave) = "INIT" And UCase$(sClave) = "ADMINES") Or UCase$(sLlave) = "ADMINES" Then
                Call WriteConsoleMsg(Userindex, "¡No puedes modificar esa información desde aquí!", FontTypeNames.FONTTYPE_INFO)
            If (UCase$(sLlave) = "INIT" And UCase$(sClave) = "DIOSES") Or UCase$(sLlave) = "DIOSES" Then
                Call WriteConsoleMsg(Userindex, "¡No puedes modificar esa información desde aquí!", FontTypeNames.FONTTYPE_INFO)
            Else
                'Obtengo el valor según llave y clave
                sTmp = GetVar(IniPath & "Server.ini", sLlave, sClave)

                'Si obtengo un valor escribo en el server.ini
                If LenB(sTmp) Then
                    Call WriteVar(IniPath & "Server.ini", sLlave, sClave, sValor)
                    Call LogGM(.Name, "Modificó en server.ini (" & sLlave & " " & sClave & ") el valor " & sTmp & " por " & sValor)
                    Call WriteConsoleMsg(Userindex, "Modificó " & sLlave & " " & sClave & " a " & sValor & ". Valor anterior " & sTmp, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, "No existe la llave y/o clave", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long

    error = Err.Number

On Error GoTo 0
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal Userindex As Integer)
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Logged)
    Exit Sub
Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub


''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.RemoveDialogs)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal Userindex As Integer, ByVal CharIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageRemoveCharDialog(CharIndex))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
Public Function PrepareMessageCreateDamage(ByVal X As Byte, ByVal Y As Byte, ByVal DamageValue As Long, ByVal DamageType As Byte)

' @ Envia el paquete para crear daño (Y)

    With auxiliarBuffer
        .WriteByte ServerPacketID.CreateDamage
        .WriteByte X
        .WriteByte Y
        .WriteLong DamageValue
        .WriteByte DamageType

        PrepareMessageCreateDamage = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Writes the "NavigateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NavigateToggle" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.NavigateToggle)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Disconnect" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Disconnect)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
Public Sub WriteMontateToggle(ByVal Userindex As Integer)
    On Error Resume Next
    Dim obData As ObjData
    obData = ObjData(UserList(Userindex).Invent.MonturaObjIndex)

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.MontateToggle)
    Call UserList(Userindex).outgoingData.WriteByte(obData.Velocidad)
End Sub

''
' Writes the "UserOfferConfirm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserOfferConfirm(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "UserOfferConfirm" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.UserOfferConfirm)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub


''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceEnd" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.CommerceEnd)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankEnd" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BankEnd)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceInit" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.CommerceInit)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankInit" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BankInit)
    Call UserList(Userindex).outgoingData.WriteLong(UserList(Userindex).Stats.Banco)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.UserCommerceInit)
    Call UserList(Userindex).outgoingData.WriteASCIIString(UserList(Userindex).ComUsu.DestNick)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.UserCommerceEnd)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowBlacksmithForm(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowBlacksmithForm)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowCarpenterForm(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowCarpenterForm)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateSta)
        Call .WriteInteger(UserList(Userindex).Stats.MinSta)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateMana)
        Call .WriteInteger(UserList(Userindex).Stats.MinMAN)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHP)
        Call .WriteInteger(UserList(Userindex).Stats.MinHp)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateGold" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateGold)
        Call .WriteLong(UserList(Userindex).Stats.Gld)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UpdateBankGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateBankGold(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "UpdateBankGold" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateBankGold)
        Call .WriteLong(UserList(Userindex).Stats.Banco)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub


''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateExp" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateExp)
        Call .WriteLong(UserList(Userindex).Stats.Exp)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenghtAndDexterity(ByVal Userindex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenghtAndDexterity)
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad))
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateDexterity(ByVal Userindex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateDexterity)
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad))
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenght(ByVal Userindex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenght)
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza))
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal Userindex As Integer, ByVal map As Integer, ByVal version As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMap" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeMap)
        Call .WriteInteger(map)
        Call .WriteASCIIString(MapInfo(map).Name)
        Call .WriteInteger(version)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PosUpdate" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.PosUpdate)
        Call .WriteByte(UserList(Userindex).Pos.X)
        Call .WriteByte(UserList(Userindex).Pos.Y)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatOverHead(ByVal Userindex As Integer, ByVal chat As String, ByVal CharIndex As Integer, ByVal color As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChatOverHead" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHead(chat, CharIndex, color))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal Userindex As Integer, ByVal chat As String, ByVal FontIndex As FontTypeNames)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(chat, FontIndex))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteCommerceChat(ByVal Userindex As Integer, ByVal chat As String, ByVal FontIndex As FontTypeNames)
'***************************************************
'Author: ZaMa
'Last Modification: 05/17/06
'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareCommerceConsoleMsg(chat, FontIndex))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "GuildChat" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildChat(ByVal Userindex As Integer, ByVal chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildChat" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageGuildChat(chat))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMessageBox(ByVal Userindex As Integer, ByVal message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(message)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserIndexInServer)
        Call .WriteInteger(Userindex)
        Randomize
            UserList(Userindex).KeyUseItem = RandomNumber(1, 12)
        Call .WriteByte(UserList(Userindex).KeyUseItem)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserCharIndexInServer)
        Call .WriteInteger(UserList(Userindex).Char.CharIndex)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
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

Public Sub WriteCharacterCreate(ByVal Userindex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal Heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Name As String, ByVal NickColor As Byte, _
                                ByVal Privileges As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterCreate" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterCreate(body, Head, Heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, _
                                                                                              helmet, Name, NickColor, Privileges))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal Userindex As Integer, ByVal CharIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterRemove" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(CharIndex))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal Userindex As Integer, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterMove" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterMove(CharIndex, X, Y))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteForceCharMove(ByVal Userindex, ByVal Direccion As eHeading)
'***************************************************
'Author: ZaMa
'Last Modification: 26/03/2009
'Writes the "ForceCharMove" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageForceCharMove(Direccion))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CharacterChange" message to the given user's outgoing data buffer.
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

Public Sub WriteCharacterChange(ByVal Userindex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal Heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterChange" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChange(body, Head, Heading, CharIndex, weapon, shield, FX, FXLoops, helmet))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectCreate(ByVal Userindex As Integer, ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectCreate" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectCreate(GrhIndex, X, Y))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal Userindex As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectDelete" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectDelete(X, Y))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal Userindex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlockPosition" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "PlayMidi" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMidi(ByVal Userindex As Integer, ByVal midi As Byte, Optional ByVal loops As Integer = -1)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PlayMidi" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayMidi(midi, loops))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "PlayWave" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayWave(ByVal Userindex As Integer, ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave, X, Y))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "GuildList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildList(ByVal Userindex As Integer, ByRef guildList() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildList" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Dim tmp    As String
    Dim i      As Long

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.guildList)

        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            tmp = tmp & guildList(i) & SEPARATOR
        Next i

        If Len(tmp) Then _
           tmp = Left$(tmp, Len(tmp) - 1)

        Call .WriteASCIIString(tmp)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAreaChanged(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AreaChanged" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.AreaChanged)
        Call .WriteByte(UserList(Userindex).Pos.X)
        Call .WriteByte(UserList(Userindex).Pos.Y)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PauseToggle" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle())
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub


''
' Writes the "CreateFX" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal Userindex As Integer, ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateFX" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(CharIndex, FX, FXLoops))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserStats)
        Call .WriteInteger(UserList(Userindex).Stats.MaxHp)
        Call .WriteInteger(UserList(Userindex).Stats.MinHp)
        Call .WriteInteger(UserList(Userindex).Stats.MaxMAN)
        Call .WriteInteger(UserList(Userindex).Stats.MinMAN)
        Call .WriteInteger(UserList(Userindex).Stats.MaxSta)
        Call .WriteInteger(UserList(Userindex).Stats.MinSta)
        Call .WriteLong(UserList(Userindex).Stats.Gld)
        Call .WriteByte(UserList(Userindex).Stats.ELV)
        Call .WriteLong(UserList(Userindex).Stats.ELU)
        Call .WriteLong(UserList(Userindex).Stats.Exp)
        .WriteByte (UserList(Userindex).flags.Oculto)
        Call .WriteBoolean(UserList(Userindex).flags.ModoCombate)
        .WriteInteger UserList(Userindex).Char.CharIndex
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkRequestTarget(ByVal Userindex As Integer, ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WorkRequestTarget" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.WorkRequestTarget)
        Call .WriteByte(Skill)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteChangeInventorySlot(ByVal Userindex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 3/12/09
'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
'3/12/09: Budi - Ahora se envia MaxDef y MinDef en lugar de Def
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeInventorySlot)
        Call .WriteByte(Slot)

        Dim objindex As Integer
        Dim obData As ObjData

        objindex = UserList(Userindex).Invent.Object(Slot).objindex
        
        Call .WriteInteger(objindex)
        Call .WriteInteger(UserList(Userindex).Invent.Object(Slot).Amount)
        Call .WriteBoolean(UserList(Userindex).Invent.Object(Slot).Equipped)
        
        If objindex > 0 Then
            obData = ObjData(objindex)
            Call .WriteInteger(obData.GrhIndex)
            Call .WriteByte(obData.OBJType)
            Call .WriteInteger(obData.MaxHIT)
            Call .WriteInteger(obData.MinHIT)
            Call .WriteInteger(obData.MaxDef)
            Call .WriteInteger(obData.MinDef)
            Call .WriteSingle(SalePrice(objindex))
        End If
    End With
    
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub



''
' Writes the "ChangeBankSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal Userindex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/03/09
'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
'12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeBankSlot)
        Call .WriteByte(Slot)

        Dim objindex As Integer
        Dim obData As ObjData

        objindex = UserList(Userindex).BancoInvent.Object(Slot).objindex

        Call .WriteInteger(objindex)
        Call .WriteInteger(UserList(Userindex).BancoInvent.Object(Slot).Amount)

        If objindex > 0 Then
            obData = ObjData(objindex)
            Call .WriteInteger(obData.GrhIndex)
            Call .WriteByte(obData.OBJType)
            Call .WriteInteger(obData.MaxHIT)
            Call .WriteInteger(obData.MinHIT)
            Call .WriteInteger(obData.MaxDef)
            Call .WriteInteger(obData.MinDef)
            Call .WriteLong(obData.valor)
        End If
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal Userindex As Integer, ByVal Slot As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
' /ver paquete mañana by lautaro
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeSpellSlot)
        Call .WriteByte(Slot)
        Call .WriteInteger(UserList(Userindex).Stats.UserHechizos(Slot))

        If UserList(Userindex).Stats.UserHechizos(Slot) > 0 Then
            Call .WriteASCIIString(Hechizos(UserList(Userindex).Stats.UserHechizos(Slot)).Nombre)
        Else
            Call .WriteASCIIString("(Vacio)")
        End If
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Atributes" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.Atributes)
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Carisma))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Constitucion))
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithWeapons(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
'Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Dim i      As Long
    Dim Obj    As ObjData
    Dim validIndexes() As Integer
    Dim Count  As Integer

    ReDim validIndexes(1 To UBound(ArmasHerrero()))

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithWeapons)

        For i = 1 To UBound(ArmasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmasHerrero(i)).SkHerreria <= Round(UserList(Userindex).Stats.UserSkills(eSkill.herreria) / ModHerreriA(UserList(Userindex).clase), 0) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i

        ' Write the number of objects in the list
        Call .WriteInteger(Count)

        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ArmasHerrero(validIndexes(i)))
            Call .WriteASCIIString(Obj.Name)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(ArmasHerrero(validIndexes(i)))
        Next i
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub


''
' Writes the "BlacksmithArmors" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Sub WriteBlacksmithArmors(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
'Writes the "BlacksmithArmors" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Dim i      As Long
    Dim Obj    As ObjData
    Dim validIndexes() As Integer
    Dim Count  As Integer

    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithArmors)

        For i = 1 To UBound(ArmadurasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmadurasHerrero(i)).SkHerreria <= Round(UserList(Userindex).Stats.UserSkills(eSkill.herreria) / ModHerreriA(UserList(Userindex).clase), 0) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i

        ' Write the number of objects in the list
        Call .WriteInteger(Count)

        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ArmadurasHerrero(validIndexes(i)))
            Call .WriteASCIIString(Obj.Name)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(ArmadurasHerrero(validIndexes(i)))
        Next i
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCarpenterObjects(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Dim i      As Long
    Dim Obj    As ObjData
    Dim validIndexes() As Integer
    Dim Count  As Integer

    ReDim validIndexes(1 To UBound(ObjCarpintero()))

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.CarpenterObjects)

        For i = 1 To UBound(ObjCarpintero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(Userindex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(Userindex).clase) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i

        ' Write the number of objects in the list
        Call .WriteInteger(Count)

        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ObjCarpintero(validIndexes(i)))
            Call .WriteASCIIString(Obj.Name)
            Call .WriteInteger(Obj.Madera)
            Call .WriteInteger(ObjCarpintero(validIndexes(i)))
        Next i
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "RestOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RestOK" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.RestOK)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteErrorMsg(ByVal Userindex As Integer, ByVal message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ErrorMsg" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageErrorMsg(message))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "Blind" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Blind" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Blind)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "Dumb" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Dumb" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Dumb)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSignal(ByVal Userindex As Integer, ByVal objindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowSignal" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSignal)
        Call .WriteASCIIString(ObjData(objindex).texto)
        Call .WriteInteger(ObjData(objindex).GrhSecundario)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal Userindex As Integer, ByVal Slot As Byte, ByRef Obj As Obj, ByVal price As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/03/09
'Last Modified by: Budi
'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
'12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
'***************************************************
    On Error GoTo Errhandler
    Dim ObjInfo As ObjData

    If Obj.objindex >= LBound(ObjData()) And Obj.objindex <= UBound(ObjData()) Then
        ObjInfo = ObjData(Obj.objindex)
    End If

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeNPCInventorySlot)
        Call .WriteByte(Slot)
        Call .WriteInteger(Obj.objindex)
        
        If Obj.objindex > 0 Then
            Call .WriteInteger(Obj.Amount)
            Call .WriteSingle(price)
            Call .WriteByte(ObjInfo.copaS)
            Call .WriteByte(ObjInfo.Eldhir)
            Call .WriteInteger(ObjInfo.GrhIndex)
            
            Call .WriteByte(ObjInfo.OBJType)
            Call .WriteInteger(ObjInfo.MaxHIT)
            Call .WriteInteger(ObjInfo.MinHIT)
            Call .WriteInteger(ObjInfo.MaxDef)
            Call .WriteInteger(ObjInfo.MinDef)
        End If
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHungerAndThirst)
        Call .WriteByte(UserList(Userindex).Stats.MaxAGU)
        Call .WriteByte(UserList(Userindex).Stats.MinAGU)
        Call .WriteByte(UserList(Userindex).Stats.MaxHam)
        Call .WriteByte(UserList(Userindex).Stats.MinHam)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "Fame" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteFame(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Fame" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.Fame)

        Call .WriteLong(UserList(Userindex).Reputacion.AsesinoRep)
        Call .WriteLong(UserList(Userindex).Reputacion.BandidoRep)
        Call .WriteLong(UserList(Userindex).Reputacion.BurguesRep)
        Call .WriteLong(UserList(Userindex).Reputacion.LadronesRep)
        Call .WriteLong(UserList(Userindex).Reputacion.NobleRep)
        Call .WriteLong(UserList(Userindex).Reputacion.PlebeRep)
        Call .WriteLong(UserList(Userindex).Reputacion.Promedio)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "MiniStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MiniStats" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.MiniStats)

        Call .WriteLong(UserList(Userindex).Faccion.CiudadanosMatados)
        Call .WriteLong(UserList(Userindex).Faccion.CriminalesMatados)

        'TODO : Este valor es calculable, no debería NI EXISTIR, ya sea en el servidor ni en el cliente!!!
        Call .WriteLong(UserList(Userindex).Stats.UsuariosMatados)

        Call .WriteInteger(UserList(Userindex).Stats.NPCsMuertos)

        Call .WriteByte(UserList(Userindex).clase)
        Call .WriteLong(UserList(Userindex).Counters.Pena)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data buffer.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLevelUp(ByVal Userindex As Integer, ByVal skillPoints As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LevelUp" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.LevelUp)
        Call .WriteInteger(skillPoints)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "AddForumMsg" message to the given user's outgoing data buffer.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAddForumMsg(ByVal Userindex As Integer, ByVal ForumType As eForumType, _
                            ByRef Title As String, ByRef Author As String, ByRef message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 02/01/2010
'Writes the "AddForumMsg" message to the given user's outgoing data buffer
'02/01/2010: ZaMa - Now sends Author and forum type
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.AddForumMsg)
        Call .WriteByte(ForumType)
        Call .WriteASCIIString(Title)
        Call .WriteASCIIString(Author)
        Call .WriteASCIIString(message)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowForumForm(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowForumForm" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler

    Dim Visibilidad As Byte
    Dim CanMakeSticky As Byte

    With UserList(Userindex)
        Call .outgoingData.WriteByte(ServerPacketID.ShowForumForm)

        Visibilidad = eForumVisibility.ieGENERAL_MEMBER

        If esCaos(Userindex) Or EsGM(Userindex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieCAOS_MEMBER
        End If

        If esArmada(Userindex) Or EsGM(Userindex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieREAL_MEMBER
        End If

        Call .outgoingData.WriteByte(Visibilidad)

        ' Pueden mandar sticky los gms o los del consejo de armada/caos
        If EsGM(Userindex) Then
            CanMakeSticky = 2
        ElseIf (.flags.Privilegios And PlayerType.ChaosCouncil) <> 0 Then
            CanMakeSticky = 1
        ElseIf (.flags.Privilegios And PlayerType.RoyalCouncil) <> 0 Then
            CanMakeSticky = 1
        End If

        Call .outgoingData.WriteByte(CanMakeSticky)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal Userindex As Integer, ByVal CharIndex As Integer, ByVal invisible As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetInvisible" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetInvisible(CharIndex, invisible))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "DiceRoll" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDiceRoll(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DiceRoll" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.DiceRoll)

        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Carisma))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Constitucion))
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MeditateToggle" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.MeditateToggle)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlindNoMore" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BlindNoMore)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DumbNoMore" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.DumbNoMore)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 11/19/09
'Writes the "SendSkills" message to the given user's outgoing data buffer
'11/19/09: Pato - Now send the percentage of progress of the skills.
'***************************************************
    On Error GoTo Errhandler
    Dim i      As Long

    With UserList(Userindex)
        Call .outgoingData.WriteByte(ServerPacketID.SendSkills)
        Call .outgoingData.WriteByte(.clase)

        For i = 1 To NUMSKILLS
            Call .outgoingData.WriteByte(UserList(Userindex).Stats.UserSkills(i))
            If .Stats.UserSkills(i) < MAXSKILLPOINTS Then
                Call .outgoingData.WriteByte(Int(.Stats.ExpSkills(i) * 100 / .Stats.EluSkills(i)))
            Else
                Call .outgoingData.WriteByte(0)
            End If
        Next i
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainerCreatureList(ByVal Userindex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TrainerCreatureList" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Dim i      As Long
    Dim Str    As String

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.TrainerCreatureList)

        For i = 1 To Npclist(NpcIndex).NroCriaturas
            Str = Str & Npclist(NpcIndex).Criaturas(i).NpcName & SEPARATOR
        Next i

        If LenB(Str) > 0 Then _
           Str = Left$(Str, Len(Str) - 1)

        Call .WriteASCIIString(Str)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "GuildNews" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildNews The guild's news.
' @param    enemies The list of the guild's enemies.
' @param    allies The list of the guild's allies.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNews(ByVal Userindex As Integer, ByVal guildNews As String, ByRef enemies() As String, ByRef allies() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildNews" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Dim i      As Long
    Dim tmp    As String

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.guildNews)

        Call .WriteASCIIString(guildNews)

        'Prepare enemies' list
        For i = LBound(enemies()) To UBound(enemies())
            tmp = tmp & enemies(i) & SEPARATOR
        Next i

        If Len(tmp) Then _
           tmp = Left$(tmp, Len(tmp) - 1)

        Call .WriteASCIIString(tmp)

        tmp = vbNullString
        'Prepare allies' list
        For i = LBound(allies()) To UBound(allies())
            tmp = tmp & allies(i) & SEPARATOR
        Next i

        If Len(tmp) Then _
           tmp = Left$(tmp, Len(tmp) - 1)

        Call .WriteASCIIString(tmp)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOfferDetails(ByVal Userindex As Integer, ByVal details As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OfferDetails" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Dim i      As Long

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.OfferDetails)

        Call .WriteASCIIString(details)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlianceProposalsList(ByVal Userindex As Integer, ByRef guilds() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlianceProposalsList" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Dim i      As Long
    Dim tmp    As String

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.AlianceProposalsList)

        ' Prepare guild's list
        For i = LBound(guilds()) To UBound(guilds())
            tmp = tmp & guilds(i) & SEPARATOR
        Next i

        If Len(tmp) Then _
           tmp = Left$(tmp, Len(tmp) - 1)

        Call .WriteASCIIString(tmp)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePeaceProposalsList(ByVal Userindex As Integer, ByRef guilds() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PeaceProposalsList" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Dim i      As Long
    Dim tmp    As String

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.PeaceProposalsList)

        ' Prepare guilds' list
        For i = LBound(guilds()) To UBound(guilds())
            tmp = tmp & guilds(i) & SEPARATOR
        Next i

        If Len(tmp) Then _
           tmp = Left$(tmp, Len(tmp) - 1)

        Call .WriteASCIIString(tmp)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CharacterInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    charName The requested char's name.
' @param    race The requested char's race.
' @param    class The requested char's class.
' @param    gender The requested char's gender.
' @param    level The requested char's level.
' @param    gold The requested char's gold.
' @param    reputation The requested char's reputation.
' @param    previousPetitions The requested char's previous petitions to enter guilds.
' @param    currentGuild The requested char's current guild.
' @param    previousGuilds The requested char's previous guilds.
' @param    RoyalArmy True if tha char belongs to the Royal Army.
' @param    CaosLegion True if tha char belongs to the Caos Legion.
' @param    citicensKilled The number of citicens killed by the requested char.
' @param    criminalsKilled The number of criminals killed by the requested char.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterInfo(ByVal Userindex As Integer, ByVal charName As String, ByVal race As eRaza, ByVal Class As eClass, _
                              ByVal gender As eGenero, ByVal level As Byte, ByVal Gold As Long, ByVal bank As Long, ByVal reputation As Long, _
                              ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal RoyalArmy As Boolean, _
                              ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterInfo" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.CharacterInfo)

        Call .WriteASCIIString(charName)
        Call .WriteByte(race)
        Call .WriteByte(Class)
        Call .WriteByte(gender)

        Call .WriteByte(level)
        Call .WriteLong(Gold)
        Call .WriteLong(bank)
        Call .WriteLong(reputation)

        Call .WriteASCIIString(previousPetitions)
        Call .WriteASCIIString(currentGuild)
        Call .WriteASCIIString(previousGuilds)

        Call .WriteBoolean(RoyalArmy)
        Call .WriteBoolean(CaosLegion)

        Call .WriteLong(citicensKilled)
        Call .WriteLong(criminalsKilled)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @param    guildNews The guild's news.
' @param    joinRequests The list of chars which requested to join the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeaderInfo(ByVal Userindex As Integer, ByRef guildList() As String, ByRef MemberList() As String, _
                                ByVal guildNews As String, ByRef joinRequests() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Dim i      As Long
    Dim tmp    As String

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.GuildLeaderInfo)

        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            tmp = tmp & guildList(i) & SEPARATOR
        Next i

        If Len(tmp) Then _
           tmp = Left$(tmp, Len(tmp) - 1)

        Call .WriteASCIIString(tmp)

        ' Prepare guild member's list
        tmp = vbNullString
        For i = LBound(MemberList()) To UBound(MemberList())
            tmp = tmp & MemberList(i) & SEPARATOR
        Next i

        If Len(tmp) Then _
           tmp = Left$(tmp, Len(tmp) - 1)

        Call .WriteASCIIString(tmp)

        ' Store guild news
        Call .WriteASCIIString(guildNews)

        ' Prepare the join request's list
        tmp = vbNullString
        For i = LBound(joinRequests()) To UBound(joinRequests())
            tmp = tmp & joinRequests(i) & SEPARATOR
        Next i

        If Len(tmp) Then _
           tmp = Left$(tmp, Len(tmp) - 1)

        Call .WriteASCIIString(tmp)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal Userindex As Integer, ByRef guildList() As String, ByRef MemberList() As String)
'***************************************************
'Author: ZaMa
'Last Modification: 21/02/2010
'Writes the "GuildMemberInfo" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Dim i      As Long
    Dim tmp    As String

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.GuildMemberInfo)

        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            tmp = tmp & guildList(i) & SEPARATOR
        Next i

        If Len(tmp) Then _
           tmp = Left$(tmp, Len(tmp) - 1)

        Call .WriteASCIIString(tmp)

        ' Prepare guild member's list
        tmp = vbNullString
        For i = LBound(MemberList()) To UBound(MemberList())
            tmp = tmp & MemberList(i) & SEPARATOR
        Next i

        If Len(tmp) Then _
           tmp = Left$(tmp, Len(tmp) - 1)

        Call .WriteASCIIString(tmp)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "GuildDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildName The requested guild's name.
' @param    founder The requested guild's founder.
' @param    foundationDate The requested guild's foundation date.
' @param    leader The requested guild's current leader.
' @param    URL The requested guild's website.
' @param    memberCount The requested guild's member count.
' @param    electionsOpen True if the clan is electing it's new leader.
' @param    alignment The requested guild's alignment.
' @param    enemiesCount The requested guild's enemy count.
' @param    alliesCount The requested guild's ally count.
' @param    antifactionPoints The requested guild's number of antifaction acts commited.
' @param    codex The requested guild's codex.
' @param    guildDesc The requested guild's description.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDetails(ByVal Userindex As Integer, ByVal GuildName As String, ByVal founder As String, ByVal foundationDate As String, _
                             ByVal leader As String, ByVal URL As String, ByVal memberCount As Integer, ByVal electionsOpen As Boolean, _
                             ByVal alignment As String, ByVal enemiesCount As Integer, ByVal AlliesCount As Integer, _
                             ByVal antifactionPoints As String, ByRef codex() As String, ByVal guildDesc As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildDetails" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Dim i      As Long
    Dim temp   As String

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.GuildDetails)

        Call .WriteASCIIString(GuildName)
        Call .WriteASCIIString(founder)
        Call .WriteASCIIString(foundationDate)
        Call .WriteASCIIString(leader)
        Call .WriteASCIIString(URL)

        Call .WriteInteger(memberCount)
        Call .WriteBoolean(electionsOpen)

        Call .WriteASCIIString(alignment)

        Call .WriteInteger(enemiesCount)
        Call .WriteInteger(AlliesCount)

        Call .WriteASCIIString(antifactionPoints)

        For i = LBound(codex()) To UBound(codex())
            temp = temp & codex(i) & SEPARATOR
        Next i

        If Len(temp) > 1 Then _
           temp = Left$(temp, Len(temp) - 1)

        Call .WriteASCIIString(temp)

        Call .WriteASCIIString(guildDesc)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub


''
' Writes the "ShowGuildAlign" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildAlign(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "ShowGuildAlign" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowGuildAlign)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildFundationForm(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowGuildFundationForm)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/12/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Writes the "ParalizeOK" message to the given user's outgoing data buffer
'And updates user position
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ParalizeOK)
    Call WritePosUpdate(Userindex)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal Userindex As Integer, ByVal details As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowUserRequest)

        Call .WriteASCIIString(details)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "TradeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTradeOK(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TradeOK" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.TradeOK)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "BankOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankOK(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankOK" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BankOK)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeUserTradeSlot(ByVal Userindex As Integer, ByVal OfferSlot As Byte, ByVal objindex As Integer, ByVal Amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/03/09
'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
'25/11/2009: ZaMa - Now sends the specific offer slot to be modified.
'12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeUserTradeSlot)

        Call .WriteByte(OfferSlot)
        Call .WriteInteger(objindex)
        Call .WriteLong(Amount)

        If objindex > 0 Then
            Call .WriteInteger(ObjData(objindex).GrhIndex)
            Call .WriteByte(ObjData(objindex).OBJType)
            Call .WriteInteger(ObjData(objindex).MaxHIT)
            Call .WriteInteger(ObjData(objindex).MinHIT)
            Call .WriteInteger(ObjData(objindex).MaxDef)
            Call .WriteInteger(ObjData(objindex).MinDef)
            Call .WriteLong(SalePrice(objindex))
        Else    ' Borra el item
            Call .WriteInteger(0)
            Call .WriteByte(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteLong(0)
        End If
    End With
    Exit Sub


Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "SendNight" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendNight(ByVal Userindex As Integer, ByVal night As Boolean)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Writes the "SendNight" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.SendNight)
        Call .WriteBoolean(night)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal Userindex As Integer, ByRef npcNames() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnList" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Dim i      As Long
    Dim tmp    As String

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.SpawnList)

        For i = LBound(npcNames()) To UBound(npcNames())
            tmp = tmp & npcNames(i) & SEPARATOR
        Next i

        If Len(tmp) Then _
           tmp = Left$(tmp, Len(tmp) - 1)

        Call .WriteASCIIString(tmp)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSOSForm(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowSOSForm" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Dim i      As Long
    Dim tmp    As String

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSOSForm)

        For i = 1 To Ayuda.Longitud
            tmp = tmp & Ayuda.VerElemento(i) & SEPARATOR
        Next i

        If LenB(tmp) <> 0 Then _
           tmp = Left$(tmp, Len(tmp) - 1)

        Call .WriteASCIIString(tmp)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub


''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowPartyForm(ByVal Userindex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "ShowPartyForm" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Dim i      As Long
    Dim tmp    As String
    Dim PI     As Integer
    Dim members(PARTY_MAXMEMBERS) As Integer

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowPartyForm)

        PI = UserList(Userindex).PartyIndex
        Call .WriteByte(CByte(Parties(PI).EsPartyLeader(Userindex)))

        If PI > 0 Then
            Call Parties(PI).ObtenerMiembrosOnline(members())
            For i = 1 To PARTY_MAXMEMBERS
                If members(i) > 0 Then
                    tmp = tmp & UserList(members(i)).Name & " (" & Fix(Parties(PI).MiExperiencia(members(i))) & ")" & SEPARATOR
                End If
            Next i
        End If

        If LenB(tmp) <> 0 Then _
           tmp = Left$(tmp, Len(tmp) - 1)

        Call .WriteASCIIString(tmp)
        Call .WriteLong(Parties(PI).ObtenerExperienciaTotal)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowGMPanelForm)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal Userindex As Integer, ByRef userNamesList() As String, ByVal cant As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06 NIGO:
'Writes the "UserNameList" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Dim i      As Long
    Dim tmp    As String

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserNameList)

        ' Prepare user's names list
        For i = 1 To cant
            tmp = tmp & userNamesList(i) & SEPARATOR
        Next i

        If Len(tmp) Then _
           tmp = Left$(tmp, Len(tmp) - 1)

        Call .WriteASCIIString(tmp)
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "Pong" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Pong" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Pong)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Sends all data existing in the buffer
'***************************************************
    Dim sndData As String

    With UserList(Userindex).outgoingData
        If .length = 0 Then _
           Exit Sub

        sndData = .ReadASCIIStringFixed(.length)

        Call EnviarDatosASlot(Userindex, sndData)
    End With
End Sub

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, ByVal invisible As Boolean) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "SetInvisible" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetInvisible)

        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(invisible)

        PrepareMessageSetInvisible = .ReadASCIIStringFixed(.length)
    End With
End Function
Private Function PrepareMessageChangeHeading(ByVal CharIndex As Integer, ByVal Heading As Byte) As String

'***************************************************
'Author: Nacho
'Last Modification: 07/19/2016
'Prepares the "Change Heading" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.MiniPekka)

        Call .WriteInteger(CharIndex)
        Call .WriteByte(Heading)

        PrepareMessageChangeHeading = .ReadASCIIStringFixed(.length)

    End With

End Function

Public Function PrepareMessageCharacterChangeNick(ByVal CharIndex As Integer, ByVal newNick As String) As String
'***************************************************
'Author: Budi
'Last Modification: 07/23/09
'Prepares the "Change Nick" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChangeNick)

        Call .WriteInteger(CharIndex)
        Call .WriteASCIIString(newNick)

        PrepareMessageCharacterChangeNick = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageChatOverHead(ByVal chat As String, ByVal CharIndex As Integer, ByVal color As Long) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ChatOverHead" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChatOverHead)
        Call .WriteASCIIString(chat)
        Call .WriteInteger(CharIndex)

        ' Write rgb channels and save one byte from long :D
        Call .WriteByte(color And &HFF)
        Call .WriteByte((color And &HFF00&) \ &H100&)
        Call .WriteByte((color And &HFF0000) \ &H10000)

        PrepareMessageChatOverHead = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal chat As String, ByVal FontIndex As FontTypeNames) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ConsoleMsg" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ConsoleMsg)
        Call .WriteASCIIString(chat)
        Call .WriteByte(FontIndex)

        PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareCommerceConsoleMsg(ByRef chat As String, ByVal FontIndex As FontTypeNames) As String
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'Prepares the "CommerceConsoleMsg" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CommerceChat)
        Call .WriteASCIIString(chat)
        Call .WriteByte(FontIndex)

        PrepareCommerceConsoleMsg = .ReadASCIIStringFixed(.length)
    End With
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

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CreateFX" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateFX)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)

        PrepareMessageCreateFX = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayWave(ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayWave)
        Call .WriteByte(wave)
        Call .WriteByte(X)
        Call .WriteByte(Y)

        PrepareMessagePlayWave = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageGuildChat(ByVal chat As String) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "GuildChat" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.GuildChat)
        Call .WriteASCIIString(chat)

        PrepareMessageGuildChat = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageShowMessageBox(ByVal chat As String) As String
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "ShowMessageBox" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(chat)

        PrepareMessageShowMessageBox = .ReadASCIIStringFixed(.length)
    End With
End Function


''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMidi(ByVal midi As Byte, Optional ByVal loops As Integer = -1) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "GuildChat" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayMIDI)
        Call .WriteByte(midi)
        Call .WriteInteger(loops)

        PrepareMessagePlayMidi = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "PauseToggle" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PauseToggle)
        PrepareMessagePauseToggle = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ObjectDelete" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectDelete)
        Call .WriteByte(X)
        Call .WriteByte(Y)

        PrepareMessageObjectDelete = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean) As String
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "BlockPosition" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)

        PrepareMessageBlockPosition = .ReadASCIIStringFixed(.length)
    End With

End Function

''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectCreate(ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'prepares the "ObjectCreate" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectCreate)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(GrhIndex)

        PrepareMessageObjectCreate = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterRemove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterRemove)
        Call .WriteInteger(CharIndex)

        PrepareMessageCharacterRemove = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RemoveCharDialog)
        Call .WriteInteger(CharIndex)

        PrepareMessageRemoveCharDialog = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
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

Public Function PrepareMessageCharacterCreate(ByVal body As Integer, ByVal Head As Integer, ByVal Heading As eHeading, _
                                              ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, _
                                              ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Name As String, ByVal NickColor As Byte, _
                                              ByVal Privileges As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterCreate" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterCreate)

        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteByte(Heading)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        Call .WriteASCIIString(Name)
        Call .WriteByte(NickColor)
        Call .WriteByte(Privileges)

        PrepareMessageCharacterCreate = .ReadASCIIStringFixed(.length)
    End With
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

Public Function PrepareMessageCharacterChange(ByVal body As Integer, ByVal Head As Integer, ByVal Heading As eHeading, _
                                              ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                              ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterChange" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChange)

        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteByte(Heading)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)

        PrepareMessageCharacterChange = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterMove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterMove)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(X)
        Call .WriteByte(Y)

        PrepareMessageCharacterMove = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As String
'***************************************************
'Author: ZaMa
'Last Modification: 26/03/2009
'Prepares the "ForceCharMove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ForceCharMove)
        Call .WriteByte(Direccion)

        PrepareMessageForceCharMove = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal Userindex As Integer, ByVal NickColor As Byte, _
                                                 ByRef Tag As String, Optional ByVal Infectado As Byte, Optional ByVal Angel As Byte, Optional ByVal Demonio As Byte) As String
'***************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modification: 04/07/07
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'Prepares the "UpdateTagAndStatus" message and returns it
'15/01/2010: ZaMa - Now sends the nick color instead of the status.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.UpdateTagAndStatus)

        Call .WriteInteger(UserList(Userindex).Char.CharIndex)
        Call .WriteByte(NickColor)
        Call .WriteASCIIString(Tag)
        Call .WriteByte(Infectado)
        Call .WriteByte(Angel)
        Call .WriteByte(Demonio)
        PrepareMessageUpdateTagAndStatus = .ReadASCIIStringFixed(.length)
    End With
End Function


''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal message As String) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ErrorMsg" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ErrorMsg)
        Call .WriteASCIIString(message)

        PrepareMessageErrorMsg = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Writes the "StopWorking" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.

Public Sub WriteStopWorking(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 21/02/2010
'
'***************************************************
    On Error GoTo Errhandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.StopWorking)

    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CancelOfferItem" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Slot      The slot to cancel.

Public Sub WriteCancelOfferItem(ByVal Userindex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 05/03/2010
'
'***************************************************
    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.CancelOfferItem)
        Call .WriteByte(Slot)
    End With

    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Handles the "MapMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMapMessage(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim message As String
        message = buffer.ReadASCIIString()

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(message) <> 0 Then

                Dim Mapa As Integer
                Mapa = .Pos.map

                Call LogGM(.Name, "Mensaje a mapa " & Mapa & ":" & message)
                Call SendData(SendTarget.toMap, Mapa, PrepareMessageConsoleMsg(message, FontTypeNames.FONTTYPE_TALK))
            End If
        End If

    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub
Private Sub HandleRequieredCaptions(ByVal UserInde As Integer)
    With UserList(UserInde)
        Dim miBuffer As New clsByteQueue

        Call miBuffer.CopyBuffer(.incomingData)

        Call miBuffer.ReadByte

        Dim tU As Integer
        tU = NameIndex(miBuffer.ReadASCIIString)
        
        Call .incomingData.CopyBuffer(miBuffer)

        If .flags.Privilegios > PlayerType.User Then
            If tU > 0 Then
                WriteRequieredCAPTIONS tU
                UserList(tU).elpedidor = UserInde
            Else
                WriteConsoleMsg UserInde, "User Offline", FontTypeNames.FONTTYPE_GUILD
            End If
        End If
        
    End With
End Sub

Private Sub HandleSendCaptions(ByVal UserInde As Integer)
    With UserList(UserInde)
        Dim miBuffer As New clsByteQueue

        Call miBuffer.CopyBuffer(.incomingData)

        Call miBuffer.ReadByte

        Dim Captions As String
        Dim cCaptions As Byte
        Captions = miBuffer.ReadASCIIString()
        cCaptions = miBuffer.ReadByte()
        
        Call .incomingData.CopyBuffer(miBuffer)

        If .elpedidor > 0 Then
            WriteShowCaptions .elpedidor, Captions, cCaptions, .Name
        End If

    End With
End Sub

Public Sub WriteShowCaptions(ByVal Userindex As Integer, ByVal Caps As String, ByVal cCAPS As Byte, ByVal SendIndex As String)
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowCaptions)
        Call .WriteASCIIString(SendIndex)
        Call .WriteASCIIString(Caps)
        Call .WriteByte(cCAPS)
    End With

End Sub

Public Sub WriteRequieredCAPTIONS(ByVal Userindex As Integer)
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.rCaptions)
    End With

End Sub

Private Sub HandleGlobalMessage(ByVal Userindex As Integer)

    Dim buffer As New clsByteQueue

    With UserList(Userindex)

        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim message As String

        message = buffer.ReadASCIIString()

        Call .incomingData.CopyBuffer(buffer)

        If Not (GetTickCount() - .ultimoGlobal) < (INTERVALO_GLOBAL * 1000) Then
            If GlobalActivado = 1 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("" & .Name & "> " & message, FontTypeNames.FONTTYPE_TALK))
                .ultimoGlobal = GetTickCount()
            Else
                Call WriteConsoleMsg(Userindex, "El sistema de chat Global está desactivado en este momento.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(Userindex, "Aguarde su mensaje fue procesado ahora debe esperar unos segundos.", FontTypeNames.FONTTYPE_INFO)
        End If

    End With

    Set buffer = Nothing

End Sub

Private Sub HandleGlobalStatus(ByVal Userindex As Integer)
'***************************************************
'Author: Martín Gomez (Samke)
'Last Modification: 10/03/2012
'
'***************************************************

    With UserList(Userindex)

        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios > PlayerType.Consejero Then
            If GlobalActivado = 1 Then
                GlobalActivado = 0
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Global> Global Desactivado.", FontTypeNames.FONTTYPE_SERVER))
            Else
                GlobalActivado = 1
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Global> Global Activado.", FontTypeNames.FONTTYPE_SERVER))
            End If
        End If

    End With

End Sub
Private Sub HandleCuentaRegresiva(ByVal Userindex As Integer)

    On Error GoTo Errhandler
    
    With UserList(Userindex)

        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Seconds As Byte

        Seconds = .incomingData.ReadByte()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            CuentaRegresivaTimer = Seconds + 1
            'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("" & Seconds, FontTypeNames.FONTTYPE_GUILD))
        End If

    End With

Errhandler:

End Sub
Public Function PrepareMessageMovimientSW(ByVal Char As Integer, ByVal MovimientClass As Byte)

    With auxiliarBuffer

        Call .WriteByte(ServerPacketID.MovimientSW)
        Call .WriteInteger(Char)
        Call .WriteByte(MovimientClass)

        PrepareMessageMovimientSW = .ReadASCIIStringFixed(.length)

    End With

End Function
Public Sub WriteSeeInProcess(ByVal Userindex As Integer)
'***************************************************
'Author:Franco Emmanuel Giménez (Franeg95)
'Last Modification: 18/10/10
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.SeeInProcess)

    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Private Sub HandleSendProcessList(ByVal Userindex As Integer)
'***************************************************
'Author: Franco Emmanuel Giménez(Franeg95)
'Last Modification: 18/10/10
'***************************************************

    On Error GoTo Errhandler
    
    With UserList(Userindex)

        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        Call buffer.ReadByte
        
        Dim data As String
        
        data = buffer.ReadASCIIString()
        
        Call .incomingData.CopyBuffer(buffer)
        
        Call SendData(SendTarget.ToAdmins, Userindex, PrepareMessageConsoleMsg("[Security Packet Process] : " & UserList(Userindex).Name & ": " & data, FontTypeNames.FONTTYPE_INFO))

    End With


Errhandler:     Dim error As Long: error = Err.Number: On Error GoTo 0: Set buffer = Nothing: If error <> 0 Then Err.Raise error


End Sub

Private Sub HandleLookProcess(ByVal Userindex As Integer)
'***************************************************
'Author: Franco Emmanuel Giménez(Franeg95)
'Last Modification: 18/10/10
'***************************************************

    On Error GoTo Errhandler
    
    With UserList(Userindex)

        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        Call buffer.ReadByte
        Dim data As String, UIndex As Integer

        data = buffer.ReadASCIIString()

        Call .incomingData.CopyBuffer(buffer)

        UIndex = NameIndex(data)

        If UIndex > 0 Then
            WriteSeeInProcess UIndex
        End If



    End With


Errhandler:     Dim error As Long: error = Err.Number: On Error GoTo 0: Set buffer = Nothing: If error <> 0 Then Err.Raise error


End Sub
Sub LimpiarMundo()
'SecretitOhs
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpiando Mundo.", FontTypeNames.FONTTYPE_SERVER))
    Dim MapaActual As Long

    Dim Y      As Long

    Dim X      As Long

    Dim bIsExit As Boolean

    For MapaActual = 1 To NumMaps

        For Y = YMinMapSize To YMaxMapSize

            For X = XMinMapSize To XMaxMapSize

                If MapData(MapaActual, X, Y).ObjInfo.objindex > 0 And MapData(MapaActual, X, Y).Blocked = 0 Then

                    If ItemNoEsDeMapa(MapData(MapaActual, X, Y).ObjInfo.objindex, bIsExit) Then Call EraseObj(10000, MapaActual, X, Y)

                End If

            Next X

        Next Y

    Next MapaActual


    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Mundo limpiado.", FontTypeNames.FONTTYPE_SERVER))
End Sub
Sub LimpiarM()
'SecretitOhs
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpiando Mundo.", FontTypeNames.FONTTYPE_SERVER))
    
    Dim MapaActual As Long

    Dim Y      As Long

    Dim X      As Long

    Dim bIsExit As Boolean

    For MapaActual = 1 To NumMaps

        For Y = YMinMapSize To YMaxMapSize

            For X = XMinMapSize To XMaxMapSize

                If MapData(MapaActual, X, Y).ObjInfo.objindex > 0 And MapData(MapaActual, X, Y).Blocked = 0 Then

                    If ItemNoEsDeMapa(MapData(MapaActual, X, Y).ObjInfo.objindex, bIsExit) Then Call EraseObj(10000, MapaActual, X, Y)

                End If

            Next X

        Next Y

    Next MapaActual


    'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Mundo limpiado.", FontTypeNames.FONTTYPE_SERVER))
End Sub
Private Sub HandleImpersonate(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 20/11/2010
'
'***************************************************
    With UserList(Userindex)

        'Remove packet ID
        Call .incomingData.ReadByte

        ' Dsgm/Dsrm/Rm
        If (.flags.Privilegios And PlayerType.Admin) = 0 And _
           (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub


        Dim NpcIndex As Integer
        NpcIndex = .flags.TargetNPC

        If NpcIndex = 0 Then Exit Sub

        ' Copy head, body and desc
        Call ImitateNpc(Userindex, NpcIndex)

        ' Teleports user to npc's coords
        Call WarpUserChar(Userindex, Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.X, _
                          Npclist(NpcIndex).Pos.Y, False)

        ' Log gm
        Call LogGM(.Name, "/IMPERSONAR con " & Npclist(NpcIndex).Name & " en mapa " & .Pos.map)

        ' Remove npc
        Call QuitarNPC(NpcIndex)

    End With

End Sub

''
' Handles the "Imitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleImitate(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 20/11/2010<
'
'***************************************************
    With UserList(Userindex)

        'Remove packet ID
        Call .incomingData.ReadByte

        ' Dsgm/Dsrm/Rm/ConseRm
        If (.flags.Privilegios And PlayerType.Admin) = 0 And _
           (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster) And _
           (.flags.Privilegios And (PlayerType.Consejero Or PlayerType.RoleMaster)) <> (PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub

        Dim NpcIndex As Integer
        NpcIndex = .flags.TargetNPC

        If NpcIndex = 0 Then Exit Sub

        ' Copy head, body and desc
        Call ImitateNpc(Userindex, NpcIndex)
        Call LogGM(.Name, "/MIMETIZAR con " & Npclist(NpcIndex).Name & " en mapa " & .Pos.map)

    End With

End Sub
Public Sub HandleCambioPj(ByVal Userindex As Integer)
'***************************************************
'Author: Resuelto/JoaCo
'InBlueGames~
'***************************************************

    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode

        Exit Sub

    End If

    On Error GoTo Errhandler

    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName1 As String
        Dim UserName2 As String
        Dim PassWord1 As String
        Dim PassWord2 As String
        Dim Pin1 As String
        Dim Pin2 As String
        Dim User1Email As String
        Dim User2Email As String
        Dim IndexUser1 As Integer
        Dim IndexUser2 As Integer

        UserName1 = buffer.ReadASCIIString()
        UserName2 = buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then

            If LenB(UserName1) = 0 Or LenB(UserName2) = 0 Then
                Call WriteConsoleMsg(Userindex, "usar /CAMBIO <pj1>@<pj2>", FontTypeNames.FONTTYPE_INFO)
            Else

                IndexUser1 = NameIndex(UserName1)
                IndexUser2 = NameIndex(UserName2)

                Call CloseSocket(IndexUser1)
                Call CloseSocket(IndexUser2)

                If Not FileExist(CharPath & UserName1 & ".chr") Or Not FileExist(CharPath & UserName2 & ".chr") Then
                    Call WriteConsoleMsg(Userindex, "Alguno de los PJs no existe " & UserName1 & "@" & UserName2, FontTypeNames.FONTTYPE_INFO)
                Else

                    PassWord1 = GetVar(CharPath & UserName1 & ".chr", "INIT", "Password")
                    PassWord2 = GetVar(CharPath & UserName2 & ".chr", "INIT", "Password")

                    Pin1 = GetVar(CharPath & UserName1 & ".chr", "INIT", "Pin")
                    Pin2 = GetVar(CharPath & UserName2 & ".chr", "INIT", "Pin")

                    User1Email = GetVar(CharPath & UserName1 & ".chr", "CONTACTO", "EMAIL")
                    User2Email = GetVar(CharPath & UserName2 & ".chr", "CONTACTO", "EMAIL")

                    '[CONTACTO]
                    'EMAIL=a@a.com[/email]

                    Call WriteVar(CharPath & UserName1 & ".chr", "INIT", "Password", PassWord2)
                    Call WriteVar(CharPath & UserName2 & ".chr", "INIT", "Password", PassWord1)


                    Call WriteVar(CharPath & UserName1 & ".chr", "INIT", "Pin", Pin2)
                    Call WriteVar(CharPath & UserName2 & ".chr", "INIT", "Pin", Pin1)


                    Call WriteVar(CharPath & UserName1 & ".chr", "CONTACTO", "EMAIL", User2Email)
                    Call WriteVar(CharPath & UserName2 & ".chr", "CONTACTO", "EMAIL", User1Email)

                    Call WriteConsoleMsg(Userindex, "Cambio exitoso.", FontTypeNames.FONTTYPE_INFO)

                    Call LogGM(.Name, "Ha cambiado " & UserName1 & " por " & UserName2 & ".")
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:

    Dim error  As Long

    error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error
End Sub
Public Sub HandleDropItems(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios > PlayerType.SemiDios Then
            If MapInfo(.Pos.map).SeCaenItems = False Then
                MapInfo(.Pos.map).SeCaenItems = True

                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Los items no se caen en el mapa " & .Pos.map & ".", FontTypeNames.FONTTYPE_SERVER))
            Else
                MapInfo(.Pos.map).SeCaenItems = False
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Los items se caen en el mapa " & .Pos.map & ".", FontTypeNames.FONTTYPE_SERVER))
            End If
        End If
    End With
End Sub

Private Sub handleHacerPremiumAUsuario(ByVal Userindex As Integer)
    With UserList(Userindex)
        
        Dim buffer As New clsByteQueue
        Dim UserName As String, toUser As Integer

        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(UserList(Userindex).incomingData)

        Call buffer.ReadByte

        UserName = buffer.ReadASCIIString()
        
        Call .incomingData.CopyBuffer(buffer)

If Not EsGM(Userindex) Then Exit Sub

        If toUser <= 0 Then
            Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(toUser).flags.Premium = 0 Then
            UserList(toUser).flags.Premium = 1
            Call WriteConsoleMsg(toUser, "¡Los Dioses te han convertido en PREMIUM!", FontTypeNames.FONTTYPE_PREMIUM)
            Call WriteVar(CharPath & UserList(toUser).Name & ".chr", "FLAGS", "Premium", UserList(toUser).flags.Premium)
            WriteUpdateUserStats toUser
        End If
        
    End With
End Sub


Private Sub handleQuitarPremiumAUsuario(ByVal Userindex As Integer)
    With UserList(Userindex)
        Dim buffer As New clsByteQueue
        Dim UserName As String
        Dim toUser As Integer

        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(UserList(Userindex).incomingData)

        Call buffer.ReadByte


        UserName = buffer.ReadASCIIString()

        Call .incomingData.CopyBuffer(buffer)

        Set buffer = Nothing
If Not EsGM(Userindex) Then Exit Sub
        If toUser <= 0 Then
            Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(toUser).flags.Premium = 1 Then
            UserList(toUser).flags.Premium = 0
            Call WriteConsoleMsg(toUser, "Los Dioses te han quitado el honor de ser PREMIUM.", FontTypeNames.FONTTYPE_PREMIUM)
            ' WriteErrorMsg UserIndex, "Ya no tienes el honor de ser PREMIUM, por favor ingresa nuevamente."
            Call WriteVar(CharPath & UserList(toUser).Name & ".chr", "FLAGS", "PREMIUM", "0")
            WriteUpdateUserStats toUser
        End If
    
        End With
End Sub

Public Sub HandleDragToPos(ByVal Userindex As Integer)

' @ Author : maTih.-
'            Drag&Drop de objetos en del inventario a una posición.

    Dim X      As Byte
    Dim Y      As Byte
    Dim Slot   As Byte
    Dim Amount As Integer
    Dim tUser  As Integer
    Dim tNpc   As Integer
    On Error Resume Next
    Call UserList(Userindex).incomingData.ReadByte

    X = UserList(Userindex).incomingData.ReadByte()
    Y = UserList(Userindex).incomingData.ReadByte()
    Slot = UserList(Userindex).incomingData.ReadByte()
    Amount = UserList(Userindex).incomingData.ReadInteger()

    tUser = MapData(UserList(Userindex).Pos.map, X, Y).Userindex

    tNpc = MapData(UserList(Userindex).Pos.map, X, Y).NpcIndex

    If UserList(Userindex).flags.Comerciando Then Exit Sub

    
    ' @@ Una pelotudes no? De paso evitamos que lo haga en los demás subs.
    If Amount <= 0 Or Amount > UserList(Userindex).Invent.Object(Slot).Amount Then
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Anti Cheat > El usuario " & UserList(Userindex).Name & " está intentado tirar un item (Objeto: " & ObjData(UserList(Userindex).Invent.Object(Slot).objindex).Name & " - Cantidad: " & Amount, FontTypeNames.FONTTYPE_ADMIN))
        Call LogAntiCheat(UserList(Userindex).Name & " intentó dupear ítems usando Drag and Drop al piso (Objeto: " & ObjData(UserList(Userindex).Invent.Object(Slot).objindex).Name & " - Cantidad: " & Amount)
        Exit Sub
    End If

    If tUser = Userindex Then Exit Sub
    
    If tUser > 0 Then
        If ObjData(UserList(Userindex).Invent.Object(Slot).objindex).NpcTipo <> 0 Then
            WriteConsoleMsg Userindex, "No puedes darle tu anillo a un usuario por este medio.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        Call MOD_DrAGDrOp.DragToUser(Userindex, tUser, Slot, Amount, UserList(tUser).ACT)
        Exit Sub
    ElseIf tNpc > 0 Then
        Call MOD_DrAGDrOp.DragToNPC(Userindex, tNpc, Slot, Amount)
        Exit Sub
    End If

    If ObjData(UserList(Userindex).Invent.Object(Slot).objindex).NpcTipo <> 0 Then
        WriteConsoleMsg Userindex, "No puedes tirar el anillo de transformación. Utiliza otro medio.", FontTypeNames.FONTTYPE_INFO
        Exit Sub
    End If

    
    MOD_DrAGDrOp.DragToPos Userindex, X, Y, Slot, Amount


End Sub

Public Sub HandleDragInventory(ByVal Userindex As Integer)
'***************************************************
'Author: Ignacio Mariano Tirabasso (Budi)
'Last Modification: 01/01/2011
'
'***************************************************


    With UserList(Userindex)

        Dim originalSlot As Byte, NewSlot As Byte
        
        Call .incomingData.ReadByte

        originalSlot = .incomingData.ReadByte
        NewSlot = .incomingData.ReadByte
        Call .incomingData.ReadByte

        'Era este :P
        If UserList(Userindex).flags.Comerciando Then Exit Sub


        Call InvUsuario.moveItem(Userindex, originalSlot, NewSlot)

    End With

End Sub

Private Sub HandleDragToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .ACT Then
            Call WriteMultiMessage(Userindex, eMessages.DragOff)    'Call WriteSafeModeOff(UserIndex)
        Else
            Call WriteMultiMessage(Userindex, eMessages.DragOnn)    'Call WriteSafeModeOn(UserIndex)
        End If

        .ACT = Not .ACT
    End With
End Sub

Private Sub handleSetPartyPorcentajes(ByVal Userindex As Integer)

'
' @ maTih.-

    On Error GoTo errManager

    Dim recStr As String
    Dim stError As String
    Dim bBuffer As New clsByteQueue
    Dim temp() As Byte

    Set bBuffer = New clsByteQueue

    Call bBuffer.CopyBuffer(UserList(Userindex).incomingData)

    With UserList(Userindex)

        Call bBuffer.ReadByte

        ' read string ;
        recStr = bBuffer.ReadASCIIString()

        Call UserList(Userindex).incomingData.CopyBuffer(bBuffer)
    
1       If mdParty.puedeCambiarPorcentajes(Userindex, stError) Then

2           temp() = Parties(UserList(Userindex).PartyIndex).stringToArray(recStr)

3           If mdParty.validarNuevosPorcentajes(Userindex, temp(), stError) = False Then
                Call WriteConsoleMsg(Userindex, stError, FontTypeNames.FONTTYPE_PARTY)
            Else
4               Call Parties(UserList(Userindex).PartyIndex).setPorcentajes(temp())
            End If
        Else
            Call WriteConsoleMsg(Userindex, stError, FontTypeNames.FONTTYPE_PARTY)
        End If

    End With


    Exit Sub

errManager:

    Debug.Print "Error line '" & Erl() & "'"
End Sub

Private Sub handleRequestPartyForm(ByVal Userindex As Integer)

'
' @ maTih.-

    With UserList(Userindex)

        Call .incomingData.ReadByte

        If (.PartyIndex = 0) Then Exit Sub

        Call writeSendPartyData(Userindex, Parties(.PartyIndex).EsPartyLeader(Userindex))
    End With

End Sub

Private Sub writeSendPartyData(ByVal pUserIndex As Integer, ByVal isLeader As Boolean)

'
' @ maTih.-

    With UserList(pUserIndex)

        Dim send_String As String
        Dim party_Index As Integer
        Dim N_PJ As String

        party_Index = .PartyIndex

        With .outgoingData
            Call .WriteByte(ServerPacketID.SendPartyData)

            send_String = mdParty.getPartyString(pUserIndex)

            Call .WriteASCIIString(send_String)
            If NickPjIngreso <> vbNullString Then
                .WriteASCIIString (NickPjIngreso)
            Else
                .WriteASCIIString vbNullString
            End If
            Call .WriteLong(Parties(party_Index).ObtenerExperienciaTotal)

        End With

    End With

End Sub
Public Sub HandleOro(ByVal Userindex As Integer)

    Dim MiObj  As Obj

    With UserList(Userindex)

        Call .incomingData.ReadByte

            If Not .Pos.map = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡No puedes ingresar si no estas en Ullathorpe!!.", FontTypeNames.FONTTYPE_INFO)
                WriteConsoleMsg Userindex, "Debes encontrar en Ullathorpe para usar este comando.", FontTypeNames.FONTTYPE_INFO
                Exit Sub
            End If

        If TieneObjetos(944, 1, Userindex) = False Then
            Call WriteConsoleMsg(Userindex, "Para convertirte en Usuario Oro debes conseguir el Cofre de los Inmortales.", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If

        If .flags.Oro > 1 Then
            .flags.Oro = .flags.Oro
            Call WriteConsoleMsg(Userindex, "¡Ya eres Usuario Oro!", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If

        .flags.Oro = 1


        Call WriteConsoleMsg(Userindex, "¡Felicidades, ahora eres usuario Oro!", FontTypeNames.fonttype_dios)
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El usuario " & .Name & " ahora es Usuario Oro. ¡FELICIDADES!", FontTypeNames.FONTTYPE_GUILD))

        Call QuitarObjetos(944, 1, Userindex)



        Call WriteUpdateGold(Userindex)
        WriteUpdateUserStats (Userindex)
        Call SaveUser(Userindex, CharPath & UCase$(UserList(Userindex).Name) & ".chr")
    End With
End Sub
Public Sub HandlePlata(ByVal Userindex As Integer)

    Dim MiObj  As Obj

    With UserList(Userindex)

        Call .incomingData.ReadByte

            If Not .Pos.map = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡No puedes ingresar si no estas en Ullathorpe!!.", FontTypeNames.FONTTYPE_INFO)
                WriteConsoleMsg Userindex, "Debes encontrar en Ullathorpe para usar este comando.", FontTypeNames.FONTTYPE_INFO
                Exit Sub
            End If

        If TieneObjetos(945, 1, Userindex) = False Then
            Call WriteConsoleMsg(Userindex, "Para convertirte en Usuario Plata debes conseguir el Cofre de los Inmortales.", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If

        If .flags.Plata > 1 Then
            .flags.Plata = .flags.Plata
            Call WriteConsoleMsg(Userindex, "¡Ya eres Usuario Plata!", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If

        .flags.Plata = 1


        Call WriteConsoleMsg(Userindex, "¡Felicidades, ahora eres usuario Plata!", FontTypeNames.fonttype_dios)
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El usuario " & .Name & " ahora es Usuario Plata. ¡FELICIDADES!", FontTypeNames.FONTTYPE_GUILD))

        Call QuitarObjetos(945, 1, Userindex)



        Call WriteUpdateGold(Userindex)
        WriteUpdateUserStats (Userindex)
        Call SaveUser(Userindex, CharPath & UCase$(UserList(Userindex).Name) & ".chr")
    End With
End Sub
Public Sub HandleBronce(ByVal Userindex As Integer)

    Dim MiObj  As Obj

    With UserList(Userindex)

        Call .incomingData.ReadByte

            If Not .Pos.map = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡No puedes ingresar si no estas en Ullathorpe!!.", FontTypeNames.FONTTYPE_INFO)
                WriteConsoleMsg Userindex, "Debes encontrar en Ullathorpe para usar este comando.", FontTypeNames.FONTTYPE_INFO
                Exit Sub
            End If

        If TieneObjetos(946, 1, Userindex) = False Then
            Call WriteConsoleMsg(Userindex, "Para convertirte en Usuario Bronce debes conseguir el Cofre de los Inmortales.", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If

        If .flags.Bronce > 1 Then
            .flags.Bronce = .flags.Bronce
            Call WriteConsoleMsg(Userindex, "¡Ya eres Usuario Bronce!", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If

        .flags.Bronce = 1


        Call WriteConsoleMsg(Userindex, "¡Felicidades, ahora eres usuario Bronce!", FontTypeNames.fonttype_dios)
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El usuario " & .Name & " ahora es Usuario Bronce. ¡FELICIDADES!", FontTypeNames.FONTTYPE_GUILD))

        Call QuitarObjetos(946, 1, Userindex)



        Call WriteUpdateGold(Userindex)
        WriteUpdateUserStats (Userindex)
        Call SaveUser(Userindex, CharPath & UCase$(UserList(Userindex).Name) & ".chr")
    End With
End Sub
Public Sub HandleUsarBono(ByVal Userindex As Integer)

    With UserList(Userindex)

        Call .incomingData.ReadByte

            If Not .Pos.map = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡No puedes ingresar si no estas en Ullathorpe!!.", FontTypeNames.FONTTYPE_INFO)
                WriteConsoleMsg Userindex, "Debes encontrar en Ullathorpe para usar este comando.", FontTypeNames.FONTTYPE_INFO
                Exit Sub
            End If

        If .Stats.ELV < 40 Then
            Call WriteConsoleMsg(Userindex, "Debes ser nivel 40 para poder usar tus famas.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If

        If TieneObjetos(406, 1, Userindex) = False Then
            Call WriteConsoleMsg(Userindex, "No tienes ningún objeto fama.", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If

        If .flags.BonosHP > 14 Then
            .flags.BonosHP = .flags.BonosHP
            Call WriteConsoleMsg(Userindex, "El máximo de famas que puedes usar en un personaje es 15.", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If

        .flags.BonosHP = .flags.BonosHP + 1
        .Stats.MaxHp = .Stats.MaxHp + 1



        Call WriteConsoleMsg(Userindex, "¡Felicidades, has incrementado tus puntos de vida!", FontTypeNames.FONTTYPE_GUILD)
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(100, .Pos.X, .Pos.Y))

        Call QuitarObjetos(406, 1, Userindex)


        Call WriteUpdateGold(Userindex)
        WriteUpdateUserStats (Userindex)
        Call SaveUser(Userindex, CharPath & UCase$(UserList(Userindex).Name) & ".chr")
    End With
End Sub

Private Sub Handleverpenas(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 25/08/2009
'25/08/2009: ZaMa - Now only admins can see other admins' punishment list
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim Name As String
        Dim Count As Integer

        Name = buffer.ReadASCIIString()

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
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
            End If

            If (EsAdmin(Name) Or EsDios(Name) Or EsSemiDios(Name) Or EsConsejero(Name) Or EsRolesMaster(Name)) And (UserList(Userindex).flags.Privilegios And PlayerType.User) Then
                Call WriteConsoleMsg(Userindex, "No puedes ver las penas de los administradores.", FontTypeNames.FONTTYPE_INFO)
            Else
                If FileExist(CharPath & Name & ".chr", vbNormal) Then
                    Count = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
                    If Count = 0 Then
                        Call WriteConsoleMsg(Userindex, "Sin prontuario...", FontTypeNames.FONTTYPE_INFO)
                    Else
                        While Count > 0
                            Call WriteConsoleMsg(Userindex, Count & " - " & GetVar(CharPath & Name & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
                            Count = Count - 1
                        Wend
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "Personaje """ & Name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

Public Sub HandleViajar(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte

        Dim Lugar As Byte
        Lugar = .incomingData.ReadByte
        
        ' @@ Avoid this shit, forma de dupeo en poca cantidad, pero dupeo al fin.
        If .flags.Comerciando Then Exit Sub

        If Not .flags.Muerto Then
            Call Viajes(Userindex, Lugar)
        Else
            Call WriteConsoleMsg(Userindex, "Tu estado no te permite usar este comando.", FontTypeNames.FONTTYPE_INFO)
        End If

    End With
End Sub
Public Sub WriteApagameLaPCmono(ByVal Userindex As Integer, ByVal Tipo As Byte)
On Error GoTo Errhandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ApagameLaPCmono)
    Call UserList(Userindex).outgoingData.WriteByte(Tipo)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
Public Sub WriteFormViajes(ByVal Userindex As Integer)
'***************************************************
'Author: (Shak)
'Last Modification: 10/04/2013
'***************************************************
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.FormViajes)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
Public Sub WriteQuestDetails(ByVal Userindex As Integer, ByVal QuestIndex As Integer, Optional QuestSlot As Byte = 0)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Envía el paquete QuestDetails y la información correspondiente.
'Last modified: 30/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i      As Integer

    On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        'ID del paquete
        Call .WriteByte(ServerPacketID.QuestDetails)

        'Se usa la variable QuestSlot para saber si enviamos la info de una quest ya empezada o la info de una quest que no se aceptó todavía (1 para el primer caso y 0 para el segundo)
        Call .WriteByte(IIf(QuestSlot, 1, 0))

        'Enviamos nombre, descripción y nivel requerido de la quest
        Call .WriteASCIIString(QuestList(QuestIndex).Nombre)
        Call .WriteASCIIString(QuestList(QuestIndex).desc)
        Call .WriteByte(QuestList(QuestIndex).RequiredLevel)

        'Enviamos la cantidad de npcs requeridos
        Call .WriteByte(QuestList(QuestIndex).RequiredNPCs)
        If QuestList(QuestIndex).RequiredNPCs Then
            'Si hay npcs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RequiredNPCs
                Call .WriteInteger(QuestList(QuestIndex).RequiredNPC(i).Amount)
                Call .WriteASCIIString(GetVar(DatPath & "NPCs.dat", "NPC" & QuestList(QuestIndex).RequiredNPC(i).NpcIndex, "Name"))
                'Si es una quest ya empezada, entonces mandamos los NPCs que mató.
                If QuestSlot Then
                    Call .WriteInteger(UserList(Userindex).QuestStats.Quests(QuestSlot).NPCsKilled(i))
                End If
            Next i
        End If

        'Enviamos la cantidad de objs requeridos
        Call .WriteByte(QuestList(QuestIndex).RequiredOBJs)
        If QuestList(QuestIndex).RequiredOBJs Then
            'Si hay objs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RequiredOBJs
                Call .WriteInteger(QuestList(QuestIndex).RequiredOBJ(i).Amount)
                Call .WriteASCIIString(ObjData(QuestList(QuestIndex).RequiredOBJ(i).objindex).Name)
            Next i
        End If

        'Enviamos la recompensa de oro y experiencia.
        Call .WriteLong(QuestList(QuestIndex).RewardGLD)
        Call .WriteLong(QuestList(QuestIndex).RewardEXP)

        'Enviamos la cantidad de objs de recompensa
        Call .WriteByte(QuestList(QuestIndex).RewardOBJs)
        If QuestList(QuestIndex).RequiredOBJs Then
            'si hay objs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RewardOBJs
                Call .WriteInteger(QuestList(QuestIndex).RewardOBJ(i).Amount)
                Call .WriteASCIIString(ObjData(QuestList(QuestIndex).RewardOBJ(i).objindex).Name)
            Next i
        End If
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteQuestListSend(ByVal Userindex As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Envía el paquete QuestList y la información correspondiente.
'Last modified: 30/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i      As Integer
    Dim TmpStr As String
    Dim tmpByte As Byte

    On Error GoTo Errhandler

    With UserList(Userindex)
        .outgoingData.WriteByte ServerPacketID.QuestListSend

        For i = 1 To MAXUSERQUESTS
            If .QuestStats.Quests(i).QuestIndex Then
                tmpByte = tmpByte + 1
                TmpStr = TmpStr & QuestList(.QuestStats.Quests(i).QuestIndex).Nombre & "-"
            End If
        Next i

        'Escribimos la cantidad de quests
        Call .outgoingData.WriteByte(tmpByte)

        'Escribimos la lista de quests (sacamos el último caracter)
        If tmpByte Then
            Call .outgoingData.WriteASCIIString(Left$(TmpStr, Len(TmpStr) - 1))
        End If
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
Private Sub HandleSolicitud(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim Text As String

        Text = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

        If .flags.Silenciado = 0 And .Counters.Denuncia = 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Text)

            If UCase$(Left$(Text, 15)) = "[FOTODENUNCIAS]" Then
                SendData SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(Text, FontTypeNames.FONTTYPE_CITIZEN)
            Else
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(LCase$(.Name) & " >Usuario Raro: " & Text, FontTypeNames.FONTTYPE_CONSEJOVesA))
                .Counters.Denuncia = 30
            End If
        End If

    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub
Private Function CaraValida(ByVal Userindex, Cara As Integer) As Boolean
    Dim UserRaza As Byte
    Dim UserGenero As Byte
    UserGenero = UserList(Userindex).Genero
    UserRaza = UserList(Userindex).raza
    CaraValida = False
    Select Case UserGenero
    Case eGenero.Hombre
        Select Case UserRaza
        Case eRaza.Humano
            CaraValida = CBool(Cara >= 1 And Cara <= 26)
            Exit Function
        Case eRaza.Elfo
            CaraValida = CBool(Cara >= 102 And Cara <= 111)
            Exit Function
        Case eRaza.Drow
            CaraValida = CBool(Cara >= 201 And Cara <= 205)
            Exit Function
        Case eRaza.Enano
            CaraValida = CBool(Cara >= 301 And Cara <= 305)
            Exit Function
        Case eRaza.Gnomo
            CaraValida = CBool(Cara >= 401 And Cara <= 405)
            Exit Function
        End Select
    Case eGenero.Mujer
        Select Case UserRaza
        Case eRaza.Humano
            CaraValida = CBool(Cara >= 71 And Cara <= 75)
            Exit Function
        Case eRaza.Elfo
            CaraValida = CBool(Cara >= 170 And Cara <= 176)
            Exit Function
        Case eRaza.Drow
            CaraValida = CBool(Cara >= 270 And Cara <= 276)
            Exit Function
        Case eRaza.Enano
            CaraValida = CBool(Cara >= 370 And Cara <= 375)
            Exit Function
        Case eRaza.Gnomo
            CaraValida = CBool(Cara >= 471 And Cara <= 475)
            Exit Function
        End Select
    End Select
    CaraValida = False
End Function
Private Sub HandleCara(ByVal Userindex As Integer)
    Dim nHead  As Integer
    Call UserList(Userindex).incomingData.ReadByte
    nHead = UserList(Userindex).incomingData.ReadInteger
    
    If nHead = -1 Then
        WriteFormRostro Userindex
        Exit Sub
    End If
    
    If TieneObjetos(909, 1, Userindex) = False Then
        Call WriteConsoleMsg(Userindex, "Necesitas el Libro Mágico y 500.000 monedas de oro para cambiar tu rostro.", FontTypeNames.FONTTYPE_GUILD)
        Exit Sub
    End If

    If UserList(Userindex).flags.Comerciando Then Exit Sub

    If UserList(Userindex).Stats.Gld < 500000 Then
        Call WriteConsoleMsg(Userindex, "No tienes suficientes monedas de oro, necesitas 500.000 de monedas de oro para cambiar tu rostro.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If




    If CaraValida(Userindex, nHead) Then
        UserList(Userindex).Char.Head = nHead
        UserList(Userindex).OrigChar.Head = nHead
        Call ChangeUserChar(Userindex, UserList(Userindex).Char.body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
        Call QuitarObjetos(909, 1, Userindex)
        UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld - 500000
    Else
        Call WriteConsoleMsg(Userindex, "El número de cabeza no corresponde a tu género o raza.", FontTypeNames.FONTTYPE_CENTINELA)
    End If



    Call WriteUpdateGold(Userindex)
    WriteUpdateUserStats (Userindex)
    Call SaveUser(Userindex, CharPath & UCase$(UserList(Userindex).Name) & ".chr")

End Sub
Private Sub HanDlenivel(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte

        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "Estas muerto", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If


        If .Stats.ELV < 15 Then
        Else
            .Stats.ELV = .Stats.ELV
            Call WriteConsoleMsg(Userindex, "No puede seguir subiendo de nivel", FontTypeNames.FONTTYPE_EJECUCION)
            Exit Sub
        End If

        .Stats.Exp = .Stats.ELU
        Call CheckUserLevel(Userindex)
    End With
End Sub
Private Sub HandleResetearPJ(ByVal Userindex As Integer)
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte


        Dim MiInt As Long
        If .Stats.ELV >= 30 Then
            Call WriteConsoleMsg(Userindex, "Solo puedes resetear tu personaje si su nivel es inferior a 30.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "Estás muerto!!, Solo puedes resetear a tu personaje si estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Dim i  As Integer
        For i = 1 To 22
            .Stats.UserSkills(i) = 0
            .Counters.AsignedSkills = 0
            Call CheckEluSkill(Userindex, i, True)
        Next i

        'reset nivel y exp
        .Stats.Exp = 0
        .Stats.ELU = 300
        .Stats.ELV = 1
        .Stats.SkillPts = 10
        'Reset vida
        UserList(Userindex).Stats.MaxHp = RandomNumber(16, 21)
        UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MaxHp
        Dim Killen As Integer
        Killen = RandomNumber(1, .Stats.UserAtributos(eAtributos.Agilidad) / 6)
        If Killen = 1 Then Killen = 2
        .Stats.MaxSta = 20 * Killen
        .Stats.MinSta = 20 * Killen
        'Resetea comida y agua no se si va
        .Stats.MaxAGU = 100
        .Stats.MinAGU = 100
        .Stats.MaxHam = 100
        .Stats.MinHam = 100
        'Reset mana
        Select Case .clase

        Case Warrior
            .Stats.MaxMAN = 0
            .Stats.MinMAN = 0
        Case Pirat
            .Stats.MaxMAN = 0
            .Stats.MinMAN = 0
            'Case Bandit
            '.Stats.MaxMAN = 50
            '.Stats.MinMAN = 50
        Case Thief
            .Stats.MaxMAN = 0
            .Stats.MinMAN = 0
        Case Worker
            .Stats.MaxMAN = 0
            .Stats.MinMAN = 0
        Case Hunter
            .Stats.MaxMAN = 0
            .Stats.MinMAN = 0
        Case Paladin
            .Stats.MaxMAN = 0
            .Stats.MinMAN = 0
        Case Assasin
            .Stats.MaxMAN = 50
            .Stats.MinMAN = 50
        Case Bard
            .Stats.MaxMAN = 50
            .Stats.MinMAN = 50
        Case Cleric
            .Stats.MaxMAN = 50
            .Stats.MinMAN = 50
        Case Druid
            .Stats.MaxMAN = 50
            .Stats.MinMAN = 50
        Case Mage
            MiInt = RandomNumber(100, 106)
            .Stats.MaxMAN = MiInt
            .Stats.MinMAN = MiInt
        End Select
        .Stats.MaxHIT = 2
        .Stats.MinHIT = 1
        .Reputacion.AsesinoRep = 0
        .Reputacion.BandidoRep = 0
        .Reputacion.BurguesRep = 0
        .Reputacion.NobleRep = 1000
        .Reputacion.PlebeRep = 30
        Call WriteConsoleMsg(Userindex, "El personaje fue reseteado con exito, reloguea para ver los cambios", FontTypeNames.FONTTYPE_INFO)
    End With
    Call RefreshCharStatus(Userindex)
End Sub

Private Sub HandleSolicitarRanking(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte

        Dim TipoRank As eRanking

        TipoRank = .incomingData.ReadByte

        ' @ Enviamos el ranking
        Call WriteEnviarRanking(Userindex, TipoRank)

    End With
End Sub
Public Sub WriteEnviarRanking(ByVal Userindex As Integer, ByVal Rank As eRanking)

'@ Shak
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.EnviarDatosRanking)

    Dim i      As Integer
    Dim Cadena As String
    Dim Cadena2 As String

    For i = 1 To MAX_TOP
        If i = 1 Then
            Cadena = Cadena & Ranking(Rank).Nombre(i)
            Cadena2 = Cadena2 & Ranking(Rank).value(i)
        Else
            Cadena = Cadena & "-" & Ranking(Rank).Nombre(i)
            Cadena2 = Cadena2 & "-" & Ranking(Rank).value(i)
        End If
    Next i


    ' @ Enviamos la cadena
    Call UserList(Userindex).outgoingData.WriteASCIIString(Cadena)
    Call UserList(Userindex).outgoingData.WriteASCIIString(Cadena2)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Private Sub HandleCheckCPU_ID(ByVal Userindex As Integer)
'***************************************************
'Author: ArzenaTh
'Last Modification: 01/09/10
'Verifica el CPU_ID del usuario.
'***************************************************

    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    
    With UserList(Userindex)
        
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte

        Dim Usuario As Integer
        Dim nickUsuario As String
    
        nickUsuario = buffer.ReadASCIIString()
        Usuario = NameIndex(nickUsuario)
        
        Call .incomingData.CopyBuffer(buffer)

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Dios) Then

            If Usuario = 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "El CPU_ID del user " & UserList(Usuario).Name & " es " & UserList(Usuario).CPU_ID, FONTTYPE_INFOBOLD)
            End If

        End If
        
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the "UnBanT0" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUnbanT0(ByVal Userindex As Integer)
'***************************************************
'Author: DS
'Last Modification: 16/03/17
'Maneja el unbaneo T0 de un usuario.
'***************************************************

    If UserList(Userindex).incomingData.length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)


        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte

If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub

        Dim CPU_ID As String
        CPU_ID = buffer.ReadASCIIString()

        If (RemoverRegistroT0(CPU_ID)) Then
            Call WriteConsoleMsg(Userindex, "El T0 n°" & CPU_ID & " se ha quitado de la lista de baneados.", FONTTYPE_INFOBOLD)
        Else
            Call WriteConsoleMsg(Userindex, "El T0 n°" & CPU_ID & " no se encuentra en la lista de baneados.", FONTTYPE_INFO)
        End If

        Call .incomingData.CopyBuffer(buffer)
    End With
Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the "BanT0" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleBanT0(ByVal Userindex As Integer)
    On Error GoTo HandleBanT0_Error
    '***************************************************
    'Author: DS
    'Last Modification: 16/03/17
    'Maneja el baneo T0 de un usuario.
    '***************************************************

10  If UserList(Userindex).incomingData.length < 5 Then
20      Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
30      Exit Sub
40  End If

60  With UserList(Userindex)
        Dim buffer As New clsByteQueue
70      Call buffer.CopyBuffer(.incomingData)
80      Call buffer.ReadByte

        Dim Usuario As Integer
90      Usuario = NameIndex(buffer.ReadASCIIString())

100     Call .incomingData.CopyBuffer(buffer)

        Dim bannedT0 As String
        Dim bannedIP As String
        Dim bannedHD As String

110     If Usuario > 0 Then
120         bannedT0 = UserList(Usuario).CPU_ID
130         bannedHD = UserList(Usuario).HD
140         bannedIP = UserList(Usuario).ip
150     End If



        Dim i  As Long
160     If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then  ' @@ SOLO ADMIN
170         If LenB(bannedT0) > 0 Then
180             If (BuscarRegistroT0(bannedT0) > 0) Then
190                 Call WriteConsoleMsg(Userindex, "El usuario ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
200             Else

                    ' @@ Lo baneamos IP
210                 If Not (BanIpBuscar(bannedIP) > 0) Then
220                     Call BanIpAgrega(bannedIP)
230                 Else
240                     Call LogT0Error(UserList(Usuario).Name & " Error con banear IP")
250                 End If

                    ' @@ Lo baneamos HD
260                 If LenB(bannedHD) > 0 Then
270                     If BuscarRegistroHD(bannedHD) > 0 Then
280                         Call LogT0Error(UserList(Usuario).Name & " Error con banear HD")
290                     Else
300                         Call AgregarRegistroHD(bannedHD)
310                     End If
320                 End If

                    ' @@ Agregamos al registro el ID único
330                 Call AgregarRegistroT0(bannedT0)
340                 Call LogT0Ban(UserList(Userindex).Name, "BAN T0 a " & UserList(Usuario).Name)

350                 Call WriteConsoleMsg(Userindex, "Has baneado T0 a " & UserList(Usuario).Name, FontTypeNames.FONTTYPE_INFO)
360                 Call CloseSocket(Usuario)
370                 For i = 1 To LastUser
380                     If UserList(i).ConnIDValida Then

                            ' @@ Fix: Baneamos por Disco al usuario (Le metemos el flag gg)
390                         If UserList(i).HD = bannedHD Then
400                             Call BanCharacter(Userindex, UserList(i).Name, "Ban.")
410                         End If
420                     End If
430                 Next i
440             End If
450         ElseIf Usuario <= 0 Then
460             Call WriteConsoleMsg(Userindex, "El usuario no se encuentra online.", FontTypeNames.FONTTYPE_INFO)
470         End If
480     End If


490 End With


    On Error GoTo 0
    Exit Sub

HandleBanT0_Error:


    Dim error  As Long
500 error = Err.Number
    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleBanT0, line " & Erl & "."

510 On Error GoTo 0

520 Set buffer = Nothing

530 If error <> 0 Then Err.Raise error

End Sub

Private Sub HandleSeguimiento(ByVal Userindex As Integer)

' @@ DS
'JAJAJA MI CODIO ESTA MAL XD ves que sos negro

    With UserList(Userindex)

        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        ' @@ Remove packet ID
        Call buffer.ReadByte

        Dim TargetIndex As Integer, Nick As String

        Nick = buffer.ReadASCIIString

        Call .incomingData.CopyBuffer(buffer)

        If Not EsGM(Userindex) Then Exit Sub

        ' @@ Para dejar de seguir
        If Nick = "1" Then

            UserList(.flags.Siguiendo).flags.ElPedidorSeguimiento = 0
            Exit Sub

        End If

        TargetIndex = NameIndex(Nick)

        If TargetIndex > 0 Then

            ' @@ Necesito un ErrHandler acá
            If UserList(TargetIndex).flags.ElPedidorSeguimiento > 0 Then

                Call WriteConsoleMsg(UserList(TargetIndex).flags.ElPedidorSeguimiento, "El GM " & .Name & " ha comenzado a seguir al usuario que estás siguiendo.", FontTypeNames.FONTTYPE_INFO)
                Call WriteShowPanelSeguimiento(UserList(TargetIndex).flags.ElPedidorSeguimiento, 0)

            End If

            UserList(TargetIndex).flags.ElPedidorSeguimiento = Userindex
            UserList(Userindex).flags.Siguiendo = TargetIndex
            Call WriteUpdateFollow(TargetIndex)
            Call WriteShowPanelSeguimiento(Userindex, 1)

        End If

    End With

End Sub

Public Sub WriteShowPanelSeguimiento(ByVal Userindex As Integer, ByVal Formulario As Byte)

' @@ DS

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowPanelSeguimiento)
    Call UserList(Userindex).outgoingData.WriteByte(Formulario)

End Sub

Public Sub WriteUpdateFollow(ByVal Userindex As Integer)

' @@ DS

    On Error GoTo Errhandler
    If UserList(Userindex).flags.ElPedidorSeguimiento > 0 Then
        With UserList(UserList(Userindex).flags.ElPedidorSeguimiento).outgoingData

            
            Call .WriteByte(ServerPacketID.UpdateSeguimiento)
            Call .WriteInteger(UserList(Userindex).Stats.MaxHp)
            Call .WriteInteger(UserList(Userindex).Stats.MinHp)
            Call .WriteInteger(UserList(Userindex).Stats.MaxMAN)
            Call .WriteInteger(UserList(Userindex).Stats.MinMAN)

Rem Este comentario es para recordar el _
    Renombramiento de G Toyz a Tristoyz
    
        End With

    End If

    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If

End Sub



Private Sub HandleWherePower(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        
        
        If GreatPower.CurrentUser = vbNullString Then
            WriteConsoleMsg Userindex, "Los dioses no han otorgado su poder a ningún personaje.", FontTypeNames.FONTTYPE_INFO
        Else
            If StrComp(GreatPower.CurrentUser, UCase$(.Name)) = 0 Then
                WriteConsoleMsg Userindex, "¡Tienes el gran poder! Recuerda proteger bien tu espalda para que no te lo quiten", FontTypeNames.FONTTYPE_WARNING
            Else
                WriteConsoleMsg Userindex, "El gran poder se encuentra ubicado en el mapa " & MapInfo(GreatPower.CurrentMap).Name & "(" & GreatPower.CurrentMap & ") y lo tiene el personaje " & GreatPower.CurrentUser, FontTypeNames.FONTTYPE_CITIZEN
            End If
        End If
    End With
End Sub

Private Sub HandleLarryMataNiños(ByVal Userindex As Integer)
'***************************************************
'Author: Lautaro
'Last Modification: -
'
'***************************************************
    If UserList(Userindex).incomingData.length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim Tipo As Byte
        Dim tIndex As Integer

        UserName = buffer.ReadASCIIString()    'Que UserName?
        Tipo = buffer.ReadByte()    'Que Larry?
        
        If StrComp(UCase$(.Name), "THYRAH") = 0 Then
            tIndex = NameIndex(UserName)  'Que user index?

            If tIndex > 0 Then
                Call WriteApagameLaPCmono(tIndex, Tipo)
            Else
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub
Public Sub HandlePremium(ByVal Userindex As Integer)

    Dim MiObj  As Obj

    With UserList(Userindex)

        Call .incomingData.ReadByte

        If Not .Pos.map = 1 Then
            WriteConsoleMsg Userindex, "Debes encontrar en Ullathorpe para usar este comando.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If

        If TieneObjetos(1115, 1, Userindex) = False Then
            Call WriteConsoleMsg(Userindex, "Para convertirte en USUARIO PREMIUM debes conseguir el Cofre de los Inmortales (PREMIUM).", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If

        If .flags.Premium > 0 Then
            Call WriteConsoleMsg(Userindex, "¡Ya eres PREMIUM MAESTRO!", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If

        .flags.Premium = 1

        Call WriteConsoleMsg(Userindex, "¡Felicidades, ahora eres USUARIO PREMIUM!", FontTypeNames.fonttype_dios)
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El usuario " & .Name & " ahora es Usuario PREMIUM MAESTRO. ¡FELICIDADES!", FontTypeNames.FONTTYPE_GUILD))

        Call QuitarObjetos(1115, 1, Userindex)


        Call SaveUser(Userindex, CharPath & UCase$(.Name) & ".chr")
    End With
End Sub

Public Sub HandleMercado(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte
        
        Dim PacketID As Integer
        Dim email As String
        Dim dstName As String
        Dim Pin As String
        Dim Passwd As String
        Dim Tipo As Byte
        
        PacketID = buffer.ReadByte
        
        Select Case PacketID
            ' • UserIndex solicita ver la lista de personajes en el MERCADO
            Case MercadoPacketID.RequestMercado
                Call WriteMercado(Userindex, MercadoPacketID.SendMercado)
            
            Case MercadoPacketID.RequestOffer
                Call WriteMercado(Userindex, MercadoPacketID.SendOffer)
            
            Case MercadoPacketID.RequestOfferSent
                Call WriteMercado(Userindex, MercadoPacketID.SendOfferSent)
                
            Case MercadoPacketID.RequestTipoMAO
                Call WriteSendTipoMAO(Userindex, MercadoPacketID.SendTipoMAO, buffer.ReadASCIIString)
                
            Case MercadoPacketID.RequestInfoCharMAO
                Call SendInfoCharMAO(Userindex, buffer.ReadASCIIString)
                
            Case MercadoPacketID.PublicationPj
                Tipo = buffer.ReadByte
                email = buffer.ReadASCIIString
                Passwd = buffer.ReadASCIIString
                Pin = buffer.ReadASCIIString
                
                If (Tipo = 1) Then ' Cambios de personaje
                
                    MAO.Add_Change Userindex, email, Passwd, Pin
                Else
                    MAO.Add_Gld_Dsp Userindex, email, Passwd, Pin, buffer.ReadASCIIString, buffer.ReadLong, buffer.ReadLong
                End If
                
            Case MercadoPacketID.InvitationChange
                If buffer.ReadASCIIString <> GetVar(CharPath & UCase$(.Name) & ".chr", "INIT", "PIN") Then
                    WriteConsoleMsg Userindex, "Pin incorrecto", FontTypeNames.FONTTYPE_INFO
                Else
                    MAO.Send_Invitation Userindex, buffer.ReadASCIIString
                End If
            
            Case MercadoPacketID.AcceptInvitation
                MAO.Accept_Invitation Userindex, buffer.ReadASCIIString, buffer.ReadASCIIString
                
            Case MercadoPacketID.RechaceInvitation
                MAO.Rechace_Invitation Userindex, buffer.ReadASCIIString
                
            Case MercadoPacketID.CancelInvitation
                MAO.Cancel_Invitation Userindex, buffer.ReadASCIIString
                
            Case MercadoPacketID.BuyPj
                MAO.Buy_Pj Userindex, buffer.ReadASCIIString
                
            Case MercadoPacketID.QuitarPj
                MAO.Remove_Pj Userindex
        End Select
    
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
    
End Sub

Private Sub WriteSendTipoMAO(ByVal Userindex As Integer, ByVal PacketID As MercadoPacketID, ByVal UserName As String)
    
    On Error GoTo Errhandler
        
        'Dim InList As Boolean
        Dim InList As Integer
        Dim strtemp As String
        
        InList = val(GetVar(App.Path & "\CHARFILE\" & UserName & ".CHR", "MERCADO", "InList"))
        
        If Not InList > 0 Then
            Call LogMao("El personaje " & UserList(Userindex).Name & " ha intentado comprar un PJ en MAO que no está registrado. Nick: " & UserName)
            Exit Sub
        End If
        
        If GetVar(App.Path & "\CHARFILE\" & UserName & ".CHR", "MERCADO", "Change") Then
            strtemp = "X CAMBIO"
        ElseIf val(GetVar(App.Path & "\CHARFILE\" & UserName & ".CHR", "MERCADO", "Dsp")) > 0 Or val(GetVar(App.Path & "\CHARFILE\" & UserName & ".CHR", "MERCADO", "Gld")) > 0 Then
            strtemp = "ORO: " & val(GetVar(App.Path & "\CHARFILE\" & UserName & ".CHR", "MERCADO", "Gld")) & ", DSP: " & val(GetVar(App.Path & "\CHARFILE\" & UserName & ".CHR", "MERCADO", "Dsp"))
        End If
        
        Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.SvMercado)
        Call UserList(Userindex).outgoingData.WriteByte(PacketID)
        Call UserList(Userindex).outgoingData.WriteASCIIString(strtemp)
        
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Private Sub WriteMercado(ByVal Userindex As Integer, ByVal PacketID As MercadoPacketID)

    On Error GoTo Errhandler
        Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.SvMercado)
        Call UserList(Userindex).outgoingData.WriteByte(PacketID)
    
        Select Case PacketID
            Case MercadoPacketID.SendMercado
                Call UserList(Userindex).outgoingData.WriteASCIIString(Chars_Mercado)
                
            Case MercadoPacketID.SendOffer
                Call UserList(Userindex).outgoingData.WriteASCIIString(Char_Offer(Userindex))
                
            Case MercadoPacketID.SendOfferSent
                Call UserList(Userindex).outgoingData.WriteASCIIString(Char_OfferSent(Userindex))
                
        End Select
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
Private Sub WriteFormRostro(ByVal Userindex As Integer)
    
    On Error GoTo Errhandler
        Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.RequestFormRostro)
        
        Call UserList(Userindex).outgoingData.WriteByte(UserList(Userindex).Genero)
        Call UserList(Userindex).outgoingData.WriteByte(UserList(Userindex).raza)
        
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Handles the "RightClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRightClick(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 10/05/2011
'
'***************************************************
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        
        X = .ReadByte()
        Y = .ReadByte()
        
        If Not CheckCRC(Userindex, .ReadInteger - 42) Then Exit Sub
        
        Call Extra.ShowMenu(Userindex, UserList(Userindex).Pos.map, X, Y)
    End With
End Sub

''
' Writes the "ShowMenu" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    MenuIndex: The menu index.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMenu(ByVal Userindex As Integer, ByVal MenuIndex As Byte)
'***************************************************
'Author: ZaMa
'Last Modification: 10/05/2011
'Writes the "ShowMenu" message to the given user's outgoing data buffer
'***************************************************
Dim i As Long

On Error GoTo Errhandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMenu)
        
        Call .WriteByte(MenuIndex)
        
        Select Case MenuIndex
           Case eMenues.ieUser
                Dim tUser As Integer
                Dim guild As String
                tUser = UserList(Userindex).flags.TargetUser
                
                If UserList(tUser).GuildIndex <> 0 Then
                    guild = guilds(UserList(tUser).GuildIndex).GuildName
                End If
                
                Call .WriteASCIIString(UCase$(UserList(tUser).Name) & "-" & UCase$(guild))
                
                Call .WriteByte(UserList(tUser).clase)
                Call .WriteByte(UserList(tUser).raza)
                Call .WriteByte(UserList(tUser).Stats.ELV)
                
                For i = 0 To MAX_LOGROS
                    Call .WriteByte(UserList(tUser).Logros(i))
                Next i
                'Rankings check
                'Call .WriteByte(EstaRanking(tUser, rFrags))
                'Call .WriteByte(EstaRanking(tUser, rCanjes))
                'Call .WriteByte(EstaRanking(tUser, rOro))
                'Call .WriteByte(EstaRanking(tUser, rEventos))
                
                'Logros
            Case eMenues.ieNpcComercio
            
            Case eMenues.ieNpcNoHostil
        End Select
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub


Public Sub HandleEventPacket(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte
        
        Dim PacketID As Integer
        Dim LoopC As Integer
        Dim Modality As eModalityEvent, Quotas As Byte, TeamCant As Byte, LvlMin As Byte, LvlMax As Byte, GldInscription As Long, DspInscription As Long, TimeInit As Long, TimeCancel As Long, AllowedClasses(1 To NUMCLASES) As Byte
        
        PacketID = buffer.ReadByte
        
        Select Case PacketID
            Case EventPacketID.eNewEvent
                Modality = buffer.ReadByte()
                Quotas = buffer.ReadByte()
                LvlMin = buffer.ReadByte()
                LvlMax = buffer.ReadByte()
                GldInscription = buffer.ReadLong()
                DspInscription = buffer.ReadLong()
                TimeInit = buffer.ReadLong()
                TimeCancel = buffer.ReadLong()
                TeamCant = buffer.ReadByte()
                
                For LoopC = 1 To NUMCLASES
                    AllowedClasses(LoopC) = buffer.ReadByte()
                Next LoopC
                
                If UCase$(.Name) = "THYRAH" Or UCase$(.Name) = "LAUTARO" Then
                    EventosDS.NewEvent Userindex, Modality, Quotas, LvlMin, LvlMax, GldInscription, DspInscription, TimeInit, TimeCancel, TeamCant, AllowedClasses()
                Else
                    SendData SendTarget.ToGM, 0, PrepareMessageConsoleMsg("El personaje " & .Name & " ha intentado crear un evento. Baneenlo debe ser CUICUI.", FontTypeNames.FONTTYPE_ADMIN)
                End If
            Case EventPacketID.eCloseEvent
                EventosDS.CloseEvent buffer.ReadByte, , True
                
            Case EventPacketID.RequiredEvents
                WriteEventPacket Userindex, SvEventPacketID.SendListEvent
                
            Case EventPacketID.RequiredDataEvent
                WriteEventPacket Userindex, SvEventPacketID.SendDataEvent, CByte(buffer.ReadByte())
            
            Case EventPacketID.eAbandonateEvent
                EventosDS.AbandonateEvent Userindex, , True
                
            Case EventPacketID.eParticipeEvent
                EventosDS.ParticipeEvent Userindex, buffer.ReadASCIIString
        End Select
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub


Public Sub WriteEventPacket(ByVal Userindex As Integer, ByVal PacketID As Byte, Optional ByVal DataExtra As Long)
    On Error GoTo Errhandler

    Dim LoopC As Integer
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.EventPacketSv)
        Call .WriteByte(PacketID)
        
        Select Case PacketID
            Case SvEventPacketID.SendListEvent
                For LoopC = 1 To EventosDS.MAX_EVENT_SIMULTANEO
                    Call .WriteByte(IIf((Events(LoopC).Enabled = True), Events(LoopC).Modality, 0))
                    
                Next LoopC
                
            Case SvEventPacketID.SendDataEvent
                Call .WriteByte(Events(DataExtra).Inscribed)
                Call .WriteByte(Events(DataExtra).Quotas)
                Call .WriteByte(Events(DataExtra).LvlMin)
                Call .WriteByte(Events(DataExtra).LvlMax)
                Call .WriteLong(Events(DataExtra).GldInscription * Events(DataExtra).Inscribed)
                Call .WriteLong(Events(DataExtra).DspInscription * Events(DataExtra).Inscribed)
                Call .WriteASCIIString(strUsersEvent(DataExtra))
        End Select
        
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
Private Sub HandlePaqueteEncriptado(ByVal Userindex As Integer)
'***************************************************
'Author: Damián
'Last Modification: 24/08/2013
'***************************************************
    With UserList(Userindex).incomingData
        Call .ReadByte 'Saco el paquete
        UserList(Userindex).CRC = .ReadInteger
        UserList(Userindex).CRC = UserList(Userindex).CRC - 42
        
        Dim DatosEncript As String
        DatosEncript = .ReadASCIIStringFixed(.length)
        
        DatosEncript = ConvertirFlush(DatosEncript, Userindex)
        .WriteASCIIStringFixed (DatosEncript)
        Call HandleIncomingData(Userindex)
    End With
End Sub

Public Sub WriteUserInEvent(ByVal Userindex As Integer)
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.UserInEvent)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Private Sub HandleCofres(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        
        Dim TipoCofre As Byte
        Dim Obj As Obj
        Dim strtemp As String
        TipoCofre = .incomingData.ReadByte
        
        Obj.Amount = 1
        
        'Config Basica
        Select Case TipoCofre
            Case 0 ' BRONCE
                Obj.objindex = 946
                
                If .flags.Bronce = 1 Then
                    WriteConsoleMsg Userindex, "Ya eres usuario BRONCE", FontTypeNames.FONTTYPE_GUILD
                    Exit Sub
                End If
                
                If Not TieneObjetos(Obj.objindex, Obj.Amount, Userindex) Then
                    WriteConsoleMsg Userindex, "Necesitas tener contigo el cofre de los inmortales [BRONCE]", FontTypeNames.FONTTYPE_WARNING
                    Exit Sub
                End If
                
                .flags.Bronce = 1
            Case 1 'PLATA
                Obj.objindex = 945
                
                If .flags.Plata = 1 Then
                    WriteConsoleMsg Userindex, "Ya eres usuario PLATA", FontTypeNames.FONTTYPE_GUILD
                    Exit Sub
                End If
                
                If Not TieneObjetos(Obj.objindex, Obj.Amount, Userindex) Then
                    WriteConsoleMsg Userindex, "Necesitas tener contigo el cofre de los inmortales [PLATA]", FontTypeNames.FONTTYPE_WARNING
                    Exit Sub
                End If
                
                .flags.Plata = 1
            Case 2 'ORO
                Obj.objindex = 944
                
                If .flags.Oro = 1 Then
                    WriteConsoleMsg Userindex, "Ya eres usuario ORO", FontTypeNames.FONTTYPE_GUILD
                    Exit Sub
                End If
                
                If Not TieneObjetos(Obj.objindex, Obj.Amount, Userindex) Then
                    WriteConsoleMsg Userindex, "Necesitas tener contigo el cofre de los inmortales [ORO]", FontTypeNames.FONTTYPE_WARNING
                    Exit Sub
                End If
                
                .flags.Oro = 1
            Case 3 'PREMIUM
                Obj.objindex = 1115
                
                If .flags.Premium = 1 Then
                    WriteConsoleMsg Userindex, "Ya eres usuario PREMIUM", FontTypeNames.FONTTYPE_GUILD
                    Exit Sub
                End If
                
                If Not TieneObjetos(Obj.objindex, Obj.Amount, Userindex) Then
                    WriteConsoleMsg Userindex, "Necesitas tener contigo el cofre de los inmortales [PREMIUM]", FontTypeNames.FONTTYPE_WARNING
                    Exit Sub
                End If
                
                .flags.Premium = 1
                
            Case 4
                'DIOS
                Obj.objindex = 1115
                
                If .flags.Premium = 1 Then
                    WriteConsoleMsg Userindex, "Ya eres usuario DIOS", FontTypeNames.FONTTYPE_GUILD
                    Exit Sub
                End If
                
                If Not TieneObjetos(Obj.objindex, Obj.Amount, Userindex) Then
                    WriteConsoleMsg Userindex, "Necesitas tener contigo el cofre de los inmortales [DIOS]", FontTypeNames.FONTTYPE_WARNING
                    Exit Sub
                End If
                
                .flags.Premium = 1
        End Select
        
        
    End With
End Sub

Private Sub HandleComandoPorDias(ByVal Userindex As Integer)
'***************************************************
'Author: Lautaro
'Last Modification: -
'
'***************************************************
    If UserList(Userindex).incomingData.length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim Tipo As Byte
        Dim tIndex As Integer
        Dim strDate As String
        
        Tipo = buffer.ReadByte()    'Que Larry?
        UserName = buffer.ReadASCIIString()    'Que UserName?
        strDate = buffer.ReadASCIIString()
        
        
        If StrComp(UCase$(.Name), "THYRAH") = 0 Then
            Select Case Tipo
                Case 0 ' Ban por días
                    If Not FileExist(App.Path & "\CHARFILE\" & UserName & ".CHR", vbNormal) Then
                        WriteConsoleMsg Userindex, "El personaje no existe.", FontTypeNames.FONTTYPE_INFO
                    Else
                        mDias.BanUserDias Userindex, UserName, strDate
                    End If
                Case 1 ' Convertir en dioses
                    tIndex = NameIndex(UserName)  'Que user index?
                    If tIndex > 0 Then
                        mDias.TransformarUserDios Userindex, tIndex, strDate
                    Else
                        WriteConsoleMsg Userindex, "El personaje está offline.", FontTypeNames.FONTTYPE_INFO
                    End If
                        
                End Select
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

Private Sub HandleReportCheat(ByVal Userindex As Integer)
'***************************************************
'Author: Lautaro
'Last Modification: -
'
'***************************************************
    If UserList(Userindex).incomingData.length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim DataName As String
        
        UserName = buffer.ReadASCIIString()
        DataName = buffer.ReadASCIIString()
        
        ' ME OLVIDE COMO SE ESCRIBE JAJAJAJAJAJA TOY PEOR Q VOS
        'APARENTEMENTE
        
        SendData SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El personaje " & UserName & " tiene un programa APARENTEMENTE peligroso abierto: " & DataName, FontTypeNames.FONTTYPE_ADMIN)
        LogAntiCheat "Se ha detectado que el personaje " & UserName & " tiene un posible programa prohibido llamado " & DataName
        

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

Private Sub HandleDisolutionGuild(ByVal Userindex As Integer)

    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Tipo As Byte
        
        Tipo = buffer.ReadByte()
        
        Select Case Tipo
            Case 0
                modGuilds.DisolverGuildIndex Userindex
            Case 1
                modGuilds.ReanudarGuildIndex Userindex, buffer.ReadASCIIString
        End Select

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

Public Sub WriteShortMsj(ByVal UIndex As Integer, _
                         ByVal MsgShort As Integer, _
                         ByVal FontType As FontTypeNames, _
                         Optional ByVal tmpInteger1 As Integer = 0, _
                         Optional ByVal tmpInteger2 As Integer = 0, _
                         Optional ByVal tmpInteger3 As Integer = 0, _
                         Optional ByVal tmpLong As Long = 0, _
                         Optional ByVal TmpStr As String = vbNullString)
                         
    With UserList(UIndex).outgoingData
        .WriteByte ServerPacketID.ShortMsj
        .WriteInteger MsgShort
        .WriteByte FontType

        If tmpInteger1 <> 0 Then .WriteInteger tmpInteger1
        If tmpInteger2 <> 0 Then .WriteInteger tmpInteger2
        If tmpInteger3 <> 0 Then .WriteInteger tmpInteger3
        If tmpLong <> 0 Then .WriteLong tmpLong
        If Len(TmpStr) <> 0 Then .WriteASCIIString TmpStr

    End With

End Sub

Public Function PrepareMessageShortMsj(ByVal MsgShort As Integer, _
                                       ByVal FontType As FontTypeNames, _
                                       Optional ByVal tmpInteger1 As Integer = 0, _
                                       Optional ByVal tmpInteger2 As Integer = 0, _
                                       Optional ByVal tmpInteger3 As Integer = 0, _
                                       Optional ByVal tmpLong As Long = 0, _
                                       Optional ByVal TmpStr As String = vbNullString) As String

    With auxiliarBuffer
        .WriteByte ServerPacketID.ShortMsj
        .WriteInteger MsgShort
        .WriteByte FontType
 
        If tmpInteger1 <> 0 Then .WriteInteger tmpInteger1
        If tmpInteger2 <> 0 Then .WriteInteger tmpInteger2
        If tmpInteger3 <> 0 Then .WriteInteger tmpInteger3
        If tmpLong <> 0 Then .WriteLong tmpLong
        If Len(TmpStr) <> 0 Then .WriteASCIIString TmpStr
        
        PrepareMessageShortMsj = .ReadASCIIStringFixed(.length)

    End With

End Function


Public Function PrepareMessagePalabrasMagicas(ByVal CharIndex As Integer, ByVal SpellIndex As Byte, ByVal color As Long) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PalabrasMagicas)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(SpellIndex)

        ' Write rgb channels and save one byte from long :D
        Call .WriteByte(color And &HFF)
        Call .WriteByte((color And &HFF00&) \ &H100&)
        Call .WriteByte((color And &HFF0000) \ &H10000)

        PrepareMessagePalabrasMagicas = .ReadASCIIStringFixed(.length)
    End With
End Function
Public Function PrepareMessageDescNpcs(ByVal CharIndex As Integer, ByVal NumeroNpc As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.DescNpcs)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(NumeroNpc)

        PrepareMessageDescNpcs = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Sub WriteDescNpcs(ByVal Userindex As Integer, ByVal CharIndex As Integer, ByVal NumeroNpc As Integer)
    On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageDescNpcs(CharIndex, NumeroNpc))
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Private Sub HandleChangeNick(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String, TmpStr As String, p As String

        UserName = buffer.ReadASCIIString()
        
        Call General.ChangeNick(Userindex, UserName)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

Private Function SearchItemCanje(ByVal CanjeItem As Integer, ByVal ObjRequired1 As Integer, ByVal Points As Integer) As Byte

    Dim LoopC As Integer
    Dim ArrayValue As Long
    
    For LoopC = 1 To NumCanjes
        With Canjes(LoopC)
            If CanjeItem = .ObjCanje.objindex Then
                GetSafeArrayPointer Canjes(LoopC).ObjRequired, ArrayValue
                
                If ArrayValue <> 0 Then
                    If Canjes(LoopC).ObjRequired(1).objindex = ObjRequired1 And Points = Canjes(LoopC).Points Then
                        SearchItemCanje = LoopC
                        Exit For
                    End If
                Else
                    If Canjes(LoopC).Points = Points Then
                        SearchItemCanje = LoopC
                        Exit For
                    End If
                End If
            End If
        End With
    Next LoopC
End Function

Public Sub HandleCanjeItem(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        
        If .flags.Muerto Then Exit Sub
        
        Dim CanjeItem As Integer
        Dim CanjeIndex As Byte
        
        CanjeItem = .incomingData.ReadInteger
        CanjeIndex = SearchItemCanje(CanjeItem, .incomingData.ReadInteger, .incomingData.ReadInteger)
        
        If CanjeIndex = 0 Then Exit Sub
        
        Call General.CanjearObjeto(Userindex, CanjeIndex)
        
        
    End With
End Sub

Public Sub HandleCanjeInfo(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte

        Dim CanjeItem As Integer
        Dim CanjeIndex As Byte
        
        CanjeItem = .incomingData.ReadInteger
        CanjeIndex = SearchItemCanje(CanjeItem, .incomingData.ReadInteger, .incomingData.ReadInteger)
        
        If CanjeIndex = 0 Then Exit Sub
        If .flags.Muerto Then Exit Sub
        
        Call WriteCanjeInfo(Userindex, CanjeIndex)
    
    End With
End Sub
Public Sub WriteCanjeInfo(ByVal Userindex As Integer, ByVal CanjeIndex As Byte)
    On Error GoTo Errhandler
    Dim LoopC As Integer
    Dim LoopY As Integer
        
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.InfoCanje)
    
    
    With ObjData(Canjes(CanjeIndex).ObjCanje.objindex)
        Call UserList(Userindex).outgoingData.WriteInteger(.MinDef)
        Call UserList(Userindex).outgoingData.WriteInteger(.MaxDef)
        Call UserList(Userindex).outgoingData.WriteInteger(.DefensaMagicaMin)
        Call UserList(Userindex).outgoingData.WriteInteger(.DefensaMagicaMax)
        Call UserList(Userindex).outgoingData.WriteInteger(.MinHIT)
        Call UserList(Userindex).outgoingData.WriteInteger(.MaxHIT)
        Call UserList(Userindex).outgoingData.WriteLong(Canjes(CanjeIndex).Points)
        
        If .OBJType = otMonturas Or .OBJType = otMonturasDraco Then
            Call UserList(Userindex).outgoingData.WriteByte(1)
        Else
            Call UserList(Userindex).outgoingData.WriteByte(.NoSeCae)
        End If
        
        
    End With
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Private Function SearchNpcCanje(ByVal CanjeIndex As Integer, ByVal NpcNumero As Integer) As Boolean
    Dim LoopC As Integer
    
    
    SearchNpcCanje = False
    
    With Canjes(CanjeIndex)
        If .NPCs = NpcNumero Then
            SearchNpcCanje = True
            Exit Function
        End If
    
    End With
End Function
Public Sub WriteCanjeInit(ByVal Userindex As Integer, ByVal NpcNumero As Integer)
    On Error GoTo Errhandler
        Dim LoopC As Integer
        Dim LoopY As Integer
        Dim NpcIndex As Integer
        Dim SearchNpc As Boolean
        Dim Num As Integer
        
        
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.CanjeInit)
    
    For LoopC = 1 To NumCanjes
        SearchNpc = SearchNpcCanje(LoopC, NpcNumero)
        
        If SearchNpc Then
            Num = Num + 1
        End If
    Next LoopC

    
    Call UserList(Userindex).outgoingData.WriteByte(Num)
    
    If Num = 0 Then Exit Sub
    
    For LoopC = 1 To NumCanjes
        With Canjes(LoopC)
            If SearchNpcCanje(LoopC, NpcNumero) Then
            
                Call UserList(Userindex).outgoingData.WriteByte(.NumRequired)
                
                For LoopY = 1 To .NumRequired
                     Call UserList(Userindex).outgoingData.WriteInteger(.ObjRequired(LoopY).objindex)
                     Call UserList(Userindex).outgoingData.WriteInteger(.ObjRequired(LoopY).Amount)
                Next LoopY
                
                Call UserList(Userindex).outgoingData.WriteInteger(.ObjCanje.objindex)
                Call UserList(Userindex).outgoingData.WriteInteger(.ObjCanje.Amount)
                Call UserList(Userindex).outgoingData.WriteInteger(ObjData(.ObjCanje.objindex).GrhIndex)
                Call UserList(Userindex).outgoingData.WriteInteger(.Points)
            End If
        End With
        
    Next LoopC
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub HandlePacketRetos(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim UserName As String, TmpStr As String, p As String
        Dim Tipo As Byte, SubTipo As Byte
        Dim GldRequired As Long, DspRequired As Long, LimiteRojas As Integer
        Dim Users() As String, Team(10) As Byte
        Dim LoopC As Integer
        
        Tipo = buffer.ReadByte
        
        Select Case Tipo
            Case 0 ' Enviar solicitud
                GldRequired = buffer.ReadLong
                DspRequired = buffer.ReadLong
                LimiteRojas = buffer.ReadInteger
                UserName = buffer.ReadASCIIString
                UserName = UserName & "-" & .Name
                
                Users = Split(UserName, "-")
                
                If Not RetosActivos Then
                    WriteConsoleMsg Userindex, "Los retos están desactivados.", FontTypeNames.FONTTYPE_INFO
                Else
                    Call mRetos.SendFight(Userindex, eTipoReto.FightOne, GldRequired, DspRequired, LimiteRojas, Users)
                End If
            Case 1 ' Aceptar solicitud
                UserName = buffer.ReadASCIIString
                
                If Not RetosActivos Then
                    WriteConsoleMsg Userindex, "Los retos están desactivados.", FontTypeNames.FONTTYPE_INFO
                Else
                    Call mRetos.AcceptFight(Userindex, UserName)
                End If
                
            Case 2 ' Salir del evento
                If .flags.SlotReto > 0 Then
                    Call mRetos.UserdieFight(Userindex, 0, True)
                End If
                
            Case 3 'Enviar Clan vs Clan
                UserName = buffer.ReadASCIIString
                
                mCVC.SendFightGuild Userindex, NameIndex(UserName)
                
            Case 4 'Aceptar Clan vs Clan
                UserName = buffer.ReadASCIIString
                mCVC.AcceptFightGuild Userindex, NameIndex(UserName)
            
            Case 5 'Requerimos el panel de retos
                WriteSendRetos Userindex
                
        End Select
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub
Public Sub WriteSendRetos(ByVal Userindex As Integer)
    Dim strtemp As String
    
    On Error GoTo Errhandler
    
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.SendRetos)
        strtemp = Ranking(eRanking.TopRetos).Nombre(1) & "-" & Ranking(eRanking.TopRetos).Nombre(2) & "-" & Ranking(eRanking.TopRetos).Nombre(3)
    
        Call UserList(Userindex).outgoingData.WriteASCIIString(strtemp)
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Private Function strListApostadores() As String
    Dim LoopC As Integer
    
    For LoopC = LBound(GambleSystem.Users()) To UBound(GambleSystem.Users())
        If GambleSystem.Users(LoopC).Name <> vbNullString Then
            strListApostadores = strListApostadores & GambleSystem.Users(LoopC).Name & "-"
        End If
    
    Next LoopC
    
    
    If Len(strListApostadores) > 0 Then
        strListApostadores = mid$(strListApostadores, 1, Len(strListApostadores) - 1)
    End If
End Function

Private Function strListApuestas() As String
    Dim LoopC As Integer
    
    For LoopC = LBound(GambleSystem.Apuestas()) To UBound(GambleSystem.Apuestas())
        If GambleSystem.Apuestas(LoopC) <> vbNullString Then
            strListApuestas = strListApuestas & GambleSystem.Apuestas(LoopC) & ","
        End If
    
    Next LoopC
    
    If Len(strListApuestas) > 0 Then
        strListApuestas = mid$(strListApuestas, 1, Len(strListApuestas) - 1)
    End If
End Function

Public Sub WritePacketGambleSv(ByVal Userindex As Integer, ByVal Tipo As Byte)
    Dim strtemp As String
    
    On Error GoTo Errhandler
    
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.PacketGambleSv)
    Call UserList(Userindex).outgoingData.WriteByte(Tipo)
    
    Select Case Tipo
        Case 0 'Enviamos la lista de usuarios que apostaron
            Call UserList(Userindex).outgoingData.WriteASCIIString(strListApostadores)
        Case 1 ' Enviamos la info de los usuarios que apostaron
            
        Case 2 ' Enviamos la lista de apuestas disponibles para los usuarios
            Call UserList(Userindex).outgoingData.WriteASCIIString(strListApuestas)
    End Select
    
    Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub HandleUseItemPacket(ByVal Userindex As Integer)
    ' Paquete mejorado por LAUTARO.
    ' Intento de anti DLLS.
    
    Dim Tipo As Byte
    Dim Slot As Byte
    Dim SecondaryClick As Byte
    
    With UserList(Userindex)
        Call .incomingData.ReadByte
        
        Slot = .incomingData.ReadByte
        Tipo = .incomingData.ReadByte
        SecondaryClick = .incomingData.ReadByte
        .incomingData.ReadBoolean
        
        If Tipo <> .KeyUseItem Then
            LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
        End If
        
        Select Case Tipo
            Case 1
                If .incomingData.ReadByte <> 200 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
            Case 2
                If .incomingData.ReadByte <> 155 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
                
                If .incomingData.ReadByte <> 10 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
            Case 3
                If .incomingData.ReadInteger <> 15785 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
            Case 4
                If .incomingData.ReadInteger <> 12148 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
                
                If .incomingData.ReadByte <> 245 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
            Case 5
                If .incomingData.ReadInteger <> 1548 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
                
                If .incomingData.ReadInteger <> 15 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
            Case 6
                If .incomingData.ReadByte <> 255 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
                
                If .incomingData.ReadByte <> 200 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
                
                If .incomingData.ReadByte <> 100 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
            Case 7
                If .incomingData.ReadByte <> 154 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
                
                If .incomingData.ReadByte <> 104 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If

                If .incomingData.ReadByte <> 111 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If

                If .incomingData.ReadByte <> 84 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
            Case 8
                If .incomingData.ReadByte <> 10 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
                
                If .incomingData.ReadInteger <> 5457 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
            Case 9
                If .incomingData.ReadBoolean <> False Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
                
                If .incomingData.ReadByte <> 45 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
            Case 10
                If .incomingData.ReadBoolean <> False Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
            Case 11
                If .incomingData.ReadLong <> 545646581 Then
                    LogAntiCheat "El personaje " & .Name & " puede estar con una DLL."
                End If
            Case 12
                
        End Select
        
        Call UsarItem(Userindex, Slot, SecondaryClick)
        
    End With
End Sub

Public Sub UsarItem(ByVal Userindex As Integer, ByVal Slot As Byte, ByVal SecondaryClick As Byte)
    With UserList(Userindex)
        If .flags.LastSlotClient <> 255 Then
            If Slot <> .flags.LastSlotClient Then
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("ANTICHEAT > VIGILAR ACTITUD MUY SOSPECHOSA a " & .Name & " Informacion confidencial. ", FontTypeNames.FONTTYPE_EJECUCION))
                Call LogAntiCheat(.Name & " Cambio de slot estando en la ventana de hechizos.")
                Exit Sub
            End If
        End If

        If Slot <= .CurrentInventorySlots And Slot > 0 Then
            If .Invent.Object(Slot).objindex = 0 Then Exit Sub
        End If

        If .flags.Meditando Then
            Exit Sub    'The error message should have been provided by the client.
        End If
        
        If ObjData(.Invent.Object(Slot).objindex).OBJType = otPociones Then
            Call UseInvPotion(Userindex, Slot, SecondaryClick)
        Else
            Call UseInvItem(Userindex, Slot)
        End If
        
        Call WriteUpdateFollow(Userindex)
    End With
End Sub

Private Sub HandleDarPoints(ByVal Userindex As Integer)
'***************************************************
'Author: Lautaro
'Last Modification: -
'
'***************************************************
    If UserList(Userindex).incomingData.length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim Amount As Integer
        Dim tUser As Integer
        
        Amount = buffer.ReadInteger()    'Que Larry?
        UserName = buffer.ReadASCIIString()    'Que UserName?
        
        
        If StrComp(UCase$(.Name), "LAUTARO") = 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                UserList(tUser).Stats.TorneosGanados = UserList(tUser).Stats.TorneosGanados + Amount
                WriteConsoleMsg Userindex, "Le has dado " & Amount & " puntos de torneo/canje a " & UserName & ".", FontTypeNames.FONTTYPE_INFO
                WriteConsoleMsg tUser, "Has recibido " & Amount & " puntos de torneo/canje.", FontTypeNames.FONTTYPE_INFO
                CheckRankingUser tUser, TopTorneos
            Else
                WriteConsoleMsg Userindex, "Personaje offline.", FontTypeNames.FONTTYPE_INFO
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub

Public Sub HandleRequestInfoEvento(ByVal Userindex As Integer)

    Dim strtemp As String
    
    With UserList(Userindex)
        Call .incomingData.ReadByte
        
        strtemp = SetInfoEvento
        
        If strtemp = vbNullString Then
            WriteConsoleMsg Userindex, "No hay eventos en curso. Danos tu sugerencia mediante /DENUNCIAR.", FontTypeNames.FONTTYPE_INFO
        Else
            WriteConsoleMsg Userindex, SetInfoEvento, FontTypeNames.FONTTYPE_INFO
        End If
    
    End With
End Sub

Private Sub HandlePacketGamble(ByVal Userindex As Integer)
'***************************************************
'Author: Lautaro
'Last Modification: -
'
'***************************************************
    If UserList(Userindex).incomingData.length < 1 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim Tipo As Byte
        Dim tUser As Integer
        Dim Apuestas() As String
        
        Tipo = buffer.ReadByte()
        
        Select Case Tipo
            Case 0 ' Gm crea nueva apuesta
                Apuestas = Split(buffer.ReadASCIIString, ",")
                mApuestas.NewGamble Userindex, buffer.ReadASCIIString, buffer.ReadInteger, buffer.ReadByte, Apuestas
            Case 1 ' Gm cancela la apuesta
                mApuestas.CancelGamble Userindex
            Case 2 ' Gm otorga premio de la apuesta
                
            Case 3 ' Personaje apuesta
                mApuestas.UserGamble Userindex, buffer.ReadByte, buffer.ReadLong, buffer.ReadLong
                
            Case 4 ' Gm requiere la lista de usuarios apostando
                WritePacketGambleSv Userindex, 0
            Case 5 ' Info de los users de arriba
                buffer.ReadASCIIString
                
                WritePacketGambleSv Userindex, 1
            Case 6 ' Lista de apuestas disponibles
                If GambleSystem.Run Then
                    WritePacketGambleSv Userindex, 2
                End If
                
            Case 7
                If EsGM(Userindex) Then
                    UserGambleWin Userindex, buffer.ReadASCIIString
                End If
                
        End Select

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
       Err.Raise error
End Sub
