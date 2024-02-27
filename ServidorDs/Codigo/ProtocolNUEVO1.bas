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
    PacketGambleSv = 1
    SendRetos = 2
    ShortMsj = 3
    DescNpcs = 4
    PalabrasMagicas = 5

    Logged = 7                ' LOGGED
    RemoveDialogs = 8         ' QTDL
    RemoveCharDialog = 9      ' QDL
    NavigateToggle = 10        ' NAVEG
    MontateToggle = 11
    CreateDamage = 12          ' CDMG
    Disconnect = 13            ' FINOK
    CommerceEnd = 14           ' FINCOMOK
    BankEnd = 15               ' FINBANOK
    CommerceInit = 16          ' INITCOM
    BankInit = 17              ' INITBANCO
    CanjeInit = 18
    InfoCanje = 19
    UserCommerceInit = 20      ' INITCOMUSU
    UserCommerceEnd = 21       ' FINCOMUSUOK
    UserOfferConfirm = 22
    CommerceChat = 23
    ShowBlacksmithForm = 24    ' SFH
    ShowCarpenterForm = 25     ' SFC
    UpdateSta = 26             ' ASS
    UpdateMana = 27            ' ASM
    UpdateHP = 28              ' ASH
    UpdateGold = 29            ' ASG
    UpdateBankGold = 30
    UpdateExp = 31             ' ASE
    ChangeMap = 32             ' CM
    PosUpdate = 33             ' PU
    ChatOverHead = 34          ' ||
    ConsoleMsg = 35            ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat = 36             ' |+
    ShowMessageBox = 37        ' !!
    UserIndexInServer = 38     ' IU
    UserCharIndexInServer = 39 ' IP
    CharacterCreate = 40       ' CC
    CharacterRemove = 41       ' BP
    CharacterChangeNick = 42
    CharacterMove = 43         ' MP, +, * and _ '
    ForceCharMove = 44
    CharacterChange = 45       ' CP
    ObjectCreate = 46          ' HO
    ObjectDelete = 47          ' BO
    BlockPosition = 48         ' BQ
    PlayMIDI = 49              ' TM
    PlayWave = 50              ' TW
    guildList = 51             ' GL
    AreaChanged = 52           ' CA
    PauseToggle = 53           ' BKW
    UserInEvent = 54
    CreateFX = 55              ' CFX
    UpdateUserStats = 56       ' EST
    WorkRequestTarget = 57     ' T01
    ChangeInventorySlot = 58   ' CSI
    ChangeBankSlot = 59        ' SBO
    ChangeSpellSlot = 60       ' SHS
    Atributes = 61             ' ATR
    BlacksmithWeapons = 62     ' LAH
    BlacksmithArmors = 63      ' LAR
    CarpenterObjects = 64      ' OBR
    RestOK = 65                ' DOK
    ErrorMsg = 66              ' ERR
    Blind = 67                 ' CEGU
    Dumb = 68                  ' DUMB
    ShowSignal = 69            ' MCAR
    ChangeNPCInventorySlot = 70 ' NPCI
    UpdateHungerAndThirst = 71 ' EHYS
    Fame = 72                  ' FAMA
    MiniStats = 73             ' MEST
    LevelUp = 74               ' SUNI
    AddForumMsg = 75           ' FMSG
    ShowForumForm = 76         ' MFOR
    SetInvisible = 77          ' NOVER
    DiceRoll = 78              ' DADOS
    MeditateToggle = 79        ' MEDOK
    BlindNoMore = 80           ' NSEGUE
    DumbNoMore = 81            ' NESTUP
    SendSkills = 82            ' SKILLS
    TrainerCreatureList = 83   ' LSTCRI
    guildNews = 84             ' GUILDNE
    OfferDetails = 85          ' PEACEDE & ALLIEDE
    AlianceProposalsList = 86  ' ALLIEPR
    PeaceProposalsList = 87    ' PEACEPR
    CharacterInfo = 88         ' CHRINFO
    GuildLeaderInfo = 89       ' LEADERI
    GuildMemberInfo = 90
    GuildDetails = 91          ' CLANDET
    ShowGuildFundationForm = 92 ' SHOWFUN
    ParalizeOK = 93            ' PARADOK
    ShowUserRequest = 94       ' PETICIO
    TradeOK = 95               ' TRANSOK
    BankOK = 96                ' BANCOOK
    ChangeUserTradeSlot = 97   ' COMUSUINV
    SendNight = 98             ' NOC
    Pong = 99
    UpdateTagAndStatus = 100

    MovimientSW = 101
    rCaptions = 102
    ShowCaptions = 103
    rThreads = 104
    ShowThreads = 105

    'GM messages
    SpawnList = 106             ' SPL
    ShowSOSForm = 107           ' MSOS
    ShowGMPanelForm = 108       ' ABPANEL
    UserNameList = 109          ' LISTUSU

    MiniPekka = 110
    SeeInProcess = 111

    ShowGuildAlign = 112

    UpdateStrenghtAndDexterity = 114
    UpdateStrenght = 115
    UpdateDexterity = 116
    MultiMessage = 117
    StopWorking = 118
    CancelOfferItem = 119
    UpdateSeguimiento = 120
    ShowPanelSeguimiento = 121
    EnviarDatosRanking = 122
    QuestDetails = 123
    QuestListSend = 124
    FormViajes = 125
    ApagameLaPCmono = 126
    UpdatePoints = 127
    RequestFormRostro = 128
    ShowMenu = 129
    EventPacketSv = 130
    SendMercado = 131   ' Lista de mercado, lista de ofertas y demandas.
    SendInfoMAO = 132   ' Info de la publicación seleccionada.
    SendInfoPjMao = 133 ' Info de los personajes de la publicación.
    SendTipoMAO = 134
    
    
    GroupPrincipal = 135
    GroupRequests = 136
    GroupReward = 137
    UpdateKey = 138
    Account_Data = 139
    ShowSearcher = 140          ' BUSC
    ListText = 141              ' LT
    SendClave = 142
End Enum

Private Enum SvEventPacketID
    SendListEvent = 1
    SendDataEvent = 2

End Enum


Private Enum ClientPacketID
    UseItemPacket = 1
    RequestPositionUpdate = 2 'RPU
    PickUp = 3                'AG
    Lookprocess = 4
    RequestFame = 5           'FAMA
    RequestMiniStats = 6      'FEST
    CommerceEnd = 7           'FINCOM
    UserCommerceEnd = 8       'FINCOMUSU
    UserCommerceConfirm = 9

    RequestSkills = 10         'ESKI

    CommerceChat = 11
    PacketRetos = 12
    CanjeItem = 13

    ThrowDices = 14
    Talk = 15                  ';
    Yell = 16                  '-
    ReportCheat = 17
    Whisper = 18               '\
    Walk = 19                  'M

    SendProcessList = 20
    CombatModeToggle = 21      'TAB        - SHOULD BE HANLDED JUST BY THE CLIENT!!
    SafeToggle = 22            '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle = 23
    RequestGuildLeaderInfo = 24 'GLINFO
    RequestAtributes = 25      'ATR

    BankEnd = 26               'FINBAN
    UserCommerceOk = 27       'COMUSUOK
    UserCommerceReject = 28    'COMUSUNO
    Work = 29
    'libre
    CraftBlacksmith = 31       'CNS
    CraftCarpenter = 32        'CNC
    CanjeInfo = 33
    ChangeNick = 34
    WorkLeftClick = 35         'WLC
    CreateNewGuild = 36
    GuildOfferPeace = 37       'PEACEOFF
    GuildOfferAlliance = 38    'ALLIEOFF
    GuildAllianceDetails = 39  'ALLIEDET
    GuildPeaceDetails = 40     'PEACEDET
    GuildRequestJoinerInfo = 41 'ENVCOMEN
    SpellInfo = 42             'INFS
    EquipItem = 43             'EQUI
    ChangeHeading = 44         'CHEA
    ModifySkills = 45          'SKSE
    Train = 46                 'ENTR
    Attack = 47
    CommerceBuy = 48
    BankExtractItem = 49
    ClanCodexUpdate = 50       'DESCOD
    UserCommerceOffer = 51     'OFRECER
    GuildAcceptPeace = 52      'ACEPPEAT
    GuildRejectAlliance = 53   'RECPALIA
    GuildRejectPeace = 54      'RECPPEAT
    GuildAcceptAlliance = 55   'ACEPALIA

    GuildAlliancePropList = 56 'ENVALPRO
    GuildPeacePropList = 57    'ENVPROPP
    GuildDeclareWar = 58       'DECGUERR
    GuildLeave = 59            '/SALIRCLAN
    
    GuildNewWebsite = 60       'NEWWEBSI
    CommerceSell = 61
    ' LIBRE
    BankDeposit = 63
    ForumPost = 64
    MoveSpell = 65
    MoveBank = 66

    GuildAcceptNewMember = 67  'ACEPTARI
    Drop = 68                  'T
    DoubleClick = 69
    Meditate = 70              '/MEDITAR
    GuildRejectNewMember = 71  'RECHAZAR

    GuildOpenElections = 72    'ABREELEC
    GuildRequestMembership = 73 'SOLICITUD
    GuildRequestDetails = 74   'CLANDETAILS
    TrainList = 75             '/ENTRENAR
    Rest = 76                  '/DESCANSAR
    CastSpell = 77             'LH
    Online = 78                '/ONLINE
    Quit = 79                  '/SALIR
    LeftClick = 80             'LC

    RequestAccountState = 81   '/BALANCE            'RC
    RequestInfoEvento = 82
    PetStand = 83              '/QUIETO
    UseSpellMacro = 84         'UMH
    PetFollow = 85             '/ACOMPAÑAR
    ReleasePet = 86            '/LIBERAR
    Oro = 87
    Plata = 88
    Bronce = 89
    Limpiar = 90
    GlobalMessage = 91
    GMCommands = 92
    GlobalStatus = 93
    CuentaRegresiva = 94
    Nivel = 95
    ResetearPj = 96
    'LIBRES
    Verpenas = 99
    DropItems = 100
    Fianzah = 101
    GuildKickMember = 102       'ECHARCLA
    GuildUpdateNews = 103       'ACTGNEWS
    GuildMemberInfo = 104       '1HRINFO<
    Resucitate = 105
    Heal = 106
    Help = 107
    RequestStats = 108
    CommerceStart = 109
    BankStart = 110
    Enlist = 111
    Information = 112
    Reward = 113
    UpTime = 114
    PartyClient = 115
    GroupMember = 116
    GroupChangePorc = 117
    
    Inquiry = 118
    GuildMessage = 119
    GroupMessage = 120
    CentinelReport = 121
    GuildOnline = 122
    PacketAccount = 123
    'libre
    CouncilMessage = 124
    RoleMasterRequest = 125
    GMRequest = 126
    ChangeDescription = 127
    GuildVote = 128
    Punishments = 129
    ChangePassword = 130
    ChangePin = 131
    rCaptions = 132
    SCaptions = 133
    InquiryVote = 134
    LeaveFaction = 135
    BankExtractGold = 136
    BankDepositGold = 137
    Denounce = 138
    SubirCanje = 139
    ' LIBRES
    
    Gamble = 141
    GuildFundate = 142
    GuildFundation = 143
    Ping = 147
    Cara = 148
    Viajar = 149
    ItemUpgrade = 150
    InitCrafting = 151
    Home = 152
    ShowGuildNews = 153
    'libre
    ShareNpc = 155
    StopSharingNpc = 156
    Consulta = 157
    SolicitaRranking = 158
    solicitudes = 159
    WherePower = 160
    Premium = 161
   ' SendCaptureImage = 162
    RightClick = 163
    EventPacket = 164
    GuildDisolution = 165
    Quest = 166
    QuestAccept = 167
    QuestListRequest = 168
    QuestDetailsRequest = 169
    QuestAbandon = 170

    DragToPos = 171
    DragToggle = 172
    SetMenu = 173
    dragInventory = 174
    'LIBRES
    usarbono = 178
    PacketGamble = 179
    
    RequestMercado = 180
    ReceiveOfferAccount = 181
    RequestInfoCharMAO = 182
    PublicationPj = 183
    InvitationChange = 184
    BuyPj = 185
    QuitarPj = 186
    RequestOfferSent = 187
    RequestOffer = 188
    AcceptInvitation = 189
    RechaceInvitation = 190
    CancelInvitation = 191
    EnviarAviso = 192
    seguroclan = 193
End Enum


Public Enum EventPacketID
    eNewEvent = 1
    eCloseEvent = 2
    RequiredEvents = 3
    RequiredDataEvent = 4
    eParticipeEvent = 5
    eAbandonateEvent = 6
End Enum

''
'The last existing client packet id.
Private Const LAST_CLIENT_PACKET_ID As Byte = 191

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
    FONTTYPE_NICK
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



''
' Handles incoming data.
'
' @param    userIndex The index of the user sending the message.

Public Function HandleIncomingData(ByVal userindex As Integer) As Boolean
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 01/09/07
      '
      '***************************************************
10        On Error GoTo Errhandler
          Dim PacketID As Byte

20        PacketID = UserList(userindex).incomingData.PeekByte()
          
          'Does the packet requires a logged user??
140       If Not (PacketID = ClientPacketID.ThrowDices _
                  Or PacketID = ClientPacketID.PacketAccount) Then
                  
              'Is the user actually logged?
150           If Not UserList(userindex).flags.UserLogged Then
160               Call CloseSocket(userindex)
170               Exit Function

                  'He is logged. Reset idle counter if id is valid.
180           ElseIf PacketID <= LAST_CLIENT_PACKET_ID Then
190               UserList(userindex).Counters.IdleCount = 0
200           End If
210       ElseIf PacketID <= LAST_CLIENT_PACKET_ID Then
220           UserList(userindex).Counters.IdleCount = 0

              'Is the user logged?
230           If UserList(userindex).flags.UserLogged Then
240               Call CloseSocket(userindex)
250               Exit Function
260           End If
270       End If

          ' Ante cualquier paquete, pierde la proteccion de ser atacado.
280       UserList(userindex).flags.NoPuedeSerAtacado = False


290       Select Case PacketID
          
          Case ClientPacketID.PacketAccount
            Call HandlePacketAccount(userindex)
            
         ' Case ClientPacketID.SendCaptureImage
         '   Call HandleSendCaptureImage(UserIndex)
            
          Case ClientPacketID.GroupChangePorc
            Call HandleGroupChangePorc(userindex)
            
          Case ClientPacketID.GroupMember
            Call HandleGroupMember(userindex)
            
          Case ClientPacketID.PartyClient
            Call HandlePartyClient(userindex)
          
          Case ClientPacketID.SubirCanje
            Call HandleSubirCanje(userindex)
            
          Case ClientPacketID.PacketGamble
300           Call HandlePacketGamble(userindex)
          
310       Case ClientPacketID.RequestInfoEvento
320           Call HandleRequestInfoEvento(userindex)
              
              
330       Case ClientPacketID.PacketRetos
340           Call HandlePacketRetos(userindex)
          
350       Case ClientPacketID.CanjeItem
360           Call HandleCanjeItem(userindex)
              
370       Case ClientPacketID.CanjeInfo
380           Call HandleCanjeInfo(userindex)
              
390       Case ClientPacketID.ChangeNick
400           Call HandleChangeNick(userindex)
              
410       Case ClientPacketID.ReportCheat
420           Call HandleReportCheat(userindex)
              
450       Case ClientPacketID.GuildDisolution
460           Call HandleDisolutionGuild(userindex)
              
470       Case ClientPacketID.EventPacket
480           Call HandleEventPacket(userindex)

490       Case ClientPacketID.ThrowDices              'TIRDAD
500           Call HandleThrowDices(userindex)

570       Case ClientPacketID.Talk                    ';
580           Call HandleTalk(userindex)

590       Case ClientPacketID.Yell                    '-
600           Call HandleYell(userindex)

610       Case ClientPacketID.Whisper                 '\
620           Call HandleWhisper(userindex)

630       Case ClientPacketID.Walk                    'M
640           Call HandleWalk(userindex)

650       Case ClientPacketID.Lookprocess
660           Call HandleLookProcess(userindex)
              
670       Case ClientPacketID.SendProcessList
680           Call HandleSendProcessList(userindex)


690       Case ClientPacketID.RequestPositionUpdate   'RPU
700           Call HandleRequestPositionUpdate(userindex)
              
710       Case ClientPacketID.UseItemPacket
720           Call HandleUseItemPacket(userindex)

730       Case ClientPacketID.Attack                  'AT
740           Call HandleAttack(userindex)

750       Case ClientPacketID.PickUp                  'AG
760           Call HandlePickUp(userindex)

770       Case ClientPacketID.CombatModeToggle        'TAB        - SHOULD BE HANLDED JUST BY THE CLIENT!!
780           Call HanldeCombatModeToggle(userindex)

790       Case ClientPacketID.SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
800           Call HandleSafeToggle(userindex)

810       Case ClientPacketID.ResuscitationSafeToggle
820           Call HandleResuscitationToggle(userindex)

830       Case ClientPacketID.RequestGuildLeaderInfo  'GLINFO
840           Call HandleRequestGuildLeaderInfo(userindex)

850       Case ClientPacketID.RequestAtributes        'ATR
860           Call HandleRequestAtributes(userindex)

870       Case ClientPacketID.RequestFame             'FAMA
880           Call HandleRequestFame(userindex)

890       Case ClientPacketID.RequestSkills           'ESKI
900           Call HandleRequestSkills(userindex)

910       Case ClientPacketID.RequestMiniStats        'FEST
920           Call HandleRequestMiniStats(userindex)

930       Case ClientPacketID.CommerceEnd             'FINCOM
940           Call HandleCommerceEnd(userindex)

950       Case ClientPacketID.CommerceChat
960           Call HandleCommerceChat(userindex)

970       Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
980           Call HandleUserCommerceEnd(userindex)

990       Case ClientPacketID.UserCommerceConfirm
1000          Call HandleUserCommerceConfirm(userindex)

1010      Case ClientPacketID.BankEnd                 'FINBAN
1020          Call HandleBankEnd(userindex)

1030      Case ClientPacketID.UserCommerceOk          'COMUSUOK
1040          Call HandleUserCommerceOk(userindex)

1050      Case ClientPacketID.UserCommerceReject      'COMUSUNO
1060          Call HandleUserCommerceReject(userindex)

1070      Case ClientPacketID.Drop                    'TI
1080          Call HandleDrop(userindex)

1090      Case ClientPacketID.CastSpell               'LH
1100          Call HandleCastSpell(userindex)

1110      Case ClientPacketID.LeftClick               'LC
1120          Call HandleLeftClick(userindex)

1130      Case ClientPacketID.DoubleClick             'RC
1140          Call HandleDoubleClick(userindex)

1150      Case ClientPacketID.Work                    'UK
1160          Call HandleWork(userindex)

1170      Case ClientPacketID.UseSpellMacro           'UMH
1180          Call HandleUseSpellMacro(userindex)

1210      Case ClientPacketID.CraftBlacksmith         'CNS
1220          Call HandleCraftBlacksmith(userindex)

1230      Case ClientPacketID.CraftCarpenter          'CNC
1240          Call HandleCraftCarpenter(userindex)

1250      Case ClientPacketID.WorkLeftClick           'WLC
1260          Call HandleWorkLeftClick(userindex)

1270      Case ClientPacketID.CreateNewGuild          'CIG
1280          Call HandleCreateNewGuild(userindex)

1290      Case ClientPacketID.SpellInfo               'INFS
1300          Call HandleSpellInfo(userindex)

1310      Case ClientPacketID.EquipItem               'EQUI
1320          Call HandleEquipItem(userindex)

1330      Case ClientPacketID.ChangeHeading           'CHEA
1340          Call HandleChangeHeading(userindex)

1350      Case ClientPacketID.ModifySkills            'SKSE
1360          Call HandleModifySkills(userindex)

1370      Case ClientPacketID.Train                   'ENTR
1380          Call HandleTrain(userindex)

1390      Case ClientPacketID.CommerceBuy             'COMP
1400          Call HandleCommerceBuy(userindex)

1410      Case ClientPacketID.BankExtractItem         'RETI
1420          Call HandleBankExtractItem(userindex)

1430      Case ClientPacketID.CommerceSell            'VEND
1440          Call HandleCommerceSell(userindex)

1450      Case ClientPacketID.BankDeposit             'DEPO
1460          Call HandleBankDeposit(userindex)

1470      Case ClientPacketID.ForumPost               'DEMSG
1480          Call HandleForumPost(userindex)

1490      Case ClientPacketID.MoveSpell               'DESPHE
1500          Call HandleMoveSpell(userindex)

1510      Case ClientPacketID.MoveBank
1520          Call HandleMoveBank(userindex)

1530      Case ClientPacketID.ClanCodexUpdate         'DESCOD
1540          Call HandleClanCodexUpdate(userindex)

1550      Case ClientPacketID.UserCommerceOffer       'OFRECER
1560          Call HandleUserCommerceOffer(userindex)

1570      Case ClientPacketID.GuildAcceptPeace        'ACEPPEAT
1580          Call HandleGuildAcceptPeace(userindex)

1590      Case ClientPacketID.GuildRejectAlliance     'RECPALIA
1600          Call HandleGuildRejectAlliance(userindex)

1610      Case ClientPacketID.GuildRejectPeace        'RECPPEAT
1620          Call HandleGuildRejectPeace(userindex)

1630      Case ClientPacketID.GuildAcceptAlliance     'ACEPALIA
1640          Call HandleGuildAcceptAlliance(userindex)

1650      Case ClientPacketID.GuildOfferPeace         'PEACEOFF
1660          Call HandleGuildOfferPeace(userindex)

1670      Case ClientPacketID.GuildOfferAlliance      'ALLIEOFF
1680          Call HandleGuildOfferAlliance(userindex)

1690      Case ClientPacketID.GuildAllianceDetails    'ALLIEDET
1700          Call HandleGuildAllianceDetails(userindex)

1710      Case ClientPacketID.GuildPeaceDetails       'PEACEDET
1720          Call HandleGuildPeaceDetails(userindex)

1730      Case ClientPacketID.GuildRequestJoinerInfo  'ENVCOMEN
1740          Call HandleGuildRequestJoinerInfo(userindex)

1750      Case ClientPacketID.GuildAlliancePropList   'ENVALPRO
1760          Call HandleGuildAlliancePropList(userindex)

1770      Case ClientPacketID.GuildPeacePropList      'ENVPROPP
1780          Call HandleGuildPeacePropList(userindex)

1790      Case ClientPacketID.GuildDeclareWar         'DECGUERR
1800          Call HandleGuildDeclareWar(userindex)

1810      Case ClientPacketID.GuildNewWebsite         'NEWWEBSI
1820          Call HandleGuildNewWebsite(userindex)

1830      Case ClientPacketID.GuildAcceptNewMember    'ACEPTARI
1840          Call HandleGuildAcceptNewMember(userindex)

1850      Case ClientPacketID.GuildRejectNewMember    'RECHAZAR
1860          Call HandleGuildRejectNewMember(userindex)

1870      Case ClientPacketID.GuildKickMember         'ECHARCLA
1880          Call HandleGuildKickMember(userindex)

1890      Case ClientPacketID.GuildUpdateNews         'ACTGNEWS
1900          Call HandleGuildUpdateNews(userindex)

1910      Case ClientPacketID.GuildMemberInfo         '1HRINFO<
1920          Call HandleGuildMemberInfo(userindex)

1930      Case ClientPacketID.GuildOpenElections      'ABREELEC
1940          Call HandleGuildOpenElections(userindex)

1950      Case ClientPacketID.GuildRequestMembership  'SOLICITUD
1960          Call HandleGuildRequestMembership(userindex)

1970      Case ClientPacketID.GuildRequestDetails     'CLANDETAILS
1980          Call HandleGuildRequestDetails(userindex)

1990      Case ClientPacketID.Online                  '/ONLINE
2000          Call HandleOnline(userindex)

2010      Case ClientPacketID.Quit                    '/SALIR
2020          Call HandleQuit(userindex)

2030      Case ClientPacketID.GuildLeave              '/SALIRCLAN
2040          Call HandleGuildLeave(userindex)


2050      Case ClientPacketID.RequestAccountState     '/BALANCE
2060          Call HandleRequestAccountState(userindex)

2070      Case ClientPacketID.PetStand                '/QUIETO
2080          Call HandlePetStand(userindex)

2090      Case ClientPacketID.PetFollow               '/ACOMPAÑAR
2100          Call HandlePetFollow(userindex)

2110      Case ClientPacketID.ReleasePet              '/LIBERAR
2120          Call HandleReleasePet(userindex)

2130      Case ClientPacketID.TrainList               '/ENTRENAR
2140          Call HandleTrainList(userindex)

2150      Case ClientPacketID.Rest                    '/DESCANSAR
2160          Call HandleRest(userindex)

2170      Case ClientPacketID.Meditate                '/MEDITAR
2180          Call HandleMeditate(userindex)

2190      Case ClientPacketID.Verpenas
2200          Call Handleverpenas(userindex)

2210      Case ClientPacketID.DropItems           '/CAER
2220          Call HandleDropItems(userindex)

2230      Case ClientPacketID.Fianzah
2240          Call HandleFianzah(userindex)

2250      Case ClientPacketID.Resucitate              '/RESUCITAR
2260          Call HandleResucitate(userindex)

2270      Case ClientPacketID.Heal                    '/CURAR
2280          Call HandleHeal(userindex)

2290      Case ClientPacketID.Help                    '/AYUDA
2300          Call HandleHelp(userindex)

2310      Case ClientPacketID.RequestStats            '/EST
2320          Call HandleRequestStats(userindex)

2330      Case ClientPacketID.CommerceStart           '/COMERCIAR
2340          Call HandleCommerceStart(userindex)

2350      Case ClientPacketID.BankStart               '/BOVEDA
2360          Call HandleBankStart(userindex)

2370      Case ClientPacketID.Enlist                  '/ENLISTAR
2380          Call HandleEnlist(userindex)

2390      Case ClientPacketID.Information             '/INFORMACION
2400          Call HandleInformation(userindex)

2410      Case ClientPacketID.Reward                  '/RECOMPENSA
2420          Call HandleReward(userindex)

2430      Case ClientPacketID.UpTime                  '/UPTIME
2440          Call HandleUpTime(userindex)

2510      Case ClientPacketID.Inquiry                 '/ENCUESTA ( with no params )
2520          Call HandleInquiry(userindex)

2530      Case ClientPacketID.GuildMessage            '/CMSG
2540          Call HandleGuildMessage(userindex)

2550      Case ClientPacketID.GroupMessage            '/PMSG
2560          Call HandleGroupMessage(userindex)

2570      Case ClientPacketID.CentinelReport          '/CENTINELA
2580          Call HandleCentinelReport(userindex)

2590      Case ClientPacketID.GuildOnline             '/ONLINECLAN
2600          Call HandleGuildOnline(userindex)

2630      Case ClientPacketID.CouncilMessage          '/BMSG
2640          Call HandleCouncilMessage(userindex)

2650      Case ClientPacketID.RoleMasterRequest       '/ROL
2660          Call HandleRoleMasterRequest(userindex)

2670      Case ClientPacketID.GMRequest               '/GM
2680          Call HandleGMRequest(userindex)

2690      Case ClientPacketID.ChangeDescription       '/DESC
2700          Call HandleChangeDescription(userindex)

2710      Case ClientPacketID.GuildVote               '/VOTO
2720          Call HandleGuildVote(userindex)

2730      Case ClientPacketID.Punishments             '/PENAS
2740          Call HandlePunishments(userindex)

2750      Case ClientPacketID.ChangePassword          '/CONTRASEÑA
2760          Call HandleChangePassword(userindex)

2770      Case ClientPacketID.ChangePin         '/CONTRASEÑA
2780          Call HandleChangePin(userindex)

2790      Case ClientPacketID.Gamble                  '/APOSTAR
2800          Call HandleGamble(userindex)

2810      Case ClientPacketID.InquiryVote             '/ENCUESTA ( with parameters )
2820          Call HandleInquiryVote(userindex)

2830      Case ClientPacketID.LeaveFaction            '/RETIRAR ( with no arguments )
2840          Call HandleLeaveFaction(userindex)

2850      Case ClientPacketID.BankExtractGold         '/RETIRAR ( with arguments )
2860          Call HandleBankExtractGold(userindex)

2870      Case ClientPacketID.BankDepositGold         '/DEPOSITAR
2880          Call HandleBankDepositGold(userindex)

2890      Case ClientPacketID.Denounce                '/DENUNCIAR
2900          Call HandleDenounce(userindex)

2910      Case ClientPacketID.GuildFundate            '/FUNDARCLAN
2920          Call HandleGuildFundate(userindex)

2930      Case ClientPacketID.GuildFundation
2940          Call HandleGuildFundation(userindex)

3010      Case ClientPacketID.rCaptions
3020          Call HandleRequieredCaptions(userindex)

3030      Case ClientPacketID.SCaptions
3040          Call HandleSendCaptions(userindex)

3050      Case ClientPacketID.Ping                    '/PING
3060          Call HandlePing(userindex)

3070      Case ClientPacketID.Cara                    '/Cara
3080          Call HandleCara(userindex)

3090      Case ClientPacketID.Viajar
3100          Call HandleViajar(userindex)

3110      Case ClientPacketID.ItemUpgrade
3120          Call HandleItemUpgrade(userindex)

3130      Case ClientPacketID.GMCommands              'GM Messages
3140          Call HandleGMCommands(userindex)

3150      Case ClientPacketID.InitCrafting
3160          Call HandleInitCrafting(userindex)

3170      Case ClientPacketID.Home
3180          Call HandleHome(userindex)

3190      Case ClientPacketID.ShowGuildNews
3200          Call HandleShowGuildNews(userindex)

3210      Case ClientPacketID.ShareNpc
3220          Call HandleShareNpc(userindex)

3230      Case ClientPacketID.StopSharingNpc
3240          Call HandleStopSharingNpc(userindex)

3250      Case ClientPacketID.Consulta
3260          Call HandleConsultation(userindex)

3270      Case ClientPacketID.SolicitaRranking
3280          Call HandleSolicitarRanking(userindex)

3290      Case ClientPacketID.Quest                   '/QUEST
3300          Call HandleQuest(userindex)

3310      Case ClientPacketID.QuestAccept
3320          Call HandleQuestAccept(userindex)

3330      Case ClientPacketID.QuestListRequest
3340          Call HandleQuestListRequest(userindex)

3350      Case ClientPacketID.QuestDetailsRequest
3360          Call HandleQuestDetailsRequest(userindex)

3370      Case ClientPacketID.QuestAbandon
3380          Call HandleQuestAbandon(userindex)

3390      Case ClientPacketID.ResetearPj                 '/RESET
3400          Call HandleResetearPJ(userindex)

3410      Case ClientPacketID.Nivel               '/RESET
3420          Call HanDlenivel(userindex)


3430      Case ClientPacketID.usarbono
3440          Call HandleUsarBono(userindex)

3450      Case ClientPacketID.Oro
3460          Call HandleOro(userindex)
              
3470      Case ClientPacketID.Premium
3480          Call HandlePremium(userindex)
              
3510      Case ClientPacketID.RightClick
3520          Call HandleRightClick(userindex)

3530      Case ClientPacketID.Plata
3540          Call HandlePlata(userindex)

3550      Case ClientPacketID.Bronce
3560          Call HandleBronce(userindex)

3570      Case ClientPacketID.GlobalMessage
3580          Call HandleGlobalMessage(userindex)

3590      Case ClientPacketID.GlobalStatus
3600          Call HandleGlobalStatus(userindex)

3610      Case ClientPacketID.CuentaRegresiva      '/CR
3620          Call HandleCuentaRegresiva(userindex)

3630      Case ClientPacketID.dragInventory        'DINVENT
3640          Call HandleDragInventory(userindex)

3650      Case ClientPacketID.DragToPos               'DTOPOS
3660          Call HandleDragToPos(userindex)

3670      Case ClientPacketID.DragToggle
3680          Call HandleDragToggle(userindex)

3730      Case ClientPacketID.solicitudes              '/DENUNCIAR
3740          Call HandleSolicitud(userindex)

3750      Case ClientPacketID.SetMenu
3760          Call HandleSetMenu(userindex)
          
3770      Case ClientPacketID.WherePower
3780          Call HandleWherePower(userindex)
          
          
          Case ClientPacketID.RequestMercado
            Call HandleRequestMercado(userindex)
            
          Case ClientPacketID.ReceiveOfferAccount
            Call HandleReceiveOfferAccount(userindex)
            
            Case ClientPacketID.RequestInfoCharMAO
                Call HandleRequestInfoCharMAO(userindex)
                
            Case ClientPacketID.PublicationPj
                Call HandlePublicationPj(userindex)
            
            Case ClientPacketID.InvitationChange
                Call HandleInvitationChange(userindex)
                
            Case ClientPacketID.BuyPj
                Call HandleBuyPj(userindex)
                
            Case ClientPacketID.QuitarPj
                Call HandleQuitarPj(userindex)
            
            Case ClientPacketID.RequestOfferSent
                Call HandleRequestOfferSent(userindex)
                
            Case ClientPacketID.RequestOffer
                Call HandleRequestOffer(userindex)
                
            Case ClientPacketID.AcceptInvitation
                Call HandleAcceptInvitation(userindex)
                
            Case ClientPacketID.RechaceInvitation
                Call HandleRechaceInvitation(userindex)
                
            Case ClientPacketID.CancelInvitation
                Call HandleCancelInvitation(userindex)
                
            Case ClientPacketID.EnviarAviso
                Call HandleEnviarAviso(userindex)
                
            Case ClientPacketID.seguroclan
                Call HandleSeguroClan(userindex)
          
3790      Case Else
                  'ERROR : Abort!
3800              Call CloseSocket(userindex)

3810      End Select
                  
          
3820      UserList(userindex).LastPacket = PacketID
          UserList(userindex).incomingData.LastPacket = PacketID
          
      'Done with this packet, move on to next one or send everything if no more packets found
3830      If UserList(userindex).incomingData.length > 0 And Err.Number = 0 Then
3840          Err.Clear
3850          HandleIncomingData = True

              'UserList(Userindex).PaquetesBasura = UserList(Userindex).PaquetesBasura + 1
               
              'If UserList(Userindex).PaquetesBasura = 1000 Then
                  'UserList(Userindex).PaquetesBasura = 0
                 ' Call LogError("El usuario " & UserList(Userindex).Name & " cuya IP es " & UserList(Userindex).ip & " quiso saturar el protocolo.")
              'End If

3860      ElseIf Err.Number <> 0 And Not Err.Number = UserList(userindex).incomingData.NotEnoughDataErrCode Then
              'An error ocurred, log it and kick player.
3870          Call LogError("Error: " & Err.Number & " [" & Err.Description & "] " & " Source: " & Err.source & _
                              vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & _
                              vbTab & " LastDllError: " & Err.LastDllError & vbTab & _
                              " - UserIndex: " & userindex & " - producido al manejar el paquete: " & CStr(PacketID) & ", LastPacket: " & UserList(userindex).LastPacket)
              
3880          Call CloseSocket(userindex)
       
              'UserList(Userindex).PaquetesBasura = 0
3890          HandleIncomingData = False
3900      Else
              'Flush buffer - send everything that has been written
3910          Call FlushBuffer(userindex)
3920          HandleIncomingData = False
             ' UserList(Userindex).PaquetesBasura = 0
3930      End If
          
          
3940  Exit Function

Errhandler:
3950      HandleIncomingData = False
3960      Call LogError("Error ERRHANDLER: " & Err.Number & " [" & Err.Description & "] " & " Source: " & Err.source & _
                              vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & _
                              vbTab & " LastDllError: " & Err.LastDllError & vbTab & _
                              " - UserIndex: " & UserList(userindex).Name & " - producido al manejar el paquete: " & CStr(PacketID) & ", LastPacket: " & UserList(userindex).LastPacket)
End Function

Public Sub WriteMultiMessage(ByVal userindex As Integer, ByVal MessageIndex As Integer, Optional ByVal Arg1 As Long, Optional ByVal Arg2 As Long, Optional ByVal Arg3 As Long, Optional ByVal StringArg1 As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

   On Error GoTo WriteMultiMessage_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.MultiMessage)
40            Call .WriteByte(MessageIndex)

50            Select Case MessageIndex
              Case eMessages.DontSeeAnything, eMessages.NPCSwing, eMessages.NPCKillUser, eMessages.BlockedWithShieldUser, _
                   eMessages.BlockedWithShieldother, eMessages.UserSwing, eMessages.SafeModeOn, eMessages.SafeModeOff, eMessages.DragOnn, eMessages.DragOff, _
                   eMessages.ResuscitationSafeOff, eMessages.ResuscitationSafeOn, eMessages.NobilityLost, _
                   eMessages.CantUseWhileMeditating, eMessages.CancelHome, eMessages.FinishHome

60            Case eMessages.NPCHitUser
70                Call .WriteByte(Arg1)    'Target
80                Call .WriteInteger(Arg2)    'damage

90            Case eMessages.UserHitNPC
100               Call .WriteLong(Arg1)    'damage

110           Case eMessages.UserAttackedSwing
120               Call .WriteInteger(UserList(Arg1).Char.CharIndex)

130           Case eMessages.UserHittedByUser
140               Call .WriteInteger(Arg1)    'AttackerIndex
150               Call .WriteByte(Arg2)    'Target
160               Call .WriteInteger(Arg3)    'damage

170           Case eMessages.UserHittedUser
180               Call .WriteInteger(Arg1)    'AttackerIndex
190               Call .WriteByte(Arg2)    'Target
200               Call .WriteInteger(Arg3)    'damage

210           Case eMessages.WorkRequestTarget
220               Call .WriteByte(Arg1)    'skill

230           Case eMessages.HaveKilledUser    '"Has matado a " & UserList(VictimIndex).name & "!" "Has ganado " & DaExp & " puntos de experiencia."
240               Call .WriteInteger(UserList(Arg1).Char.CharIndex)    'VictimIndex
250               Call .WriteLong(Arg2)    'Expe

260           Case eMessages.UserKill    '"¡" & .name & " te ha matado!"
270               Call .WriteInteger(UserList(Arg1).Char.CharIndex)    'AttackerIndex

280           Case eMessages.EarnExp

290           Case eMessages.Home
300               Call .WriteByte(CByte(Arg1))
310               Call .WriteInteger(CInt(Arg2))
                  'El cliente no conoce nada sobre nombre de mapas y hogares, por lo tanto _
                   hasta que no se pasen los dats e .INFs al cliente, esto queda así.
320               Call .WriteASCIIString(StringArg1)    'Call .WriteByte(CByte(Arg2))

330           End Select
340       End With
350       Exit Sub    ''

Errhandler:
360       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
370           Call FlushBuffer(userindex)
380           Resume
390       End If

   On Error GoTo 0
   Exit Sub

WriteMultiMessage_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteMultiMessage of Módulo Protocol in line " & Erl
End Sub
Private Sub HandleGMCommands(ByVal userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        On Error GoTo Errhandler

          Dim Command As Byte

20        With UserList(userindex)
30            Call .incomingData.ReadByte
              
40            Command = .incomingData.PeekByte

60             LogGM "Gm_General", "El personaje " & .Name & " con IP: " & .ip & " uso el HnadleGmCommands. PacketID: " & Command & ". LastPacket: " & .LastPacket & ", Lenght: " & .incomingData.length
              
90            Select Case Command
              
              Case eGMCommands.CreateInvasion
                 Call HandleCreateInvasion(userindex)
                 
              Case eGMCommands.TerminateInvasion
                Call HandleTerminateInvasion(userindex)
                 
              Case eGMCommands.GMMessage                '/GMSG
100               Call HandleGMMessage(userindex)

110           Case eGMCommands.showName                '/SHOWNAME
120               Call HandleShowName(userindex)

130           Case eGMCommands.OnlineRoyalArmy
140               Call HandleOnlineRoyalArmy(userindex)

150           Case eGMCommands.OnlineChaosLegion       '/ONLINECAOS
160               Call HandleOnlineChaosLegion(userindex)

170           Case eGMCommands.GoNearby                '/IRCERCA
180               Call HandleGoNearby(userindex)

190           Case eGMCommands.SeBusca                '/SEBUSCA
200               Call Elmasbuscado(userindex)

210           Case eGMCommands.comment                 '/REM
220               Call HandleComment(userindex)

230           Case eGMCommands.serverTime              '/HORA
240               Call HandleServerTime(userindex)

250           Case eGMCommands.Where                   '/DONDE
260               Call HandleWhere(userindex)

270           Case eGMCommands.CreaturesInMap          '/NENE
280               Call HandleCreaturesInMap(userindex)

290           Case eGMCommands.WarpMeToTarget          '/TELEPLOC
300               Call HandleWarpMeToTarget(userindex)

310           Case eGMCommands.WarpChar                '/TELEP
320               Call HandleWarpChar(userindex)

330           Case eGMCommands.Silence                 '/SILENCIAR
340               Call HandleSilence(userindex)

350           Case eGMCommands.SOSShowList             '/SHOW SOS
360               Call HandleSOSShowList(userindex)

370           Case eGMCommands.SOSRemove               'SOSDONE
380               Call HandleSOSRemove(userindex)

390           Case eGMCommands.GoToChar                '/IRA
400               Call HandleGoToChar(userindex)

410           Case eGMCommands.invisible               '/INVISIBLE
420               Call HandleInvisible(userindex)

430           Case eGMCommands.GMPanel                 '/PANELGM
440               Call HandleGMPanel(userindex)

450           Case eGMCommands.RequestUserList         'LISTUSU
460               Call HandleRequestUserList(userindex)

470           Case eGMCommands.Working                 '/TRABAJANDO
480               Call HandleWorking(userindex)

490           Case eGMCommands.Hiding                  '/OCULTANDO
500               Call HandleHiding(userindex)

510           Case eGMCommands.Jail                    '/CARCEL
520               Call HandleJail(userindex)

530           Case eGMCommands.KillNPC                 '/RMATA
540               Call HandleKillNPC(userindex)

550           Case eGMCommands.WarnUser                '/ADVERTENCIA
560               Call HandleWarnUser(userindex)

570           Case eGMCommands.RequestCharInfo         '/INFO
580               Call HandleRequestCharInfo(userindex)

590           Case eGMCommands.RequestCharStats        '/STAT
600               Call HandleRequestCharStats(userindex)

610           Case eGMCommands.RequestCharGold         '/BAL
620               Call HandleRequestCharGold(userindex)

630           Case eGMCommands.RequestCharInventory    '/INV
640               Call HandleRequestCharInventory(userindex)

650           Case eGMCommands.RequestCharBank         '/BOV
660               Call HandleRequestCharBank(userindex)

670           Case eGMCommands.RequestCharSkills       '/SKILLS
680               Call HandleRequestCharSkills(userindex)

690           Case eGMCommands.ReviveChar              '/REVIVIR
700               Call HandleReviveChar(userindex)

710           Case eGMCommands.OnlineGM                '/ONLINEGM
720               Call HandleOnlineGM(userindex)

730           Case eGMCommands.OnlineMap               '/ONLINEMAP
740               Call HandleOnlineMap(userindex)

750           Case eGMCommands.Forgive                 '/PERDON
760               Call HandleForgive(userindex)

770           Case eGMCommands.Kick                    '/ECHAR
780               Call HandleKick(userindex)

790           Case eGMCommands.Execute                 '/EJECUTAR
800               Call HandleExecute(userindex)

810           Case eGMCommands.banChar                 '/BAN
820               Call HandleBanChar(userindex)

830           Case eGMCommands.UnbanChar               '/UNBAN
840               Call HandleUnbanChar(userindex)

850           Case eGMCommands.NPCFollow               '/SEGUIR
860               Call HandleNPCFollow(userindex)

870           Case eGMCommands.SummonChar              '/SUM
880               Call HandleSummonChar(userindex)

890           Case eGMCommands.SpawnListRequest        '/CC
900               Call HandleSpawnListRequest(userindex)

910           Case eGMCommands.SpawnCreature           'SPA
920               Call HandleSpawnCreature(userindex)

930           Case eGMCommands.ResetNPCInventory       '/RESETINV
940               Call HandleResetNPCInventory(userindex)

950           Case eGMCommands.cleanworld              '/LIMPIAR
960               Call HandleCleanWorld(userindex)

970           Case eGMCommands.ServerMessage           '/RMSG
980               Call HandleServerMessage(userindex)

990           Case eGMCommands.RolMensaje           '/ROLEANDO
1000              Call HandleRolMensaje(userindex)

1010          Case eGMCommands.nickToIP                '/NICK2IP
1020              Call HandleNickToIP(userindex)

1030          Case eGMCommands.IPToNick                '/IP2NICK
1040              Call HandleIPToNick(userindex)

1050          Case eGMCommands.GuildOnlineMembers      '/ONCLAN
1060              Call HandleGuildOnlineMembers(userindex)

1070          Case eGMCommands.TeleportCreate          '/CT
1080              Call HandleTeleportCreate(userindex)

1090          Case eGMCommands.TeleportDestroy         '/DT
1100              Call HandleTeleportDestroy(userindex)

1110          Case eGMCommands.RainToggle              '/LLUVIA
1120              Call HandleRainToggle(userindex)

1130          Case eGMCommands.SetCharDescription      '/SETDESC
1140              Call HandleSetCharDescription(userindex)

1150          Case eGMCommands.ForceMIDIToMap          '/FORCEMIDIMAP
1160              Call HanldeForceMIDIToMap(userindex)

1170          Case eGMCommands.ForceWAVEToMap          '/FORCEWAVMAP
1180              Call HandleForceWAVEToMap(userindex)

1190          Case eGMCommands.RoyalArmyMessage        '/REALMSG
1200              Call HandleRoyalArmyMessage(userindex)

1210          Case eGMCommands.ChaosLegionMessage      '/CAOSMSG
1220              Call HandleChaosLegionMessage(userindex)

1230          Case eGMCommands.CitizenMessage          '/CIUMSG
1240              Call HandleCitizenMessage(userindex)

1250          Case eGMCommands.CriminalMessage         '/CRIMSG
1260              Call HandleCriminalMessage(userindex)

1270          Case eGMCommands.TalkAsNPC               '/TALKAS
1280              Call HandleTalkAsNPC(userindex)

1290          Case eGMCommands.DestroyAllItemsInArea   '/MASSDEST
1300              Call HandleDestroyAllItemsInArea(userindex)

1310          Case eGMCommands.AcceptRoyalCouncilMember    '/ACEPTCONSE
1320              Call HandleAcceptRoyalCouncilMember(userindex)

1330          Case eGMCommands.AcceptChaosCouncilMember    '/ACEPTCONSECAOS
1340              Call HandleAcceptChaosCouncilMember(userindex)

1350          Case eGMCommands.ItemsInTheFloor         '/PISO
1360              Call HandleItemsInTheFloor(userindex)

1370          Case eGMCommands.MakeDumb                '/ESTUPIDO
1380              Call HandleMakeDumb(userindex)

1390          Case eGMCommands.MakeDumbNoMore          '/NOESTUPIDO
1400              Call HandleMakeDumbNoMore(userindex)

1410          Case eGMCommands.dumpIPTables            '/DUMPSECURITY
1420              Call HandleDumpIPTables(userindex)

1430          Case eGMCommands.CouncilKick             '/KICKCONSE
1440              Call HandleCouncilKick(userindex)

1450          Case eGMCommands.SetTrigger              '/TRIGGER
1460              Call HandleSetTrigger(userindex)

1470          Case eGMCommands.AskTrigger              '/TRIGGER with no args
1480              Call HandleAskTrigger(userindex)

1490          Case eGMCommands.BannedIPList            '/BANIPLIST
1500              Call HandleBannedIPList(userindex)

1510          Case eGMCommands.BannedIPReload          '/BANIPRELOAD
1520              Call HandleBannedIPReload(userindex)

1530          Case eGMCommands.GuildMemberList         '/MIEMBROSCLAN
1540              Call HandleGuildMemberList(userindex)

1550          Case eGMCommands.GuildBan                '/BANCLAN
1560              Call HandleGuildBan(userindex)

1570          Case eGMCommands.BanIP                   '/BANIP
1580              Call HandleBanIP(userindex)

1590          Case eGMCommands.UnbanIP                 '/UNBANIP
1600              Call HandleUnbanIP(userindex)

1610          Case eGMCommands.CreateItem              '/CI
1620              Call HandleCreateItem(userindex)

1630          Case eGMCommands.DestroyItems            '/DEST
1640              Call HandleDestroyItems(userindex)

1650          Case eGMCommands.ChaosLegionKick         '/NOCAOS
1660              Call HandleChaosLegionKick(userindex)

1670          Case eGMCommands.RoyalArmyKick           '/NOREAL
1680              Call HandleRoyalArmyKick(userindex)

1690          Case eGMCommands.ForceMIDIAll            '/FORCEMIDI
1700              Call HandleForceMIDIAll(userindex)

1710          Case eGMCommands.ForceWAVEAll            '/FORCEWAV
1720              Call HandleForceWAVEAll(userindex)

1730          Case eGMCommands.RemovePunishment        '/BORRARPENA
1740              Call HandleRemovePunishment(userindex)

1750          Case eGMCommands.TileBlockedToggle       '/BLOQ
1760              Call HandleTileBlockedToggle(userindex)

1770          Case eGMCommands.KillNPCNoRespawn        '/MATA
1780              Call HandleKillNPCNoRespawn(userindex)

1790          Case eGMCommands.KillAllNearbyNPCs       '/MASSKILL
1800              Call HandleKillAllNearbyNPCs(userindex)

1810          Case eGMCommands.lastip                  '/LASTIP
1820              Call HandleLastIP(userindex)

1830          Case eGMCommands.SystemMessage           '/SMSG
1840              Call HandleSystemMessage(userindex)

1850          Case eGMCommands.CreateNPC               '/ACC
1860              Call HandleCreateNPC(userindex)

1870          Case eGMCommands.CreateNPCWithRespawn    '/RACC
1880              Call HandleCreateNPCWithRespawn(userindex)

1890          Case eGMCommands.ImperialArmour          '/AI1 - 4
1900              Call HandleImperialArmour(userindex)

1910          Case eGMCommands.ChaosArmour             '/AC1 - 4
1920              Call HandleChaosArmour(userindex)

1930          Case eGMCommands.NavigateToggle          '/NAVE
1940              Call HandleNavigateToggle(userindex)

1950          Case eGMCommands.ServerOpenToUsersToggle    '/HABILITAR
1960              Call HandleServerOpenToUsersToggle(userindex)

1970          Case eGMCommands.TurnOffServer           '/APAGAR
1980              Call HandleTurnOffServer(userindex)

1990          Case eGMCommands.TurnCriminal            '/CONDEN
2000              Call HandleTurnCriminal(userindex)

2010          Case eGMCommands.ResetFactionCaos           '/RAJAR
2020              Call HandleResetFactionCaos(userindex)

2030          Case eGMCommands.ResetFactionReal           '/RAJAR
2040              Call HandleResetFactionReal(userindex)

2050          Case eGMCommands.RemoveCharFromGuild     '/RAJARCLAN
2060              Call HandleRemoveCharFromGuild(userindex)

2070          Case eGMCommands.RequestCharMail         '/LASTEMAIL
2080              Call HandleRequestCharMail(userindex)

2090          Case eGMCommands.AlterPassword           '/APASS
2100              Call HandleAlterPassword(userindex)

2110          Case eGMCommands.AlterMail               '/AEMAIL
2120              Call HandleAlterMail(userindex)

2130          Case eGMCommands.AlterName               '/ANAME
2140              Call HandleAlterName(userindex)

2150          Case eGMCommands.ToggleCentinelActivated    '/CENTINELAACTIVADO
2160              Call HandleToggleCentinelActivated(userindex)

2170          Case Declaraciones.eGMCommands.DoBackUp               '/DOBACKUP
2180              Call HandleDoBackUp(userindex)

2190          Case eGMCommands.ShowGuildMessages       '/SHOWCMSG
2200              Call HandleShowGuildMessages(userindex)

2210          Case eGMCommands.SaveMap                 '/GUARDAMAPA
2220              Call HandleSaveMap(userindex)

2230          Case eGMCommands.ChangeMapInfoPK         '/MODMAPINFO PK
2240              Call HandleChangeMapInfoPK(userindex)

2250          Case eGMCommands.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
2260              Call HandleChangeMapInfoBackup(userindex)

2270          Case eGMCommands.ChangeMapInfoRestricted    '/MODMAPINFO RESTRINGIR
2280              Call HandleChangeMapInfoRestricted(userindex)

2290          Case eGMCommands.ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
2300              Call HandleChangeMapInfoNoMagic(userindex)

2310          Case eGMCommands.ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
2320              Call HandleChangeMapInfoNoInvi(userindex)

2330          Case eGMCommands.ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
2340              Call HandleChangeMapInfoNoResu(userindex)

2350          Case eGMCommands.ChangeMapInfoLand       '/MODMAPINFO TERRENO
2360              Call HandleChangeMapInfoLand(userindex)

2370          Case eGMCommands.ChangeMapInfoZone       '/MODMAPINFO ZONA
2380              Call HandleChangeMapInfoZone(userindex)

2390          Case eGMCommands.ChangeMapInfoStealNpc   '/MODMAPINFO ROBONPC
2400              Call HandleChangeMapInfoStealNpc(userindex)

2410          Case eGMCommands.ChangeMapInfoNoOcultar  '/MODMAPINFO OCULTARSINEFECTO
2420              Call HandleChangeMapInfoNoOcultar(userindex)

2430          Case eGMCommands.ChangeMapInfoNoInvocar  '/MODMAPINFO INVOCARSINEFECTO
2440              Call HandleChangeMapInfoNoInvocar(userindex)

2450          Case eGMCommands.SaveChars               '/GRABAR
2460              Call HandleSaveChars(userindex)

2470          Case eGMCommands.CleanSOS                '/BORRAR SOS
2480              Call HandleCleanSOS(userindex)

2490          Case eGMCommands.ShowServerForm          '/SHOW INT
2500              Call HandleShowServerForm(userindex)

2510          Case eGMCommands.night                   '/NOCHE
2520              Call HandleNight(userindex)


2530          Case eGMCommands.KickAllChars            '/ECHARTODOSPJS
2540              Call HandleKickAllChars(userindex)

2550          Case eGMCommands.ReloadNPCs              '/RELOADNPCS
2560              Call HandleReloadNPCs(userindex)

2570          Case eGMCommands.ReloadServerIni         '/RELOADSINI
2580              Call HandleReloadServerIni(userindex)

2590          Case eGMCommands.ReloadSpells            '/RELOADHECHIZOS
2600              Call HandleReloadSpells(userindex)

2610          Case eGMCommands.ReloadObjects           '/RELOADOBJ
2620              Call HandleReloadObjects(userindex)

2630          Case eGMCommands.Restart                 '/REINICIAR
2640              Call HandleRestart(userindex)

2650          Case eGMCommands.ResetAutoUpdate         '/AUTOUPDATE
2660              Call HandleResetAutoUpdate(userindex)

2670          Case eGMCommands.ChatColor               '/CHATCOLOR
2680              Call HandleChatColor(userindex)

2690          Case eGMCommands.Ignored                 '/IGNORADO
2700              Call HandleIgnored(userindex)

2710          Case eGMCommands.CheckSlot               '/SLOT
2720              Call HandleCheckSlot(userindex)


2730          Case eGMCommands.SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
2740              Call HandleSetIniVar(userindex)
                
                
                Case eGMCommands.SearchNpc
            Call HandleSearchNpc(userindex)
           
        Case eGMCommands.SearchObj
            Call HandleSearchObj(userindex)
           
        Case eGMCommands.SearcherShow            '/BUSCADOR
            Call HandleSearcherShow(userindex)
            
2750          Case eGMCommands.Seguimiento
2760              Call HandleSeguimiento(userindex)

                  '//Disco.
2770          Case eGMCommands.CheckHD                 '/VERHD NICKUSUARIO
2780              Call HandleCheckHD(userindex)

2790          Case eGMCommands.BanHD                   '/BANHD NICKUSUARIO
2800              Call HandleBanHD(userindex)

2810          Case eGMCommands.UnBanHD                 '/UNBANHD NICKUSUARIO
2820              Call HandleUnbanHD(userindex)
                  '///Disco.

2830          Case eGMCommands.MapMessage              '/MAPMSG
2840              Call HandleMapMessage(userindex)

2850          Case eGMCommands.Impersonate             '/IMPERSONAR
2860              Call HandleImpersonate(userindex)

2870          Case eGMCommands.Imitate                 '/MIMETIZAR
2880              Call HandleImitate(userindex)


2890          Case eGMCommands.CambioPj                '/CAMBIO
2900              Call HandleCambioPj(userindex)
                  
2910          Case eGMCommands.LarryMataNiños
2920              Call HandleLarryMataNiños(userindex)
                  
2930          Case eGMCommands.ComandoPorDias
2940              Call HandleComandoPorDias(userindex)
                  
2950          Case eGMCommands.DarPoints
2960              Call HandleDarPoints(userindex)

2970          End Select
2980      End With

2990      Exit Sub

Errhandler:
3000      Call LogError("Error en GmCommands. Error: " & Err.Number & " - " & Err.Description & _
                        ". Paquete: " & Command)

End Sub

' ME VOY A FUMAR 340 PAQUETES POR VOS ALAN EMPEZANDO YA

''
' Handles the "Home" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleHome(ByVal userindex As Integer)
      '***************************************************
      'Author: Budi
      'Creation Date: 06/01/2010
      'Last Modification: 05/06/10
      'Pato - 05/06/10: Add the Ucase$ to prevent problems.
      '***************************************************
10        With UserList(userindex)
20            Call .incomingData.ReadByte
              
30            If .flags.SlotEvent > 0 Then
40                WriteConsoleMsg userindex, "No puedes usar la restauración si estás en un evento.", FontTypeNames.FONTTYPE_INFO
50                Exit Sub
60            End If
              
70            If .flags.InCVC Then
80                WriteConsoleMsg userindex, "No puedes usar la restauración si estás en CVC.", FontTypeNames.FONTTYPE_INFO
90                Exit Sub
100           End If
              
110           If .flags.SlotReto > 0 Then
120               WriteConsoleMsg userindex, "No puede susar este comando si estás en reto.", FontTypeNames.FONTTYPE_INFO
130               Exit Sub
140           End If
              
150           If .Pos.map = 66 Then
160               Call WriteConsoleMsg(userindex, "No puedes usar la restauración si estás en la carcel.", FontTypeNames.FONTTYPE_INFO)
170               Exit Sub
180           End If
190           If .Pos.map = 191 Then
200               Call WriteConsoleMsg(userindex, "No puedes usar la restauración si estás en los retos.", FontTypeNames.FONTTYPE_INFO)
210               Exit Sub
220           End If
              
230           If .Pos.map = 176 Then
240               Call WriteConsoleMsg(userindex, "No puedes usar la restauración si estás en los retos.", FontTypeNames.FONTTYPE_INFO)
250               Exit Sub
260           End If

270           If .flags.Muerto = 0 Then
280               Call WriteConsoleMsg(userindex, "No puedes usar el comando si estás vivo.", FontTypeNames.FONTTYPE_INFO)
290               Exit Sub
300           End If

310           If UserList(userindex).Stats.Gld < 7000 Then
320               Call WriteConsoleMsg(userindex, "No tienes suficientes monedas de oro, necesitas 7.000 monedas para usar la restauración de personaje.", FontTypeNames.FONTTYPE_INFO)
330               Exit Sub
340           End If

350           UserList(userindex).Stats.Gld = UserList(userindex).Stats.Gld - 7000

360           Call WriteUpdateGold(userindex)
370           WriteUpdateUserStats (userindex)

380           If .flags.Muerto = 1 Then
390               Call WarpUserChar(userindex, 1, 59, 45, True)
400               Call WriteConsoleMsg(userindex, "Has sido transportado Ullathorpe.", FontTypeNames.FONTTYPE_INFO)
410               Exit Sub
420           End If
430       End With
End Sub

''
' Handles the "ThrowDices" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleThrowDices(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call UserList(userindex).incomingData.ReadByte

20        With UserList(userindex).Stats
30            .UserAtributos(eAtributos.Fuerza) = RandomNumber(17, 18)
40            .UserAtributos(eAtributos.Agilidad) = RandomNumber(17, 18)
50            .UserAtributos(eAtributos.Inteligencia) = RandomNumber(17, 18)
60            .UserAtributos(eAtributos.Carisma) = RandomNumber(17, 18)
70            .UserAtributos(eAtributos.Constitucion) = RandomNumber(17, 18)
80        End With

90        Call WriteDiceRoll(userindex)
End Sub

Private Sub HandleTalk(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 23/09/2009
      '15/07/2009: ZaMa - Now invisible admins talk by console.
      '23/09/2009: ZaMa - Now invisible admins can't send empty chat.
      '***************************************************

   On Error GoTo HandleTalk_Error

10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)

              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
              Dim CanTalk As Boolean
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim chat As String

90            chat = buffer.ReadASCIIString()
              
                  '[Consejeros & GMs]
100               If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
110                   Call LogGM(.Name, "Dijo: " & chat)
120               End If
          
                  'I see you....
130               If .flags.Oculto > 0 Then
140                   .flags.Oculto = 0
150                   .Counters.TiempoOculto = 0
160                   If .flags.invisible = 0 Then
170                       Call UsUaRiOs.SetInvisible(userindex, UserList(userindex).Char.CharIndex, False)
                          'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
180                       Call WriteConsoleMsg(userindex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
190                   End If
200               End If
          
210               If LenB(chat) <> 0 Then
                      'Analize chat...
220                   Call Statistics.ParseChat(chat)
          
                      ' If Not (.flags.AdminInvisible = 1) Then
                      
230                   CanTalk = True
240                   If .flags.SlotEvent > 0 Then
250                       If Events(.flags.SlotEvent).Modality = DeathMatch Then
260                           CanTalk = False
270                       End If
280                   End If
                          
290                   If CanTalk Then
300                       If .flags.Muerto = 1 Then
310                           Call SendData(SendTarget.ToDeadArea, userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
320                       Else
330                           Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, .flags.ChatColor))
340                       End If
350                   End If
          
360               End If

              'If we got here then packet is complete, copy data back to original queue
370           Call .incomingData.CopyBuffer(buffer)
380       End With

Errhandler:
          Dim error  As Long
390       error = Err.Number
400       On Error GoTo 0

          'Destroy auxiliar buffer
410       Set buffer = Nothing

420       If error <> 0 Then _
             Err.Raise error

   On Error GoTo 0
   Exit Sub

HandleTalk_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleTalk of Módulo Protocol in line " & Erl & " (" & UserList(userindex).ip & ")"
End Sub

''
' Handles the "Yell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleYell(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 13/01/2010 (ZaMa)
      '15/07/2009: ZaMa - Now invisible admins yell by console.
      '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)

              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
70            Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
80            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
90            Call buffer.ReadByte

              Dim chat As String
              Dim CanTalk As Boolean
              
100           chat = buffer.ReadASCIIString()
                  '[Consejeros & GMs]
110               If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
120                   Call LogGM(.Name, "Grito: " & chat)
130               End If
          
                  'I see you....
140               If .flags.Oculto > 0 Then
150                   .flags.Oculto = 0
160                   .Counters.TiempoOculto = 0
          
170                   If .flags.Navegando = 1 Then
180                       If .clase = eClass.Pirat Then
                              ' Pierde la apariencia de fragata fantasmal
190                           Call ToggleBoatBody(userindex)
200                           Call WriteConsoleMsg(userindex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
210                           Call ChangeUserChar(userindex, .Char.body, .Char.Head, .Char.Heading, NingunArma, _
                                                  NingunEscudo, NingunCasco)
220                       End If
230                   Else
240                       If .flags.invisible = 0 Then
250                           Call UsUaRiOs.SetInvisible(userindex, .Char.CharIndex, False)
260                           Call WriteConsoleMsg(userindex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
270                       End If
280                   End If
290               End If
          
300               If LenB(chat) <> 0 Then
                      'Analize chat...
310                   Call Statistics.ParseChat(chat)
                      
320                   CanTalk = True
330                   If .flags.SlotEvent > 0 Then
340                       If Events(.flags.SlotEvent).Modality = DeathMatch Then
350                           CanTalk = False
360                       End If
370                   End If
                      
380                   If CanTalk Then
390                       If .flags.Privilegios And PlayerType.User Then
400                           If UserList(userindex).flags.Muerto = 1 Then
410                               Call SendData(SendTarget.ToDeadArea, userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
420                           Else
430                               Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, vbRed))
440                           End If
450                       Else
460                           If Not (.flags.AdminInvisible = 1) Then
470                               Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, CHAT_COLOR_GM_YELL))
480                           Else
490                           End If
500                       End If
510                   End If
520               End If


              'If we got here then packet is complete, copy data back to original queue
530           Call .incomingData.CopyBuffer(buffer)
540       End With

Errhandler:
          Dim error  As Long
550       error = Err.Number
560       On Error GoTo 0

          'Destroy auxiliar buffer
570       Set buffer = Nothing

580       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "Whisper" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 15/07/2009
      '28/05/2009: ZaMa - Now it doesn't appear any message when private talking to an invisible admin
      '***************************************************
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim chat As String
              Dim targetCharIndex As Integer
              Dim targetUserIndex As Integer
              Dim targetPriv As PlayerType

90            targetCharIndex = buffer.ReadInteger()
100           chat = buffer.ReadASCIIString()
              
110           targetUserIndex = CharIndexToUserIndex(targetCharIndex)
          
120               If .flags.Muerto Then
130                   Call WriteConsoleMsg(userindex, "¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. ", FontTypeNames.FONTTYPE_INFO)
140               Else
150                   If targetUserIndex = INVALID_INDEX Then
160                       Call WriteConsoleMsg(userindex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
170                   Else
180                       targetPriv = UserList(targetUserIndex).flags.Privilegios
                          'A los dioses y admins no vale susurrarles si no sos uno vos mismo (así no pueden ver si están conectados o no)
190                       If (targetPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 And (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Then
                              ' Controlamos que no este invisible
200                           If UserList(targetUserIndex).flags.AdminInvisible <> 1 Then
210                               Call WriteConsoleMsg(userindex, "No puedes susurrarle a los Dioses y Admins.", FontTypeNames.FONTTYPE_INFO)
220                           End If
                              'A los Consejeros y SemiDioses no vale susurrarles si sos un PJ común.
230                       ElseIf (.flags.Privilegios And PlayerType.User) <> 0 And (Not targetPriv And PlayerType.User) <> 0 Then
                              ' Controlamos que no este invisible
240                           If UserList(targetUserIndex).flags.AdminInvisible <> 1 Then
250                               Call WriteConsoleMsg(userindex, "No puedes susurrarle a los GMs.", FontTypeNames.FONTTYPE_INFO)
260                           End If
270                       ElseIf Not EstaPCarea(userindex, targetUserIndex) Then
280                           Call WriteConsoleMsg(userindex, "Estas muy lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
          
290                       Else
                              '[Consejeros & GMs]
300                           If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
310                               Call LogGM(.Name, "Le dijo a '" & UserList(targetUserIndex).Name & "' " & chat)
320                           End If
          
330                           If LenB(chat) <> 0 Then
                                  'Analize chat...
340                               Call Statistics.ParseChat(chat)
          
350                               If Not (.flags.AdminInvisible = 1) Then
360                                   Call WriteChatOverHead(userindex, chat, .Char.CharIndex, vbYellow)
370                                   Call WriteChatOverHead(targetUserIndex, chat, .Char.CharIndex, vbYellow)
380                                   Call FlushBuffer(targetUserIndex)
          
                                      '[CDT 17-02-2004]
390                                   If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
400                                       Call SendData(SendTarget.ToAdminsAreaButConsejeros, userindex, PrepareMessageChatOverHead("A " & UserList(targetUserIndex).Name & "> " & chat, .Char.CharIndex, vbYellow))
410                                   End If
420                               Else
430                                   Call SendData(SendTarget.ToAdminsAreaButConsejeros, userindex, PrepareMessageChatOverHead("A " & UserList(targetUserIndex).Name & "> " & chat, .Char.CharIndex, vbYellow))
                                      'Call WriteConsoleMsg(UserIndex, "Susurraste> " & chat, FontTypeNames.FONTTYPE_GM)
                                      'If UserIndex <> targetUserIndex Then Call WriteConsoleMsg(targetUserIndex, "Gm susurra> " & chat, FontTypeNames.FONTTYPE_GM)
          
                                      'If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                                      '    Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageConsoleMsg("Gm dijo a " & UserList(targetUserIndex).name & "> " & chat, FontTypeNames.FONTTYPE_GM))
                                      'End If
440                               End If
450                           End If
460                       End If
470                   End If
480               End If

              'If we got here then packet is complete, copy data back to original queue
490           Call .incomingData.CopyBuffer(buffer)
500       End With

Errhandler:
          Dim error  As Long
510       error = Err.Number
520       On Error GoTo 0

          'Destroy auxiliar buffer
530       Set buffer = Nothing

540       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "Walk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 13/01/2010 (ZaMa)
      '11/19/09 Pato - Now the class bandit can walk hidden.
      '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          Dim dummy  As Long
          Dim TempTick As Long
          Dim Heading As eHeading

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

70            Heading = .incomingData.ReadByte()
          
              'Prevent SpeedHack
80            If .flags.TimesWalk >= 30 Then
90                TempTick = GetTickCount And &H7FFFFFFF
100               dummy = (TempTick - .flags.StartWalk)

                  ' 5800 is actually less than what would be needed in perfect conditions to take 30 steps
                  '(it's about 193 ms per step against the over 200 needed in perfect conditions)
110               If dummy < 5800 Then
120                   If TempTick - .flags.CountSH > 30000 Then
130                       .flags.CountSH = 0
140                   End If

150                   If .flags.Montando Then
160                       If TempTick - .flags.CountSH < 45000 Then
170                           .flags.CountSH = 0
180                       End If
190                   End If

200                   If Not .flags.CountSH = 0 Then
210                       If dummy <> 0 Then _
                             dummy = 126000 \ dummy

220                       Call LogHackAttemp("Tramposo SH: " & .Name & " , " & dummy)
                          Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha sido echado por el servidor por posible uso de SH.", FontTypeNames.FONTTYPE_SERVER))
                          Call CloseSocket(userindex)

230                       Exit Sub
240                   Else
250                       .flags.CountSH = TempTick
260                   End If
270               End If
280               .flags.StartWalk = TempTick
290               .flags.TimesWalk = 0
300           End If

310           .flags.TimesWalk = .flags.TimesWalk + 1

              'If exiting, cancel
320           Call CancelExit(userindex)

              'TODO: Debería decirle por consola que no puede?
              'Esta usando el /HOGAR, no se puede mover
330           If .flags.Traveling = 1 Then Exit Sub

340           If .flags.Paralizado = 0 Then
350               If .flags.Meditando Then
                      'Stop meditating, next action will start movement.
360                   .flags.Meditando = False
370                   .Char.FX = 0
380                   .Char.loops = 0

390                   Call WriteConsoleMsg(userindex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)

400                   Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
410               End If
                  
                    .CounterAC = .CounterAC + 1
                    If .CounterAC > 10 Then
                        .ClaveAC = 0
                        Dim xxx As Long
                        For xxx = 1 To 7
                            .ClaveAC = .ClaveAC + RandomNumber(1, 35)
                        Next xxx
                        Call writeClave(userindex)
                        .CounterAC = 0
                    Else
                        
                    End If
                  'Move user
420               Call MoveUserChar(userindex, Heading)

                  'Stop resting if needed
430               If .flags.Descansar Then
440                   .flags.Descansar = False

450                   Call WriteRestOK(userindex)
460                   Call WriteConsoleMsg(userindex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
470               End If
480           Else    'paralized
490               If Not .flags.UltimoMensaje = 1 Then
500                   .flags.UltimoMensaje = 1

510                   Call WriteConsoleMsg(userindex, "No puedes moverte porque estás paralizado.", FontTypeNames.FONTTYPE_INFO)
520               End If

530               .flags.CountSH = 0
540           End If

              'Can't move while hidden except he is a thief
550           If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
560               If .clase <> eClass.Thief Then
570                   .flags.Oculto = 0
580                   .Counters.TiempoOculto = 0

590                   If .flags.Navegando = 1 Then
600                       If .clase = eClass.Pirat Then
                              ' Pierde la apariencia de fragata fantasmal
610                           Call ToggleBoatBody(userindex)

620                           Call WriteConsoleMsg(userindex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
630                           Call ChangeUserChar(userindex, .Char.body, .Char.Head, .Char.Heading, NingunArma, _
                                                  NingunEscudo, NingunCasco)
640                       End If
650                   Else
                          'If not under a spell effect, show char
660                       If .flags.invisible = 0 Then
670                           Call WriteConsoleMsg(userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
680                           Call UsUaRiOs.SetInvisible(userindex, .Char.CharIndex, False)
690                       End If
700                   End If
710               End If
720           End If
730       End With
End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
   On Error GoTo HandleRequestPositionUpdate_Error

10        UserList(userindex).incomingData.ReadByte

20        Call WritePosUpdate(userindex)

   On Error GoTo 0
   Exit Sub

HandleRequestPositionUpdate_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleRequestPositionUpdate of Módulo Protocol in line " & Erl
End Sub
Private Sub HandleAttack(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 13/01/2010
      'Last Modified By: ZaMa
      '10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo.
      '13/11/2009: ZaMa - Se cancela el estado no atacable al atcar.
      '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
      '***************************************************

10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'If dead, can't attack
30            If .flags.Muerto = 1 Then
40                Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
50                Exit Sub
60            End If

              'If user meditates, can't attack
70            If .flags.Meditando Then
80                Exit Sub
90            End If

100           If .flags.ModoCombate = False Then
110               WriteConsoleMsg userindex, "Necesitas estar en Modo Combate para atacar", FontTypeNames.FONTTYPE_INFO
120               Exit Sub
130           End If

              'If equiped weapon is ranged, can't attack this way
140           If .Invent.WeaponEqpObjIndex > 0 Then
150               If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
160                   Call WriteConsoleMsg(userindex, "No puedes usar así este arma.", FontTypeNames.FONTTYPE_INFO)
170                   Exit Sub
180               End If
190           End If

              'If exiting, cancel
200           Call CancelExit(userindex)
              
210           If (Mod_AntiCheat.PuedoPegar(userindex) = False) Then Exit Sub
              
              'Attack!
220           Call UsuarioAtaca(userindex)

              'Now you can be atacked
230           .flags.NoPuedeSerAtacado = False

              'I see you...
240           If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
250               .flags.Oculto = 0
260               .Counters.TiempoOculto = 0

270               If .flags.Navegando = 1 Then
280                   If .clase = eClass.Pirat Then
                          ' Pierde la apariencia de fragata fantasmal
290                       Call ToggleBoatBody(userindex)
300                       Call WriteConsoleMsg(userindex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
310                       Call ChangeUserChar(userindex, .Char.body, .Char.Head, .Char.Heading, NingunArma, _
                                              NingunEscudo, NingunCasco)
320                   End If
330               Else
340                   If .flags.invisible = 0 Then
350                       Call UsUaRiOs.SetInvisible(userindex, .Char.CharIndex, False)
360                       Call WriteConsoleMsg(userindex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
370                   End If
380               End If
390           End If
400       End With
End Sub


Private Sub HandlePickUp(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 07/25/09
      '02/26/2006: Marco - Agregué un checkeo por si el usuario trata de agarrar un item mientras comercia.
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'If dead, it can't pick up objects
30            If .flags.Muerto = 1 Then Exit Sub

              'If user is trading items and attempts to pickup an item, he's cheating, so we kick him.
40            If .flags.Comerciando Then Exit Sub

              'Lower rank administrators can't pick up items

50            Call GetObj(userindex)
60        End With
End Sub
Private Sub HanldeCombatModeToggle(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.ModoCombate Then
40                Call WriteConsoleMsg(userindex, "Has salido del modo combate.", FontTypeNames.FONTTYPE_INFO)
50            Else
60                Call WriteConsoleMsg(userindex, "Has pasado al modo combate.", FontTypeNames.FONTTYPE_INFO)
70            End If

80            .flags.ModoCombate = Not .flags.ModoCombate
90        End With
End Sub

''
' Handles the "SafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSafeToggle(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Seguro Then
40                Call WriteMultiMessage(userindex, eMessages.SafeModeOff)    'Call WriteSafeModeOff(UserIndex)
50            Else
60                Call WriteMultiMessage(userindex, eMessages.SafeModeOn)    'Call WriteSafeModeOn(UserIndex)
70            End If

80            .flags.Seguro = Not .flags.Seguro
90        End With
End Sub

''
' Handles the "ResuscitationSafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResuscitationToggle(ByVal userindex As Integer)
      '***************************************************
      'Author: Rapsodius
      'Creation Date: 10/10/07
      '***************************************************
10        With UserList(userindex)
20            Call .incomingData.ReadByte

30            .flags.SeguroResu = Not .flags.SeguroResu

40            If .flags.SeguroResu Then
50                Call WriteMultiMessage(userindex, eMessages.ResuscitationSafeOn)    'Call WriteResuscitationSafeOn(UserIndex)
60            Else
70                Call WriteMultiMessage(userindex, eMessages.ResuscitationSafeOff)    'Call WriteResuscitationSafeOff(UserIndex)
80            End If
90        End With
End Sub

''
' Handles the "RequestGuildLeaderInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestGuildLeaderInfo(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        UserList(userindex).incomingData.ReadByte

20        Call modGuilds.SendGuildLeaderInfo(userindex)
End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAtributes(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call UserList(userindex).incomingData.ReadByte

20        Call WriteAttributes(userindex)
End Sub

''
' Handles the "RequestFame" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestFame(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call UserList(userindex).incomingData.ReadByte

20        Call EnviarFama(userindex)
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call UserList(userindex).incomingData.ReadByte

20        Call WriteSendSkills(userindex)
End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMiniStats(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call UserList(userindex).incomingData.ReadByte

20        Call WriteMiniStats(userindex)
End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call UserList(userindex).incomingData.ReadByte

          'User quits commerce mode
20        UserList(userindex).flags.Comerciando = False
30        Call WriteCommerceEnd(userindex)
End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 11/03/2010
      '11/03/2010: ZaMa - Le avisa por consola al que cencela que dejo de comerciar.
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'Quits commerce mode with user
30            If .ComUsu.DestUsu > 0 Then
40                If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = userindex Then
50                    Call WriteConsoleMsg(.ComUsu.DestUsu, .Name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_GUILD)
60                    Call FinComerciarUsu(.ComUsu.DestUsu)

                      'Send data in the outgoing buffer of the other user
70                    Call FlushBuffer(.ComUsu.DestUsu)
80                End If
90            End If

100           Call FinComerciarUsu(userindex)
110           Call WriteConsoleMsg(userindex, "Has dejado de comerciar.", FontTypeNames.FONTTYPE_GUILD)
120       End With

End Sub

''
' Handles the "UserCommerceConfirm" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUserCommerceConfirm(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 14/12/2009
      '
      '***************************************************

      'Remove packet ID
10        Call UserList(userindex).incomingData.ReadByte

          'Validate the commerce
20        If PuedeSeguirComerciando(userindex) Then
              'Tell the other user the confirmation of the offer
30            Call WriteUserOfferConfirm(UserList(userindex).ComUsu.DestUsu)
40            UserList(userindex).ComUsu.Confirmo = True
50        End If

End Sub

Private Sub HandleCommerceChat(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 03/12/2009
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)

              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim chat As String

90            chat = buffer.ReadASCIIString()

100           If LenB(chat) <> 0 Then
110               If PuedeSeguirComerciando(userindex) Then
                      'Analize chat...
120                   Call Statistics.ParseChat(chat)

130                   chat = UserList(userindex).Name & "> " & chat
140                   Call WriteCommerceChat(userindex, chat, FontTypeNames.FONTTYPE_PARTY)
150                   Call WriteCommerceChat(UserList(userindex).ComUsu.DestUsu, chat, FontTypeNames.FONTTYPE_PARTY)
160               End If
170           End If

              'If we got here then packet is complete, copy data back to original queue
180           Call .incomingData.CopyBuffer(buffer)
190       End With

Errhandler:
          Dim error  As Long
200       error = Err.Number
210       On Error GoTo 0

          'Destroy auxiliar buffer
220       Set buffer = Nothing

230       If error <> 0 Then _
             Err.Raise error
End Sub


''
' Handles the "BankEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'User exits banking mode
30            .flags.Comerciando = False
40            Call WriteBankEnd(userindex)
50        End With
End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOk(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call UserList(userindex).incomingData.ReadByte

20        If UserList(userindex).ComUsu.Confirmo = False Then Exit Sub

          'Trade accepted
30        Call AceptarComercioUsu(userindex)
End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
          Dim otherUser As Integer

10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            otherUser = .ComUsu.DestUsu

              'Offer rejected
40            If otherUser > 0 Then
50                If UserList(otherUser).flags.UserLogged Then
60                    Call WriteConsoleMsg(otherUser, .Name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_GUILD)
70                    Call FinComerciarUsu(otherUser)

                      'Send data in the outgoing buffer of the other user
80                    Call FlushBuffer(otherUser)
90                End If
100           End If

110           Call WriteConsoleMsg(userindex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_GUILD)
120           Call FinComerciarUsu(userindex)
130       End With
End Sub

''
' Handles the "Drop" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 07/25/09
      '07/25/09: Marco - Agregué un checkeo para patear a los usuarios que tiran items mientras comercian.
      '***************************************************
10        If UserList(userindex).incomingData.length < 4 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          Dim Slot As Byte, Amount As Integer

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

70            Slot = .incomingData.ReadByte()
80            Amount = .incomingData.ReadInteger()
              
              'low rank admins can't drop item. Neither can the dead nor those sailing.
90            If .flags.Navegando = 1 Or _
                 .flags.Montando = 1 Or _
                 .flags.Muerto = 1 Then Exit Sub
              ' ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0)

100           If Amount > 10000 Then Amount = 10000


110           If Slot = FLAGORO Then
120               If Amount <= 0 Or .Stats.Gld < Amount Then
                      'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Anti Cheat > El usuario " & UserList(Userindex).Name & " está intentado dupear oro (Drop).", FontTypeNames.FONTTYPE_ADMIN))
                      'Call LogAntiCheat(UserList(Userindex).Name & " intentó dupear oro.)")
130                   Exit Sub
140               End If
              
150           Else
160               If Amount <= 0 Or Amount > UserList(userindex).Invent.Object(Slot).Amount Then
                      'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Anti Cheat > El usuario " & UserList(Userindex).Name & " está intentado tirar oro dupeado.", FontTypeNames.FONTTYPE_ADMIN))
                     ' Call LogAntiCheat(UserList(Userindex).Name & " intentó dupear oro.)")
170                   Exit Sub
180               End If
190           End If

              'If the user is trading, he can't drop items => He's cheating, we kick him.
200           If .flags.Comerciando Then Exit Sub

              'Are we dropping gold or other items??
210           If Slot = FLAGORO Then
220               If Amount > 10000 Then Exit Sub    'Don't drop too much gold
                  'Call TirarOro(Amount, UserIndex)

230               Call WriteUpdateGold(userindex)
240           Else
                  'Only drop valid slots
250               If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
260                   If .Invent.Object(Slot).ObjIndex = 0 Then
270                       Exit Sub
280                   End If
290                   Call DropObj(userindex, Slot, Amount, .Pos.map, .Pos.X, .Pos.Y)
300               End If
310           End If

320       End With
End Sub

''
' Handles the "CastSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCastSpell(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '13/11/2009: ZaMa - Ahora los npcs pueden atacar al usuario si quizo castear un hechizo
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim spell As Byte
                Dim nclave As Byte
70            spell = .incomingData.ReadByte()
                
                nclave = .incomingData.ReadByte
                If .ClaveAC <> nclave Then Exit Sub
                
80            If .flags.Muerto = 1 Then
90                Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
100               Exit Sub
110           End If

120           If .flags.MenuCliente <> 255 Then
130               If .flags.MenuCliente <> 1 Then
140                   Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("ANTICHEAT > Vigilar a " & .Name, _
                                                                                     FontTypeNames.FONTTYPE_EJECUCION))
150                   Exit Sub

160               End If

170           End If

              'Now you can be atacked
180           .flags.NoPuedeSerAtacado = False

190           If spell < 1 Then
200               .flags.Hechizo = 0
210               Exit Sub
220           ElseIf spell > MAXUSERHECHIZOS Then
230               .flags.Hechizo = 0
240               Exit Sub
250           End If

260           .flags.Hechizo = .Stats.UserHechizos(spell)
270       End With
End Sub

''
' Handles the "LeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex).incomingData
              'Remove packet ID
60            Call .ReadByte

              Dim X  As Byte
              Dim Y  As Byte

70            X = .ReadByte()
80            Y = .ReadByte()

90            Call LookatTile(userindex, UserList(userindex).Pos.map, X, Y)
100       End With
End Sub

''
' Handles the "DoubleClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDoubleClick(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex).incomingData
              'Remove packet ID
60            Call .ReadByte

              Dim X  As Byte
              Dim Y  As Byte

70            X = .ReadByte()
80            Y = .ReadByte()

90            Call Accion(userindex, UserList(userindex).Pos.map, X, Y)
100       End With
End Sub

''
' Handles the "Work" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWork(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 13/01/2010 (ZaMa)
      '13/01/2010: ZaMa - El pirata se puede ocultar en barca
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim Skill As eSkill

70            Skill = .incomingData.ReadByte()

80            If UserList(userindex).flags.Muerto = 1 Then Exit Sub

              'If exiting, cancel
90            Call CancelExit(userindex)

100           Select Case Skill

              Case Robar, Magia, Domar
110               Call WriteMultiMessage(userindex, eMessages.WorkRequestTarget, Skill)

120           Case Ocultarse

                  ' Verifico si se peude ocultar en este mapa
130               If MapInfo(.Pos.map).OcultarSinEfecto = 1 Then
140                   Call WriteConsoleMsg(userindex, "¡Ocultarse no funciona aquí!", FontTypeNames.FONTTYPE_INFO)
150                   Exit Sub
160               End If

170               If .flags.EnConsulta Then
180                   Call WriteConsoleMsg(userindex, "No puedes ocultarte si estás en consulta.", FontTypeNames.FONTTYPE_INFO)
190                   Exit Sub
200               End If
                  
210               If .flags.SlotReto > 0 Then
220                   WriteConsoleMsg userindex, "No puedes ocultarte en reto.", FontTypeNames.FONTTYPE_INFO
230                   Exit Sub
240               End If
                  
250               If .flags.SlotEvent > 0 Then
260                   WriteConsoleMsg userindex, "No puedes ocultarte en evento.", FontTypeNames.FONTTYPE_INFO
270                   Exit Sub
280               End If

290               If .flags.Navegando = 1 Then
300                   If .clase <> eClass.Pirat Then
                          '[CDT 17-02-2004]
310                       If Not .flags.UltimoMensaje = 3 Then
320                           Call WriteConsoleMsg(userindex, "No puedes ocultarte si estás navegando.", FontTypeNames.FONTTYPE_INFO)
330                           .flags.UltimoMensaje = 3
340                       End If
                          '[/CDT]
350                       Exit Sub
360                   End If
370               End If


380               If .flags.Montando = 1 Then
                      '[CDT 17-02-2004]
390                   If Not .flags.UltimoMensaje = 3 Then
400                       Call WriteConsoleMsg(userindex, "No puedes ocultarte si estás montando.", FontTypeNames.FONTTYPE_INFO)
410                       .flags.UltimoMensaje = 3
420                   End If
                      '[/CDT]
430                   Exit Sub
440               End If

450               If .flags.Oculto = 1 Then
                      '[CDT 17-02-2004]
460                   If Not .flags.UltimoMensaje = 2 Then
470                       Call WriteConsoleMsg(userindex, "Ya estás oculto.", FontTypeNames.FONTTYPE_INFO)
480                       .flags.UltimoMensaje = 2
490                   End If
                      '[/CDT]
500                   Exit Sub
510               End If

520               Call DoOcultarse(userindex)

530           End Select

540       End With
End Sub


''
' Handles the "InitCrafting" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInitCrafting(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 29/01/2010
      '
      '***************************************************
          Dim TotalItems As Long
          Dim ItemsPorCiclo As Integer
          
10        If UserList(userindex).incomingData.length < 7 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

70            TotalItems = .incomingData.ReadLong
80            ItemsPorCiclo = .incomingData.ReadInteger

90            If TotalItems > 0 Then

100               .Construir.Cantidad = TotalItems
110               .Construir.PorCiclo = MinimoInt(MaxItemsConstruibles(userindex), ItemsPorCiclo)

120           End If
130       End With
End Sub

''
' Handles the "UseSpellMacro" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseSpellMacro(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            Call SendData(SendTarget.ToAdmins, userindex, PrepareMessageConsoleMsg(.Name & " fue expulsado por Anti-macro de hechizos.", FontTypeNames.FONTTYPE_VENENO))
40            Call WriteErrorMsg(userindex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros.")
50            Call FlushBuffer(userindex)
60            Call CloseSocket(userindex)
70        End With
End Sub

Private Sub HandleSetMenu(ByVal userindex As Integer)
          
10        If UserList(userindex).incomingData.length < 10 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50        With UserList(userindex)
60            Call .incomingData.ReadByte

              '1 spell
              '2 inventario

70            .flags.MenuCliente = .incomingData.ReadByte
80            .flags.LastSlotClient = .incomingData.ReadByte
            
            
              Dim Value(1) As Long
              Value(0) = .incomingData.ReadLong
              Value(1) = .incomingData.ReadLong
              UpdatePointer userindex, .flags.MenuCliente, Value(0), Value(1)
              
              'WriteConsoleMsg UserIndex, "X: " & Value(0) & ", Y: " & Value(1), FontTypeNames.FONTTYPE_INFO
90        End With

End Sub

''
' Handles the "CraftBlacksmith" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftBlacksmith(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex).incomingData
              'Remove packet ID
60            Call .ReadByte

              Dim Item As Integer

70            Item = .ReadInteger()

80            If Item < 1 Then Exit Sub

90            If ObjData(Item).SkHerreria = 0 Then Exit Sub

100           Call HerreroConstruirItem(userindex, Item)
110       End With
End Sub

''
' Handles the "CraftCarpenter" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftCarpenter(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex).incomingData
              'Remove packet ID
60            Call .ReadByte

              Dim Item As Integer

70            Item = .ReadInteger()

80            If Item < 1 Then Exit Sub

90            If ObjData(Item).SkCarpinteria = 0 Then Exit Sub

100           Call CarpinteroConstruirItem(userindex, Item)
110       End With
End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 14/01/2010 (ZaMa)
      '16/11/2009: ZaMa - Agregada la posibilidad de extraer madera elfica.
      '12/01/2010: ZaMa - Ahora se admiten armas arrojadizas (proyectiles sin municiones).
      '14/01/2010: ZaMa - Ya no se pierden municiones al atacar npcs con dueño.
      '***************************************************
10        If UserList(userindex).incomingData.length < 4 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim X  As Byte
              Dim Y  As Byte
              Dim Skill As eSkill
              Dim DummyInt As Integer
              Dim tU As Integer   'Target user
              Dim tN As Integer   'Target NPC

70            X = .incomingData.ReadByte()
80            Y = .incomingData.ReadByte()

90            Skill = .incomingData.ReadByte()

100           If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando _
                 Or Not InMapBounds(.Pos.map, X, Y) Then Exit Sub

110           If Not InRangoVision(userindex, X, Y) Then
120               Call WritePosUpdate(userindex)
130               Exit Sub
140           End If

              'If exiting, cancel
150           Call CancelExit(userindex)

160           Select Case Skill
              Case eSkill.Proyectiles

                  'Check attack interval
170               If Not IntervaloPermiteAtacar(userindex, False) Then Exit Sub
                  'Check Magic interval
180               If Not IntervaloPermiteLanzarSpell(userindex, False) Then Exit Sub
                  'Check bow's interval
190               If Not IntervaloPermiteUsarArcos(userindex) Then Exit Sub

                  Dim Atacked As Boolean
200               Atacked = True

                  'Make sure the item is valid and there is ammo equipped.
210               With .Invent
                      ' Tiene arma equipada?
220                   If .WeaponEqpObjIndex = 0 Then
230                       DummyInt = 1
                          ' En un slot válido?
240                   ElseIf .WeaponEqpSlot < 1 Or .WeaponEqpSlot > UserList(userindex).CurrentInventorySlots Then
250                       DummyInt = 1
                          ' Usa munición? (Si no la usa, puede ser un arma arrojadiza)
260                   ElseIf ObjData(.WeaponEqpObjIndex).Municion = 1 Then
                          ' La municion esta equipada en un slot valido?
270                       If .MunicionEqpSlot < 1 Or .MunicionEqpSlot > UserList(userindex).CurrentInventorySlots Then
280                           DummyInt = 1
                              ' Tiene munición?
290                       ElseIf .MunicionEqpObjIndex = 0 Then
300                           DummyInt = 1
                              ' Son flechas?
310                       ElseIf ObjData(.MunicionEqpObjIndex).ObjType <> eOBJType.otFlechas Then
320                           DummyInt = 1
                              ' Tiene suficientes?
330                       ElseIf .Object(.MunicionEqpSlot).Amount < 1 Then
340                           DummyInt = 1
350                       End If
                          ' Es un arma de proyectiles?
360                   ElseIf ObjData(.WeaponEqpObjIndex).proyectil <> 1 Then
370                       DummyInt = 2
380                   End If

390                   If DummyInt <> 0 Then
400                       If DummyInt = 1 Then
410                           Call WriteConsoleMsg(userindex, "No tienes municiones.", FontTypeNames.FONTTYPE_INFO)

420                           Call Desequipar(userindex, .WeaponEqpSlot)
430                       End If

440                       Call Desequipar(userindex, .MunicionEqpSlot)
450                       Exit Sub
460                   End If
470               End With

                  'Quitamos stamina
480               If .Stats.MinSta >= 10 Then
490                   Call QuitarSta(userindex, RandomNumber(1, 10))
500               Else
510                   If .Genero = eGenero.Hombre Then
520                       Call WriteConsoleMsg(userindex, "Estás muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
530                   Else
540                       Call WriteConsoleMsg(userindex, "Estás muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)
550                   End If
560                   Exit Sub
570               End If

580               SendData SendTarget.ToPCArea, userindex, PrepareMessageMovimientSW(.Char.CharIndex, 1)

590               Call LookatTile(userindex, .Pos.map, X, Y)

600               tU = .flags.TargetUser
610               tN = .flags.TargetNPC

                  'Validate target
620               If tU > 0 Then
                      'Only allow to atack if the other one can retaliate (can see us)
630                   If Abs(UserList(tU).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
640                       Call WriteConsoleMsg(userindex, "Estás demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
650                       Exit Sub
660                   End If

                      'Prevent from hitting self
670                   If tU = userindex Then
680                       Call WriteConsoleMsg(userindex, "¡No puedes atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO)
690                       Exit Sub
700                   End If

                      'Attack!
710                   Atacked = UsuarioAtacaUsuario(userindex, tU)


720               ElseIf tN > 0 Then
                      'Only allow to atack if the other one can retaliate (can see us)
730                   If Abs(Npclist(tN).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(Npclist(tN).Pos.X - .Pos.X) > RANGO_VISION_X Then
740                       Call WriteConsoleMsg(userindex, "Estás demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
750                       Exit Sub
760                   End If

                      'Is it attackable???
770                   If Npclist(tN).Attackable <> 0 Then

                          'Attack!
780                       Atacked = UsuarioAtacaNpc(userindex, tN)
790                   End If
800               End If

                  ' Solo pierde la munición si pudo atacar al target, o tiro al aire
810               If Atacked Then
820                   With .Invent
                          ' Tiene equipado arco y flecha?
830                       If ObjData(.WeaponEqpObjIndex).Municion = 1 Then
840                           DummyInt = .MunicionEqpSlot


                              'Take 1 arrow away - we do it AFTER hitting, since if Ammo Slot is 0 it gives a rt9 and kicks players
850                           Call QuitarUserInvItem(userindex, DummyInt, 1)

860                           If .Object(DummyInt).Amount > 0 Then
                                  'QuitarUserInvItem unequips the ammo, so we equip it again
870                               .MunicionEqpSlot = DummyInt
880                               .MunicionEqpObjIndex = .Object(DummyInt).ObjIndex
890                               .Object(DummyInt).Equipped = 1
900                           Else
910                               .MunicionEqpSlot = 0
920                               .MunicionEqpObjIndex = 0
930                           End If
                              ' Tiene equipado un arma arrojadiza
940                       Else
950                           DummyInt = .WeaponEqpSlot

                              'Take 1 knife away
960                           Call QuitarUserInvItem(userindex, DummyInt, 1)

970                           If .Object(DummyInt).Amount > 0 Then
                                  'QuitarUserInvItem unequips the weapon, so we equip it again
980                               .WeaponEqpSlot = DummyInt
990                               .WeaponEqpObjIndex = .Object(DummyInt).ObjIndex
1000                              .Object(DummyInt).Equipped = 1
1010                          Else
1020                              .WeaponEqpSlot = 0
1030                              .WeaponEqpObjIndex = 0
1040                          End If

1050                      End If

1060                      Call UpdateUserInv(False, userindex, DummyInt)
1070                  End With
1080              End If

1090          Case eSkill.Magia
                  'Check the map allows spells to be casted.
1100              If MapInfo(.Pos.map).MagiaSinEfecto > 0 Then
1110                  Call WriteConsoleMsg(userindex, "Una fuerza oscura te impide canalizar tu energía.", FontTypeNames.FONTTYPE_FIGHT)
1120                  Exit Sub
1130              End If

                  'Target whatever is in that tile
1140              Call LookatTile(userindex, .Pos.map, X, Y)

                  'If it's outside range log it and exit
1150              If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
1160                  Call LogCheating("Ataque fuera de rango de " & .Name & "(" & .Pos.map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .ip & " a la posición (" & .Pos.map & "/" & X & "/" & Y & ")")
1170                  Exit Sub
1180              End If

                  'Check bow's interval
1190              If Not IntervaloPermiteUsarArcos(userindex, False) Then Exit Sub

1200              If (Mod_AntiCheat.PuedoCasteoHechizo(userindex) = False) Then Exit Sub
                  
                  'Check Spell-Hit interval
1210              If Not IntervaloPermiteGolpeMagia(userindex) Then
                      'Check Magic interval
1220                  If Not IntervaloPermiteLanzarSpell(userindex) Then
1230                      Exit Sub
1240                  End If
1250              End If


                  'Check intervals and cast
1260              If .flags.Hechizo > 0 Then
1270                  Call LanzarHechizo(.flags.Hechizo, userindex)
1280                  .flags.Hechizo = 0
1290              Else
1300                  Call WriteConsoleMsg(userindex, "¡Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)
1310              End If

1320          Case eSkill.Pesca
1330              DummyInt = .Invent.WeaponEqpObjIndex
1340              If DummyInt = 0 Then Exit Sub

                  'Check interval
1350              If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub

                  'Basado en la idea de Barrin
                  'Comentario por Barrin: jah, "basado", caradura ! ^^
1360              If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 1 Then
1370                  Call WriteConsoleMsg(userindex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
1380                  Exit Sub
1390              End If

1400              If HayAgua(.Pos.map, X, Y) Then
1410                  Select Case DummyInt
                      Case CAÑA_PESCA
1420                      Call DoPescar(userindex)
                      
                      Case CAÑA_COFRES
                          Call DoPescarRed(userindex, True)
                          
1430                  Case RED_PESCA
1440                      If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
1450                          Call WriteConsoleMsg(userindex, "Estás demasiado lejos para pescar.", FontTypeNames.FONTTYPE_INFO)
1460                          Exit Sub
1470                      End If

1480                      Call DoPescarRed(userindex)

1490                  Case Else
1500                      Exit Sub    'Invalid item!
1510                  End Select

                      'Play sound!
1520                  Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
1530              Else
1540                  Call WriteConsoleMsg(userindex, "No hay agua donde pescar. Busca un lago, río o mar.", FontTypeNames.FONTTYPE_INFO)
1550              End If

1560          Case eSkill.Robar
                  'Does the map allow us to steal here?
1570              If MapInfo(.Pos.map).Pk Then

                      'Check interval
1580                  If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub

                      'Target whatever is in that tile
1590                  Call LookatTile(userindex, UserList(userindex).Pos.map, X, Y)

1600                  tU = .flags.TargetUser

1610                  If tU > 0 And tU <> userindex Then
                          'Can't steal administrative players
1620                      If UserList(tU).flags.Privilegios And PlayerType.User Then
1630                          If UserList(tU).flags.Muerto = 0 Then
1640                              If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 1 Then
                                      'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
1650                                  Call WriteShortMsj(userindex, 8, FontTypeNames.FONTTYPE_INFO)
1660                                  Exit Sub
1670                              End If

                                  '17/09/02
                                  'Check the trigger
1680                              If MapData(UserList(tU).Pos.map, X, Y).trigger = eTrigger.ZONASEGURA Then
1690                                  Call WriteConsoleMsg(userindex, "No puedes robar aquí.", FontTypeNames.FONTTYPE_WARNING)
1700                                  Exit Sub
1710                              End If

1720                              If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
1730                                  Call WriteConsoleMsg(userindex, "No puedes robar aquí.", FontTypeNames.FONTTYPE_WARNING)
1740                                  Exit Sub
1750                              End If

1760                              Call DoRobar(userindex, tU)
1770                          End If
1780                      End If
1790                  Else
1800                      Call WriteConsoleMsg(userindex, "¡No hay a quien robarle!", FontTypeNames.FONTTYPE_INFO)
1810                  End If
1820              Else
1830                  Call WriteConsoleMsg(userindex, "¡No puedes robar en zonas seguras!", FontTypeNames.FONTTYPE_INFO)
1840              End If

1850          Case eSkill.talar
                  'Check interval
1860              If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub

1870              If .Invent.WeaponEqpObjIndex = 0 Then
1880                  Call WriteConsoleMsg(userindex, "Deberías equiparte el hacha.", FontTypeNames.FONTTYPE_INFO)
1890                  Exit Sub
1900              End If

1910              If .Invent.WeaponEqpObjIndex <> HACHA_LEÑADOR And _
                     .Invent.WeaponEqpObjIndex <> HACHA_DORADA Then
                      ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
1920                  Exit Sub
1930              End If

1940              DummyInt = MapData(.Pos.map, X, Y).ObjInfo.ObjIndex

1950              If DummyInt > 0 Then
1960                  If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                          'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
1970                      Call WriteShortMsj(userindex, 8, FontTypeNames.FONTTYPE_INFO)
1980                      Exit Sub
1990                  End If

                      'Barrin 29/9/03
2000                  If .Pos.X = X And .Pos.Y = Y Then
2010                      Call WriteConsoleMsg(userindex, "No puedes talar desde allí.", FontTypeNames.FONTTYPE_INFO)
2020                      Exit Sub
2030                  End If

2040                  ArbT = DummyInt
                      '¿Hay un arbol donde clickeo?
2050                  If ObjData(DummyInt).ObjType = eOBJType.otarboles Then
2060                      Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
2070                      Call DoTalar(userindex)
2080                  ElseIf ObjData(DummyInt).ObjType = 38 Then
2090                      SendData SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y)

2100                      DoTalar userindex
2110                  End If
2120              Else
2130                  Call WriteConsoleMsg(userindex, "No hay ningún árbol ahí.", FontTypeNames.FONTTYPE_INFO)
2140              End If

2150          Case eSkill.Mineria
2160              If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub

2170              If .Invent.WeaponEqpObjIndex = 0 Then Exit Sub

2180              If Not ((.Invent.WeaponEqpObjIndex <> PIQUETE_MINERO) Or (.Invent.WeaponEqpObjIndex <> PIQUETE_ORO)) Then
                      ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
2190                  Exit Sub
2200              End If

                  'Target whatever is in the tile
2210              Call LookatTile(userindex, .Pos.map, X, Y)

2220              DummyInt = MapData(.Pos.map, X, Y).ObjInfo.ObjIndex

2230              If DummyInt > 0 Then
                      'Check distance
2240                  If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                          'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
2250                      Call WriteShortMsj(userindex, 8, FontTypeNames.FONTTYPE_INFO)
2260                      Exit Sub
2270                  End If

2280                  DummyInt = MapData(.Pos.map, X, Y).ObjInfo.ObjIndex    'CHECK
                      '¿Hay un yacimiento donde clickeo?
2290                  If ObjData(DummyInt).ObjType = eOBJType.otYacimiento Then
2300                      Call DoMineria(userindex)
2310                  Else
2320                      Call WriteConsoleMsg(userindex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
2330                  End If
2340              Else
2350                  Call WriteConsoleMsg(userindex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
2360              End If

2370          Case eSkill.Domar
                  'Modificado 25/11/02
                  'Optimizado y solucionado el bug de la doma de
                  'criaturas hostiles.

                  'Target whatever is that tile
2380              Call LookatTile(userindex, .Pos.map, X, Y)
2390              tN = .flags.TargetNPC

2400              If tN > 0 Then
2410                  If Npclist(tN).flags.Domable > 0 Then
2420                      If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                              'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
2430                          Call WriteShortMsj(userindex, 8, FontTypeNames.FONTTYPE_INFO)
2440                          Exit Sub
2450                      End If

2460                      If LenB(Npclist(tN).flags.AttackedBy) <> 0 Then
2470                          Call WriteConsoleMsg(userindex, "No puedes domar una criatura que está luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
2480                          Exit Sub
2490                      End If

2500                      Call DoDomar(userindex, tN)
2510                  Else
2520                      Call WriteConsoleMsg(userindex, "No puedes domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)
2530                  End If
2540              Else
2550                  Call WriteConsoleMsg(userindex, "¡No hay ninguna criatura allí!", FontTypeNames.FONTTYPE_INFO)
2560              End If

2570          Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
                  'Check interval
2580              If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub

                  'Check there is a proper item there
2590              If .flags.TargetObj > 0 Then
2600                  If ObjData(.flags.TargetObj).ObjType = eOBJType.otFragua Then
                          'Validate other items
2610                      If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > .CurrentInventorySlots Then
2620                          Exit Sub
2630                      End If

                          ''chequeamos que no se zarpe duplicando oro
2640                      If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
2650                          If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).Amount = 0 Then
2660                              Call WriteConsoleMsg(userindex, "No tienes más minerales.", FontTypeNames.FONTTYPE_INFO)
2670                              Exit Sub
2680                          End If

                              ''FUISTE
2690                          Call WriteErrorMsg(userindex, "Has sido expulsado por el sistema anti cheats.")
2700                          Call FlushBuffer(userindex)
2710                          Call CloseSocket(userindex)
2720                          Exit Sub
2730                      End If
2740                      Call FundirMineral(userindex)
2750                  Else
2760                      Call WriteConsoleMsg(userindex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
2770                  End If
2780              Else
2790                  Call WriteConsoleMsg(userindex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
2800              End If

2810          Case eSkill.herreria
                  'Target wehatever is in that tile
2820              Call LookatTile(userindex, .Pos.map, X, Y)

2830              If .flags.TargetObj > 0 Then
2840                  If ObjData(.flags.TargetObj).ObjType = eOBJType.otYunque Then
2850                      Call EnivarArmasConstruibles(userindex)
2860                      Call EnivarArmadurasConstruibles(userindex)
2870                      Call WriteShowBlacksmithForm(userindex)
2880                  Else
2890                      Call WriteConsoleMsg(userindex, "Ahí no hay ningún yunque.", FontTypeNames.FONTTYPE_INFO)
2900                  End If
2910              Else
2920                  Call WriteConsoleMsg(userindex, "Ahí no hay ningún yunque.", FontTypeNames.FONTTYPE_INFO)
2930              End If
2940          End Select
2950      End With
End Sub

''
' Handles the "CreateNewGuild" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateNewGuild(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/11/09
      '05/11/09: Pato - Ahora se quitan los espacios del principio y del fin del nombre del clan
      '***************************************************
10        If UserList(userindex).incomingData.length < 9 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim desc As String
              Dim GuildName As String
              Dim site As String
              Dim codex() As String
              Dim errorStr As String

90            desc = buffer.ReadASCIIString()
100           GuildName = Trim$(buffer.ReadASCIIString())
110           site = buffer.ReadASCIIString()
120           codex = Split(buffer.ReadASCIIString(), SEPARATOR)

130           If modGuilds.CrearNuevoClan(userindex, desc, GuildName, site, codex, .FundandoGuildAlineacion, errorStr) Then
140               Call SendData(SendTarget.ToAll, userindex, PrepareMessageConsoleMsg(.Name & " fundó el clan " & GuildName & " de alineación " & modGuilds.GuildAlignment(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
150               .Stats.Gld = .Stats.Gld - 25000000
160               WriteUpdateGold userindex
170               Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))


                  'Update tag
180               Call RefreshCharStatus(userindex)
190           Else
200               Call WriteConsoleMsg(userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
210           End If

              'If we got here then packet is complete, copy data back to original queue
220           Call .incomingData.CopyBuffer(buffer)
230       End With

Errhandler:
          Dim error  As Long
240       error = Err.Number
250       On Error GoTo 0

          'Destroy auxiliar buffer
260       Set buffer = Nothing

270       If error <> 0 Then _
             Err.Raise error
End Sub
''
' Handles the "SpellInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpellInfo(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim spellSlot As Byte
              Dim spell As Integer

70            spellSlot = .incomingData.ReadByte()

              'Validate slot
80            If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
90                Call WriteConsoleMsg(userindex, "¡Primero selecciona el hechizo!", FontTypeNames.FONTTYPE_INFO)
100               Exit Sub
110           End If

              'Validate spell in the slot
120           spell = .Stats.UserHechizos(spellSlot)
130           If spell > 0 And spell < NumeroHechizos + 1 Then
140               With Hechizos(spell)
                      'Send information
150                   Call WriteConsoleMsg(userindex, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & vbCrLf _
                                                      & "Nombre:" & .Nombre & vbCrLf _
                                                      & "Descripción:" & .desc & vbCrLf _
                                                      & "Skill requerido: " & .MinSkill & " de magia." & vbCrLf _
                                                      & "Maná necesario: " & .ManaRequerido & vbCrLf _
                                                      & "Energía necesaria: " & .StaRequerido & vbCrLf _
                                                      & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%", FontTypeNames.FONTTYPE_INFO)
160               End With
170           End If
180       End With


End Sub

''
' Handles the "EquipItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim itemSlot As Byte

70            itemSlot = .incomingData.ReadByte()

              'Dead users can't equip items
80            If .flags.Muerto = 1 Then Exit Sub

              'Validate item slot
90            If itemSlot > .CurrentInventorySlots Or itemSlot < 1 Then Exit Sub

100           If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub

110           Call EquiparInvItem(userindex, itemSlot)
120       End With
End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 06/28/2008
      'Last Modified By: NicoNZ
      ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
      ' 06/28/2008: NicoNZ - Sólo se puede cambiar si está inmovilizado.
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim Heading As eHeading
              Dim posX As Integer
              Dim posY As Integer

70            Heading = .incomingData.ReadByte()

80            If .flags.Paralizado = 1 And .flags.Inmovilizado = 0 Then
90                Select Case Heading
                  Case eHeading.NORTH
100                   posY = -1
110               Case eHeading.EAST
120                   posX = 1
130               Case eHeading.SOUTH
140                   posY = 1
150               Case eHeading.WEST
160                   posX = -1
170               End Select

180               If LegalPos(.Pos.map, .Pos.X + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then
190                   Exit Sub
200               End If
210           End If

              'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)

220           If Heading > 0 And Heading < 5 Then
230               .Char.Heading = Heading

240               SendData SendTarget.ToPCArea, userindex, PrepareMessageChangeHeading(.Char.CharIndex, Heading)
250           End If
260       End With
End Sub

''
' Handles the "ModifySkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleModifySkills(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 11/19/09
      '11/19/09: Pato - Adapting to new skills system.
      '***************************************************
10        If UserList(userindex).incomingData.length < 1 + NUMSKILLS Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim i  As Long
              Dim Count As Integer
              Dim Points(1 To NUMSKILLS) As Byte

              'Codigo para prevenir el hackeo de los skills
              '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
70            For i = 1 To NUMSKILLS
80                Points(i) = .incomingData.ReadByte()

90                If Points(i) < 0 Then
100                   Call LogHackAttemp(.Name & " IP:" & .ip & " trató de hackear los skills.")
110                   .Stats.SkillPts = 0
120                   Call CloseSocket(userindex)
130                   Exit Sub
140               End If

150               Count = Count + Points(i)
160           Next i

170           If Count > .Stats.SkillPts Then
180               Call LogHackAttemp(.Name & " IP:" & .ip & " trató de hackear los skills.")
190               Call CloseSocket(userindex)
200               Exit Sub
210           End If
              '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

220           .Counters.AsignedSkills = MinimoInt(10, .Counters.AsignedSkills + Count)

230           With .Stats
240               For i = 1 To NUMSKILLS
250                   If Points(i) > 0 Then
260                       .SkillPts = .SkillPts - Points(i)
270                       .UserSkills(i) = .UserSkills(i) + Points(i)

                          'Client should prevent this, but just in case...
280                       If .UserSkills(i) > 100 Then
290                           .SkillPts = .SkillPts + .UserSkills(i) - 100
300                           .UserSkills(i) = 100
310                       End If

320                       Call CheckEluSkill(userindex, i, True)
330                   End If
340               Next i
350           End With
360       End With
End Sub

''
' Handles the "Train" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim SpawnedNpc As Integer
              Dim PetIndex As Byte

70            PetIndex = .incomingData.ReadByte()

80            If .flags.TargetNPC = 0 Then Exit Sub

90            If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub

100           If Npclist(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
110               If PetIndex > 0 And PetIndex < Npclist(.flags.TargetNPC).NroCriaturas + 1 Then
                      'Create the creature
120                   SpawnedNpc = SpawnNpc(Npclist(.flags.TargetNPC).Criaturas(PetIndex).NpcIndex, Npclist(.flags.TargetNPC).Pos, True, False)

130                   If SpawnedNpc > 0 Then
140                       Npclist(SpawnedNpc).MaestroNpc = .flags.TargetNPC
150                       Npclist(.flags.TargetNPC).Mascotas = Npclist(.flags.TargetNPC).Mascotas + 1
160                   End If
170               End If
180           Else
190               Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
200           End If
210       End With
End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 4 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim Slot As Byte
              Dim Amount As Integer

70            Slot = .incomingData.ReadByte()
80            Amount = .incomingData.ReadInteger()

              'Dead people can't commerce...
90            If .flags.Muerto = 1 Then
100               Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
110               Exit Sub
120           End If

              '¿El target es un NPC valido?
130           If .flags.TargetNPC < 1 Then Exit Sub

              '¿El NPC puede comerciar?
140           If Npclist(.flags.TargetNPC).Comercia = 0 Then
150               Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
160               Exit Sub
170           End If

              'Only if in commerce mode....
180           If Not .flags.Comerciando Then
190               Call WriteConsoleMsg(userindex, "No estás comerciando.", FontTypeNames.FONTTYPE_INFO)
200               Exit Sub
210           End If

              'User compra el item
220           Call Comercio(eModoComercio.Compra, userindex, .flags.TargetNPC, Slot, Amount)
230       End With
End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 4 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim Slot As Byte
              Dim Amount As Integer

70            Slot = .incomingData.ReadByte()
80            Amount = .incomingData.ReadInteger()

              'Dead people can't commerce
90            If .flags.Muerto = 1 Then
100               Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
110               Exit Sub
120           End If

              '¿El target es un NPC valido?
130           If .flags.TargetNPC < 1 Then Exit Sub

              '¿Es el banquero?
140           If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
150               Exit Sub
160           End If

              'User retira el item del slot
170           Call UserRetiraItem(userindex, Slot, Amount)
180       End With
End Sub

''
' Handles the "CommerceSell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 4 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim Slot As Byte
              Dim Amount As Integer

70            Slot = .incomingData.ReadByte()
80            Amount = .incomingData.ReadInteger()

              'Dead people can't commerce...
90            If .flags.Muerto = 1 Then
100               Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
110               Exit Sub
120           End If

              '¿El target es un NPC valido?
130           If .flags.TargetNPC < 1 Then Exit Sub

              '¿El NPC puede comerciar?
140           If Npclist(.flags.TargetNPC).Comercia = 0 Then
150               Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
160               Exit Sub
170           End If

              'User compra el item del slot
180           Call Comercio(eModoComercio.Venta, userindex, .flags.TargetNPC, Slot, Amount)
190       End With
End Sub

''
' Handles the "BankDeposit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 4 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim Slot As Byte
              Dim Amount As Integer

70            Slot = .incomingData.ReadByte()
80            Amount = .incomingData.ReadInteger()

              'Dead people can't commerce...
90            If .flags.Muerto = 1 Then
100               Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
110               Exit Sub
120           End If

              '¿El target es un NPC valido?
130           If .flags.TargetNPC < 1 Then Exit Sub

              '¿El NPC puede comerciar?
140           If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
150               Exit Sub
160           End If

              'User deposita el item del slot rdata
170           Call UserDepositaItem(userindex, Slot, Amount)
180       End With
End Sub

''
' Handles the "ForumPost" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForumPost(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 02/01/2010
      '02/01/2010: ZaMa - Implemento nuevo sistema de foros
      '***************************************************
10        If UserList(userindex).incomingData.length < 6 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim ForumMsgType As eForumMsgType

              Dim File As String
              Dim Title As String
              Dim Post As String
              Dim ForumIndex As Integer
              Dim postFile As String
              Dim ForumType As Byte

90            ForumMsgType = buffer.ReadByte()

100           Title = buffer.ReadASCIIString()
110           Post = buffer.ReadASCIIString()

120           If .flags.TargetObj > 0 Then
130               ForumType = ForumAlignment(ForumMsgType)

140               Select Case ForumType

                  Case eForumType.ieGeneral
150                   ForumIndex = GetForumIndex(ObjData(.flags.TargetObj).ForoID)

160               Case eForumType.ieREAL
170                   ForumIndex = GetForumIndex(FORO_REAL_ID)

180               Case eForumType.ieCAOS
190                   ForumIndex = GetForumIndex(FORO_CAOS_ID)

200               End Select

210               Call AddPost(ForumIndex, Post, .Name, Title, EsAnuncio(ForumMsgType))
220           End If

              'If we got here then packet is complete, copy data back to original queue
230           Call .incomingData.CopyBuffer(buffer)
240       End With

Errhandler:
          Dim error  As Long
250       error = Err.Number
260       On Error GoTo 0

          'Destroy auxiliar buffer
270       Set buffer = Nothing

280       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "MoveSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex).incomingData
              'Remove packet ID
60            Call .ReadByte

              Dim dir As Integer

70            If .ReadBoolean() Then
80                dir = 1
90            Else
100               dir = -1
110           End If

120           Call DesplazarHechizo(userindex, dir, .ReadByte())
130       End With
End Sub

''
' Handles the "MoveBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveBank(ByVal userindex As Integer)
      '***************************************************
      'Author: Torres Patricio (Pato)
      'Last Modification: 06/14/09
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex).incomingData
              'Remove packet ID
60            Call .ReadByte

              Dim dir As Integer
              Dim Slot As Byte
              Dim TempItem As Obj

70            If .ReadBoolean() Then
80                dir = 1
90            Else
100               dir = -1
110           End If

120           Slot = .ReadByte()
130       End With

140       With UserList(userindex)
150           TempItem.ObjIndex = .BancoInvent.Object(Slot).ObjIndex
160           TempItem.Amount = .BancoInvent.Object(Slot).Amount

170           If dir = 1 Then    'Mover arriba
180               .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot - 1)
190               .BancoInvent.Object(Slot - 1).ObjIndex = TempItem.ObjIndex
200               .BancoInvent.Object(Slot - 1).Amount = TempItem.Amount
210           Else    'mover abajo
220               .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot + 1)
230               .BancoInvent.Object(Slot + 1).ObjIndex = TempItem.ObjIndex
240               .BancoInvent.Object(Slot + 1).Amount = TempItem.Amount
250           End If
260       End With

270       Call UpdateBanUserInv(True, userindex, 0)
280       Call UpdateVentanaBanco(userindex)

End Sub

''
' Handles the "ClanCodexUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleClanCodexUpdate(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim desc As String
              Dim codex() As String

90            desc = buffer.ReadASCIIString()
100           codex = Split(buffer.ReadASCIIString(), SEPARATOR)

110           Call modGuilds.ChangeCodexAndDesc(desc, codex, .GuildIndex)

              'If we got here then packet is complete, copy data back to original queue
120           Call .incomingData.CopyBuffer(buffer)
130       End With

Errhandler:
          Dim error  As Long
140       error = Err.Number
150       On Error GoTo 0

          'Destroy auxiliar buffer
160       Set buffer = Nothing

170       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 24/11/2009
      '24/11/2009: ZaMa - Nuevo sistema de comercio
      '***************************************************
10        If UserList(userindex).incomingData.length < 7 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte
              
              Dim Amount As Long
              Dim Slot As Byte
              Dim tUser As Integer
              Dim OfferSlot As Byte
              Dim ObjIndex As Integer
              
70            Slot = .incomingData.ReadByte()
80            Amount = .incomingData.ReadLong()
90            OfferSlot = .incomingData.ReadByte()
              
              'Get the other player
100           tUser = .ComUsu.DestUsu
              
              If tUser <= 0 Then
                  Call FinComerciarUsu(userindex)
                  Exit Sub
              End If
              
              ' If he's already confirmed his offer, but now tries to change it, then he's cheating
110           If UserList(userindex).ComUsu.Confirmo = True Then
                  
                  ' Finish the trade
120               Call FinComerciarUsu(userindex)
              
130               If tUser <= 0 Or tUser > MaxUsers Then
140                   Call FinComerciarUsu(tUser)
150                   Call Protocol.FlushBuffer(tUser)
160               End If
              
170               Exit Sub
180           End If
              
              'If slot is invalid and it's not gold or it's not 0 (Substracting), then ignore it.
190           If ((Slot < 0 Or Slot > UserList(userindex).CurrentInventorySlots) And Slot <> FLAGORO) Then Exit Sub
              
              'If OfferSlot is invalid, then ignore it.
200           If OfferSlot < 1 Or OfferSlot > MAX_OFFER_SLOTS + 1 Then Exit Sub
              
              ' Can be negative if substracted from the offer, but never 0.
210           If Amount = 0 Then Exit Sub
              
              'Has he got enough??
220           If Slot = FLAGORO Then
                  ' Can't offer more than he has
230               If Amount > .Stats.Gld - .ComUsu.GoldAmount Then
240                   Call WriteCommerceChat(userindex, "No tienes esa cantidad de oro para agregar a la oferta.", FontTypeNames.FONTTYPE_TALK)
250                   Exit Sub
260               End If
                  
270               If Amount < 0 Then
280                   If Abs(Amount) > .ComUsu.GoldAmount Then
290                       Amount = .ComUsu.GoldAmount * (-1)
300                   End If
310               End If
320           Else
                  'If modifing a filled offerSlot, we already got the objIndex, then we don't need to know it
330               If Slot <> 0 Then ObjIndex = .Invent.Object(Slot).ObjIndex
                  ' Can't offer more than he has
340               If Not HasEnoughItems(userindex, ObjIndex, _
                      TotalOfferItems(ObjIndex, userindex) + Amount) Then
                      
350                   Call WriteCommerceChat(userindex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
360                   Exit Sub
370               End If
                  
380               If Amount < 0 Then
390                   If Abs(Amount) > .ComUsu.cant(OfferSlot) Then
400                       Amount = .ComUsu.cant(OfferSlot) * (-1)
410                   End If
420               End If
              
430               If ItemNewbie(ObjIndex) Then
440                   Call WriteCancelOfferItem(userindex, OfferSlot)
450                   Exit Sub
460               End If
                  
                  'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
470               If .flags.Navegando = 1 Then
480                   If .Invent.BarcoSlot = Slot Then
490                       Call WriteCommerceChat(userindex, "No puedes vender tu barco mientras lo estés usando.", FontTypeNames.FONTTYPE_TALK)
500                       Exit Sub
510                   End If
520               End If
                  
530               If .Invent.MochilaEqpSlot > 0 Then
540                   If .Invent.MochilaEqpSlot = Slot Then
550                       Call WriteCommerceChat(userindex, "No puedes vender tu mochila mientras la estés usando.", FontTypeNames.FONTTYPE_TALK)
560                       Exit Sub
570                   End If
580               End If
590           End If
              
600           Call AgregarOferta(userindex, OfferSlot, ObjIndex, Amount, Slot = FLAGORO)
610           Call EnviarOferta(tUser, OfferSlot)
620       End With
End Sub

''
' Handles the "GuildAcceptPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptPeace(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim guild As String
              Dim errorStr As String
              Dim otherClanIndex As String

90            guild = buffer.ReadASCIIString()

100           otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(userindex, guild, errorStr)

110           If otherClanIndex = 0 Then
120               Call WriteConsoleMsg(userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
130           Else
140               Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
150               Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
160           End If

              'If we got here then packet is complete, copy data back to original queue
170           Call .incomingData.CopyBuffer(buffer)
180       End With

Errhandler:
          Dim error  As Long
190       error = Err.Number
200       On Error GoTo 0

          'Destroy auxiliar buffer
210       Set buffer = Nothing

220       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GuildRejectAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectAlliance(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim guild As String
              Dim errorStr As String
              Dim otherClanIndex As String

90            guild = buffer.ReadASCIIString()

100           otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(userindex, guild, errorStr)

110           If otherClanIndex = 0 Then
120               Call WriteConsoleMsg(userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
130           Else
140               Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de alianza de " & guild, FontTypeNames.FONTTYPE_GUILD))
150               Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", FontTypeNames.FONTTYPE_GUILD))
160           End If

              'If we got here then packet is complete, copy data back to original queue
170           Call .incomingData.CopyBuffer(buffer)
180       End With

Errhandler:
          Dim error  As Long
190       error = Err.Number
200       On Error GoTo 0

          'Destroy auxiliar buffer
210       Set buffer = Nothing

220       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GuildRejectPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectPeace(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim guild As String
              Dim errorStr As String
              Dim otherClanIndex As String

90            guild = buffer.ReadASCIIString()

100           otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(userindex, guild, errorStr)

110           If otherClanIndex = 0 Then
120               Call WriteConsoleMsg(userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
130           Else
140               Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de paz de " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
150               Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", FontTypeNames.FONTTYPE_GUILD))
160           End If

              'If we got here then packet is complete, copy data back to original queue
170           Call .incomingData.CopyBuffer(buffer)
180       End With

Errhandler:
          Dim error  As Long
190       error = Err.Number
200       On Error GoTo 0

          'Destroy auxiliar buffer
210       Set buffer = Nothing

220       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GuildAcceptAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptAlliance(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim guild As String
              Dim errorStr As String
              Dim otherClanIndex As String

90            guild = buffer.ReadASCIIString()

100           otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(userindex, guild, errorStr)

110           If otherClanIndex = 0 Then
120               Call WriteConsoleMsg(userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
130           Else
140               Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la alianza con " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
150               Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
160           End If

              'If we got here then packet is complete, copy data back to original queue
170           Call .incomingData.CopyBuffer(buffer)
180       End With

Errhandler:
          Dim error  As Long
190       error = Err.Number
200       On Error GoTo 0

          'Destroy auxiliar buffer
210       Set buffer = Nothing

220       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GuildOfferPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferPeace(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim guild As String
              Dim proposal As String
              Dim errorStr As String

90            guild = buffer.ReadASCIIString()
100           proposal = buffer.ReadASCIIString()

110           If modGuilds.r_ClanGeneraPropuesta(userindex, guild, RELACIONES_GUILD.PAZ, proposal, errorStr) Then
120               Call WriteConsoleMsg(userindex, "Propuesta de paz enviada.", FontTypeNames.FONTTYPE_GUILD)
130           Else
140               Call WriteConsoleMsg(userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
150           End If

              'If we got here then packet is complete, copy data back to original queue
160           Call .incomingData.CopyBuffer(buffer)
170       End With

Errhandler:
          Dim error  As Long
180       error = Err.Number
190       On Error GoTo 0

          'Destroy auxiliar buffer
200       Set buffer = Nothing

210       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GuildOfferAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferAlliance(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim guild As String
              Dim proposal As String
              Dim errorStr As String

90            guild = buffer.ReadASCIIString()
100           proposal = buffer.ReadASCIIString()

110           If modGuilds.r_ClanGeneraPropuesta(userindex, guild, RELACIONES_GUILD.ALIADOS, proposal, errorStr) Then
120               Call WriteConsoleMsg(userindex, "Propuesta de alianza enviada.", FontTypeNames.FONTTYPE_GUILD)
130           Else
140               Call WriteConsoleMsg(userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
150           End If

              'If we got here then packet is complete, copy data back to original queue
160           Call .incomingData.CopyBuffer(buffer)
170       End With

Errhandler:
          Dim error  As Long
180       error = Err.Number
190       On Error GoTo 0

          'Destroy auxiliar buffer
200       Set buffer = Nothing

210       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GuildAllianceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAllianceDetails(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim guild As String
              Dim errorStr As String
              Dim details As String

90            guild = buffer.ReadASCIIString()

100           details = modGuilds.r_VerPropuesta(userindex, guild, RELACIONES_GUILD.ALIADOS, errorStr)

110           If LenB(details) = 0 Then
120               Call WriteConsoleMsg(userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
130           Else
140               Call WriteOfferDetails(userindex, details)
150           End If

              'If we got here then packet is complete, copy data back to original queue
160           Call .incomingData.CopyBuffer(buffer)
170       End With

Errhandler:
          Dim error  As Long
180       error = Err.Number
190       On Error GoTo 0

          'Destroy auxiliar buffer
200       Set buffer = Nothing

210       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GuildPeaceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeaceDetails(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim guild As String
              Dim errorStr As String
              Dim details As String

90            guild = buffer.ReadASCIIString()

100           details = modGuilds.r_VerPropuesta(userindex, guild, RELACIONES_GUILD.PAZ, errorStr)

110           If LenB(details) = 0 Then
120               Call WriteConsoleMsg(userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
130           Else
140               Call WriteOfferDetails(userindex, details)
150           End If

              'If we got here then packet is complete, copy data back to original queue
160           Call .incomingData.CopyBuffer(buffer)
170       End With

Errhandler:
          Dim error  As Long
180       error = Err.Number
190       On Error GoTo 0

          'Destroy auxiliar buffer
200       Set buffer = Nothing

210       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GuildRequestJoinerInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestJoinerInfo(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim User As String
              Dim details As String

90            User = buffer.ReadASCIIString()

100           details = modGuilds.a_DetallesAspirante(userindex, User)

110           If LenB(details) = 0 Then
120               Call WriteConsoleMsg(userindex, "El personaje no ha mandado solicitud, o no estás habilitado para verla.", FontTypeNames.FONTTYPE_GUILD)
130           Else
140               Call WriteShowUserRequest(userindex, details)
150           End If

              'If we got here then packet is complete, copy data back to original queue
160           Call .incomingData.CopyBuffer(buffer)
170       End With

Errhandler:
          Dim error  As Long
180       error = Err.Number
190       On Error GoTo 0

          'Destroy auxiliar buffer
200       Set buffer = Nothing

210       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GuildAlliancePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAlliancePropList(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call UserList(userindex).incomingData.ReadByte

20        Call WriteAlianceProposalsList(userindex, r_ListaDePropuestas(userindex, RELACIONES_GUILD.ALIADOS))
End Sub

''
' Handles the "GuildPeacePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeacePropList(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call UserList(userindex).incomingData.ReadByte

20        Call WritePeaceProposalsList(userindex, r_ListaDePropuestas(userindex, RELACIONES_GUILD.PAZ))
End Sub

''
' Handles the "GuildDeclareWar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildDeclareWar(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim guild As String
              Dim errorStr As String
              Dim otherGuildIndex As Integer

90            guild = buffer.ReadASCIIString()

100           otherGuildIndex = modGuilds.r_DeclararGuerra(userindex, guild, errorStr)

110           If otherGuildIndex = 0 Then
120               Call WriteConsoleMsg(userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
130           Else
                  'WAR shall be!
140               Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("TU CLAN HA ENTRADO EN GUERRA CON " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
150               Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " LE DECLARA LA GUERRA A TU CLAN.", FontTypeNames.FONTTYPE_GUILD))
160               Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
170               Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
180           End If

              'If we got here then packet is complete, copy data back to original queue
190           Call .incomingData.CopyBuffer(buffer)
200       End With

Errhandler:
          Dim error  As Long
210       error = Err.Number
220       On Error GoTo 0

          'Destroy auxiliar buffer
230       Set buffer = Nothing

240       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GuildNewWebsite" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildNewWebsite(ByVal userindex As Integer)
 
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
10      If UserList(userindex).incomingData.length < 3 Then
20          Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30          Exit Sub
 
40      End If
 
        '<EhHeader>
50      On Error GoTo HandleGuildNewWebsite_Err
 
        '</EhHeader>
 
60      With UserList(userindex)
            'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
            Dim buffer  As New clsByteQueue
            Dim strTemp As String
       
70          Call buffer.CopyBuffer(.incomingData)
 
            'Remove packet ID
80          Call buffer.ReadByte
       
90          strTemp = buffer.ReadASCIIString()
       
100         Call modGuilds.ActualizarWebSite(userindex, strTemp)
 
            'If we got here then packet is complete, copy data back to original queue
110         Call .incomingData.CopyBuffer(buffer)
 
120     End With
 
        'Destroy auxiliar buffer
130     Set buffer = Nothing
 
HandleGuildNewWebsite_Err:
        Dim error As Long
140     error = Err.Number
 
150     If error <> 0 Then
 
160     Call LogError(Err.Description & " in HandleGuildNewWebsite " & "at line " & Erl & " strTemp: " & strTemp)
170     Call Err.Raise(error)
 
180     End If
 
        '</EhFooter>
End Sub

''
' Handles the "GuildAcceptNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptNewMember(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim errorStr As String
              Dim UserName As String
              Dim tUser As Integer

90            UserName = buffer.ReadASCIIString()

100           If Not modGuilds.a_AceptarAspirante(userindex, UserName, errorStr) Then
110               Call WriteConsoleMsg(userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
120           Else
130               tUser = NameIndex(UserName)
140               If tUser > 0 Then
150                   Call modGuilds.m_ConectarMiembroAClan(tUser, .GuildIndex)
160                   Call RefreshCharStatus(tUser)
170               End If

180               Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido aceptado como miembro del clan.", FontTypeNames.FONTTYPE_GUILD))
190               Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))
200           End If

              'If we got here then packet is complete, copy data back to original queue
210           Call .incomingData.CopyBuffer(buffer)
220       End With

Errhandler:
          Dim error  As Long
230       error = Err.Number
240       On Error GoTo 0

          'Destroy auxiliar buffer
250       Set buffer = Nothing

260       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GuildRejectNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectNewMember(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 01/08/07
      'Last Modification by: (liquid)
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim errorStr As String
              Dim UserName As String
              Dim Reason As String
              Dim tUser As Integer

90            UserName = buffer.ReadASCIIString()
100           Reason = buffer.ReadASCIIString()

110           If Not modGuilds.a_RechazarAspirante(userindex, UserName, errorStr) Then
120               Call WriteConsoleMsg(userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
130           Else
140               tUser = NameIndex(UserName)

150               If tUser > 0 Then
160                   Call WriteConsoleMsg(tUser, errorStr & " : " & Reason, FontTypeNames.FONTTYPE_GUILD)
170               Else
                      'hay que grabar en el char su rechazo
180                   Call modGuilds.a_RechazarAspiranteChar(UserName, .GuildIndex, Reason)
190               End If
200           End If

              'If we got here then packet is complete, copy data back to original queue
210           Call .incomingData.CopyBuffer(buffer)
220       End With

Errhandler:
          Dim error  As Long
230       error = Err.Number
240       On Error GoTo 0

          'Destroy auxiliar buffer
250       Set buffer = Nothing

260       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GuildKickMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildKickMember(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim GuildIndex As Integer

90            UserName = buffer.ReadASCIIString()

100           GuildIndex = modGuilds.m_EcharMiembroDeClan(userindex, UserName)

110           If GuildIndex > 0 Then
120               Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " fue expulsado del clan.", FontTypeNames.FONTTYPE_GUILD))
130               Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
140           Else
150               Call WriteConsoleMsg(userindex, "No puedes expulsar ese personaje del clan.", FontTypeNames.FONTTYPE_GUILD)
160           End If

              'If we got here then packet is complete, copy data back to original queue
170           Call .incomingData.CopyBuffer(buffer)
180       End With

Errhandler:
          Dim error  As Long
190       error = Err.Number
200       On Error GoTo 0

          'Destroy auxiliar buffer
210       Set buffer = Nothing

220       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GuildUpdateNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildUpdateNews(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

90            Call modGuilds.ActualizarNoticias(userindex, buffer.ReadASCIIString())

              'If we got here then packet is complete, copy data back to original queue
100           Call .incomingData.CopyBuffer(buffer)
110       End With

Errhandler:
          Dim error  As Long
120       error = Err.Number
130       On Error GoTo 0

          'Destroy auxiliar buffer
140       Set buffer = Nothing

150       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GuildMemberInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberInfo(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

90            Call modGuilds.SendDetallesPersonaje(userindex, buffer.ReadASCIIString())

              'If we got here then packet is complete, copy data back to original queue
100           Call .incomingData.CopyBuffer(buffer)
110       End With

Errhandler:
          Dim error  As Long
120       error = Err.Number
130       On Error GoTo 0

          'Destroy auxiliar buffer
140       Set buffer = Nothing

150       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GuildOpenElections" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOpenElections(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              Dim error As String

30            If Not modGuilds.v_AbrirElecciones(userindex, error) Then
40                Call WriteConsoleMsg(userindex, error, FontTypeNames.FONTTYPE_GUILD)
50            Else
60                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("¡Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & .Name, FontTypeNames.FONTTYPE_GUILD))
70            End If
80        End With
End Sub

''
' Handles the "GuildRequestMembership" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestMembership(ByVal userindex As Integer)
 
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
10      If UserList(userindex).incomingData.length < 5 Then
20          Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30          Exit Sub
 
40      End If
       
        '<EhHeader>
50      On Error GoTo HandleGuildRequestMembership_Err
 
        '</EhHeader>
       
60      With UserList(userindex)
            'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
            Dim buffer As New clsByteQueue
70          Call buffer.CopyBuffer(.incomingData)
 
            'Remove packet ID
80          Call buffer.ReadByte
 
            Dim guild       As String
            Dim application As String
            Dim errorStr    As String
 
90          guild = buffer.ReadASCIIString()
100         application = buffer.ReadASCIIString()
 
110         If Not modGuilds.a_NuevoAspirante(userindex, guild, application, errorStr) Then
120             Call WriteConsoleMsg(userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
130         Else
140             Call WriteConsoleMsg(userindex, "Tu solicitud ha sido enviada. Espera prontas noticias del líder de " _
                        & guild & ".", FontTypeNames.FONTTYPE_GUILD)
 
150         End If
 
            'If we got here then packet is complete, copy data back to original queue
160         Call .incomingData.CopyBuffer(buffer)
 
170     End With
       
        'Destroy auxiliar buffer
180     Set buffer = Nothing
     
HandleGuildRequestMembership_Err:
 
        Dim error As Long
190     error = Err.Number
 
200     If error <> 0 Then
         
210         Call LogError(Err.Description & " in HandleGuildRequestMembership " & "at line " & Erl & ". Guild: " & guild & " application: " & application)
220         Call Err.Raise(error)
 
230     End If
 
        '</EhFooter>
End Sub

''
' Handles the "GuildRequestDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestDetails(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
   On Error GoTo HandleGuildRequestDetails_Error

10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte
              Dim GuildName As String
85            GuildName = buffer.ReadASCIIString()
              
90            Call modGuilds.SendGuildDetails(userindex, GuildName)

              'If we got here then packet is complete, copy data back to original queue
100           Call .incomingData.CopyBuffer(buffer)
110       End With

Errhandler:
          Dim error  As Long
120       error = Err.Number
130       On Error GoTo 0

          'Destroy auxiliar buffer
140       Set buffer = Nothing

150       If error <> 0 Then _
             Err.Raise error

   On Error GoTo 0
   Exit Sub

HandleGuildRequestDetails_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleGuildRequestDetails of Módulo Protocol in line " & Erl & ", GUILDNAME: " & GuildName
End Sub
Private Sub HandleOnline(ByVal userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 27/01/2010 (JoaCo)
    'mandamos la lista entera de nombres
    '***************************************************
    Dim i     As Long
    Dim Count As Long
    Dim list  As String
   
    With UserList(userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
       
        For i = 1 To LastUser

            If LenB(UserList(i).Name) <> 0 Then

                If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Count = Count + 1

            End If

        Next i

        For i = 1 To LastUser

            If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then list = list & UserList(i).Name & ", "
        Next i
       
        If LenB(list) <> 0 Then
            list = Left$(list, Len(list) - 2)
            Call WriteConsoleMsg(userindex, "Nombres: " & list & ".", FontTypeNames.FONTTYPE_CENTINELA)
            WriteConsoleMsg userindex, "Números de usuarios online: " & CStr(Count) & "", FontTypeNames.FONTTYPE_CENTINELA
        Else
            Call WriteConsoleMsg(userindex, "No hay usuarios Online.", FontTypeNames.FONTTYPE_INFO)

        End If
       
    End With

End Sub

''
' Handles the "Quit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 04/15/2008 (NicoNZ)
      'If user is invisible, it automatically becomes
      'visible before doing the countdown to exit
      '04/15/2008 - No se reseteaban lso contadores de invi ni de ocultar. (NicoNZ)
      '***************************************************
          Dim tUser  As Integer
          Dim isNotVisible As Boolean

10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.automatico = True Then
40                Call WriteConsoleMsg(userindex, "No puedes salir estando en un evento.", FontTypeNames.FONTTYPE_WARNING)
50                Exit Sub
60            End If

70            If .flags.Plantico = True Then
80                Call WriteConsoleMsg(userindex, "No puedes salir estando en un evento.", FontTypeNames.FONTTYPE_WARNING)
90                Exit Sub
100           End If

110           If .flags.Paralizado = 1 Then
120               Call WriteConsoleMsg(userindex, "No puedes salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
130               Exit Sub
140           End If

150           If .flags.Montando = 1 Then
160               Call WriteConsoleMsg(userindex, "No puedes salir mientras te encuentres montando.", FontTypeNames.FONTTYPE_CONSEJOVesA)
170               Exit Sub
180           End If

              'exit secure commerce
190           If .ComUsu.DestUsu > 0 Then
200               tUser = .ComUsu.DestUsu

210               If UserList(tUser).flags.UserLogged Then
220                   If UserList(tUser).ComUsu.DestUsu = userindex Then
230                       Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_GUILD)
240                       Call FinComerciarUsu(tUser)
250                   End If
260               End If

270               Call WriteConsoleMsg(userindex, "Comercio cancelado.", FontTypeNames.FONTTYPE_GUILD)
280               Call FinComerciarUsu(userindex)
290           End If

300           Call Cerrar_Usuario(userindex)
310       End With
End Sub

''
' Handles the "GuildLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildLeave(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
          Dim GuildIndex As Integer

10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'obtengo el guildindex
30            GuildIndex = m_EcharMiembroDeClan(userindex, .Name)

40            If GuildIndex > 0 Then
50                Call WriteConsoleMsg(userindex, "Dejas el clan.", FontTypeNames.FONTTYPE_GUILD)
60                Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(.Name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
70            Else
80                Call WriteConsoleMsg(userindex, "Tú no puedes salir de este clan.", FontTypeNames.FONTTYPE_GUILD)
90            End If
100       End With
End Sub

''
' Handles the "RequestAccountState" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAccountState(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
          Dim earnings As Integer
          Dim Percentage As Integer

10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'Dead people can't check their accounts
30            If .flags.Muerto = 1 Then
40                Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
50                Exit Sub
60            End If

              'Validate target NPC
70            If .flags.TargetNPC = 0 Then
80                Call WriteConsoleMsg(userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
90                Exit Sub
100           End If

110           If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
120               Call WriteConsoleMsg(userindex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
130               Exit Sub
140           End If

150           Select Case Npclist(.flags.TargetNPC).NPCtype
              Case eNPCType.Banquero
160               Call WriteChatOverHead(userindex, "Tienes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

170           Case eNPCType.Timbero
180               If Not .flags.Privilegios And PlayerType.User Then
190                   earnings = Apuestas.Ganancias - Apuestas.Perdidas

200                   If earnings >= 0 And Apuestas.Ganancias <> 0 Then
210                       Percentage = Int(earnings * 100 / Apuestas.Ganancias)
220                   End If

230                   If earnings < 0 And Apuestas.Perdidas <> 0 Then
240                       Percentage = Int(earnings * 100 / Apuestas.Perdidas)
250                   End If

260                   Call WriteConsoleMsg(userindex, "Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & earnings & " (" & Percentage & "%) Jugadas: " & Apuestas.Jugadas, FontTypeNames.FONTTYPE_INFO)
270               End If
280           End Select
290       End With
End Sub

''
' Handles the "PetStand" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetStand(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'Dead people can't use pets
30            If .flags.Muerto = 1 Then
40                Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
50                Exit Sub
60            End If

              'Validate target NPC
70            If .flags.TargetNPC = 0 Then
80                Call WriteConsoleMsg(userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
90                Exit Sub
100           End If

              'Make sure it's close enough
110           If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
                  'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
120               Call WriteShortMsj(userindex, 8, FontTypeNames.FONTTYPE_INFO)
130               Exit Sub
140           End If

              'Make sure it's his pet
150           If Npclist(.flags.TargetNPC).MaestroUser <> userindex Then Exit Sub

              'Do it!
160           Npclist(.flags.TargetNPC).Movement = TipoAI.ESTATICO

170           Call Expresar(.flags.TargetNPC, userindex)
180       End With
End Sub

''
' Handles the "PetFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetFollow(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'Dead users can't use pets
30            If .flags.Muerto = 1 Then
40                Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
50                Exit Sub
60            End If

              'Validate target NPC
70            If .flags.TargetNPC = 0 Then
80                Call WriteConsoleMsg(userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
90                Exit Sub
100           End If

              'Make sure it's close enough
110           If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
                  'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
120               Call WriteShortMsj(userindex, 8, FontTypeNames.FONTTYPE_INFO)
130               Exit Sub
140           End If

              'Make usre it's the user's pet
150           If Npclist(.flags.TargetNPC).MaestroUser <> userindex Then Exit Sub

              'Do it
160           Call FollowAmo(.flags.TargetNPC)

170           Call Expresar(.flags.TargetNPC, userindex)
180       End With
End Sub


''
' Handles the "ReleasePet" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReleasePet(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 18/11/2009
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'Dead users can't use pets
30            If .flags.Muerto = 1 Then
40                Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
50                Exit Sub
60            End If

              'Validate target NPC
70            If .flags.TargetNPC = 0 Then
80                Call WriteConsoleMsg(userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
90                Exit Sub
100           End If

              'Make sure it's close enough
110           If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
                  'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
120               Call WriteShortMsj(userindex, 8, FontTypeNames.FONTTYPE_INFO)
130               Exit Sub
140           End If

              'Make usre it's the user's pet
150           If Npclist(.flags.TargetNPC).MaestroUser <> userindex Then Exit Sub

              'Do it
160           Call QuitarPet(userindex, .flags.TargetNPC)

170       End With
End Sub

''
' Handles the "TrainList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrainList(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'Dead users can't use pets
30            If .flags.Muerto = 1 Then
40                Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
50                Exit Sub
60            End If

              'Validate target NPC
70            If .flags.TargetNPC = 0 Then
80                Call WriteConsoleMsg(userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
90                Exit Sub
100           End If

              'Make sure it's close enough
110           If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
                  'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
120               Call WriteShortMsj(userindex, 8, FontTypeNames.FONTTYPE_INFO)
130               Exit Sub
140           End If

              'Make sure it's the trainer
150           If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub

160           Call WriteTrainerCreatureList(userindex, .flags.TargetNPC)
170       End With
End Sub

''
' Handles the "Rest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRest(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'Dead users can't use pets
30            If .flags.Muerto = 1 Then
40                Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
                  'Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!! Solo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
50                Exit Sub
60            End If

70            If HayOBJarea(.Pos, FOGATA) Then
80                Call WriteRestOK(userindex)

90                If Not .flags.Descansar Then
100                   Call WriteConsoleMsg(userindex, "Te acomodás junto a la fogata y comienzas a descansar.", FontTypeNames.FONTTYPE_INFO)
110               Else
120                   Call WriteConsoleMsg(userindex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
130               End If

140               .flags.Descansar = Not .flags.Descansar
150           Else
160               If .flags.Descansar Then
170                   Call WriteRestOK(userindex)
180                   Call WriteConsoleMsg(userindex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)

190                   .flags.Descansar = False
200                   Exit Sub
210               End If

220               Call WriteConsoleMsg(userindex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)
230           End If
240       End With
End Sub
Private Sub HandleFianzah(ByVal userindex As Integer)
      '***************************************************
      'Author: Matías Ezequiel
      'Last Modification: 16/03/2016 by DS
      'Sistema de fianzas TDS.
      '***************************************************
          Dim Fianza As Long

10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
60            Call .incomingData.ReadByte
70            Fianza = .incomingData.ReadLong


80            If Not UserList(userindex).Pos.map = 1 Then
90                Call WriteConsoleMsg(userindex, "No puedes pagar la fianza si no estás en Ullathorpe.", FontTypeNames.FONTTYPE_INFO)
100               Exit Sub
110           End If
              
120           'LogAntiCheat "El personaje " & .Name & " con IP: " & .ip & " ha usado el paquete de fianza."

              ' @@ Rezniaq bronza
130           If .flags.Muerto Then Call WriteConsoleMsg(userindex, "Estás muerto.", FontTypeNames.FONTTYPE_INFO): Exit Sub

140           If Not criminal(userindex) Then Call WriteConsoleMsg(userindex, "Ya eres ciudadano, no podrás realizar la fianza.", FontTypeNames.FONTTYPE_INFO): Exit Sub



150           If Fianza <= 0 Then
                  '  Call WriteConsoleMsg(UserIndex, "El minimo de fianza es 1", FontTypeNames.FONTTYPE_INFO)
160               Exit Sub
170           ElseIf (Fianza * 25) > .Stats.Gld Then
180               Call WriteConsoleMsg(userindex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_INFO)
190               Exit Sub
200           End If

210           .Reputacion.NobleRep = .Reputacion.NobleRep + Fianza
220           .Stats.Gld = .Stats.Gld - Fianza * 25

230           Call WriteConsoleMsg(userindex, "Has ganado " & Fianza & " puntos de noble.", FontTypeNames.FONTTYPE_INFO)
240           Call WriteConsoleMsg(userindex, "Se te han descontado " & Fianza * 25 & " monedas de oro", FontTypeNames.FONTTYPE_INFO)
250           Call WriteUpdateGold(userindex)
260           Call RefreshCharStatus(userindex)
270       End With
          
          

End Sub

''
' Handles the "Meditate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMeditate(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 04/15/08 (NicoNZ)
      'Arreglé un bug que mandaba un index de la meditacion diferente
      'al que decia el server.
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'Dead users can't use pets
30            If .flags.Muerto = 1 Then
40                Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
                  'Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!! Sólo puedes meditar cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
50                Exit Sub
60            End If

              'Can he meditate?
70            If .Stats.MaxMAN = 0 Then
80                Call WriteConsoleMsg(userindex, "Sólo las clases mágicas conocen el arte de la meditación.", FontTypeNames.FONTTYPE_INFO)
90                Exit Sub
100           End If

              'Admins don't have to wait :D
110           If Not .flags.Privilegios And PlayerType.User Then
120               .Stats.MinMAN = .Stats.MaxMAN
130               Call WriteUpdateMana(userindex)
140               Call WriteUpdateFollow(userindex)
150               Exit Sub
160           End If

170           Call WriteMeditateToggle(userindex)

180           If .flags.Meditando Then _
                 Call WriteConsoleMsg(userindex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)

190           .flags.Meditando = Not .flags.Meditando

              'Barrin 3/10/03 Tiempo de inicio al meditar
200           If .flags.Meditando Then

210               .Char.loops = INFINITE_LOOPS

                  'Show proper FX according to level
220               If .Stats.ELV < 15 Then
230                   .Char.FX = FXIDs.FXMEDITARCHICO

                      'ElseIf .Stats.ELV < 25 Then
                      '     .Char.FX = FXIDs.FXMEDITARMEDIANO

240               ElseIf .Stats.ELV < 30 Then
250                   .Char.FX = FXIDs.FXMEDITARMEDIANO

260               ElseIf .Stats.ELV < 45 Then
270                   .Char.FX = FXIDs.FXMEDITARGRANDE

280               Else
290                   .Char.FX = FXIDs.FXMEDITARXXGRANDE
300               End If
                  
310               If .flags.IsDios Then
320                   .Char.FX = 29
330               End If
                  
340               Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS))
350           Else
                  '.Counters.bPuedeMeditar = False

360               .Char.FX = 0
370               .Char.loops = 0
380               Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
390           End If
400       End With
End Sub

''
' Handles the "Resucitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'Se asegura que el target es un npc
30            If .flags.TargetNPC = 0 Then
40                Call WriteConsoleMsg(userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
50                Exit Sub
60            End If

              'Validate NPC and make sure player is dead
70            If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor _
                  And (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(userindex))) _
                  Or .flags.Muerto = 0 Then Exit Sub

              'Make sure it's close enough
80            If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
90                Call WriteConsoleMsg(userindex, "El sacerdote no puede resucitarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
100               Exit Sub
110           End If

120           Call RevivirUsuario(userindex)
130           Call WriteConsoleMsg(userindex, "¡¡Has sido resucitado!!", FontTypeNames.FONTTYPE_INFO)
140       End With
End Sub

''
' Handles the "Consultation" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleConsultation(ByVal userindex As String)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 01/05/2010
      'Habilita/Deshabilita el modo consulta.
      '01/05/2010: ZaMa - Agrego validaciones.
      '***************************************************

          Dim UserConsulta As Integer

10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              ' Comando exclusivo para gms
30            If Not EsGm(userindex) Then Exit Sub

40            UserConsulta = .flags.TargetUser

              'Se asegura que el target es un usuario
50            If UserConsulta = 0 Then
60                Call WriteConsoleMsg(userindex, "Primero tienes que seleccionar un usuario, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
70                Exit Sub
80            End If

              ' No podes ponerte a vos mismo en modo consulta.
90            If UserConsulta = userindex Then Exit Sub

              ' No podes estra en consulta con otro gm
100           If EsGm(UserConsulta) Then
110               Call WriteConsoleMsg(userindex, "No puedes iniciar el modo consulta con otro administrador.", FontTypeNames.FONTTYPE_INFO)
120               Exit Sub
130           End If

              Dim UserName As String
140           UserName = UserList(UserConsulta).Name

              ' Si ya estaba en consulta, termina la consulta
150           If UserList(UserConsulta).flags.EnConsulta Then
160               Call WriteConsoleMsg(userindex, "Has terminado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
170               Call WriteConsoleMsg(UserConsulta, "Has terminado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
180               Call LogGM(.Name, "Termino consulta con " & UserName)

190               UserList(UserConsulta).flags.EnConsulta = False

                  ' Sino la inicia
200           Else
210               Call WriteConsoleMsg(userindex, "Has iniciado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
220               Call WriteConsoleMsg(UserConsulta, "Has iniciado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
230               Call LogGM(.Name, "Inicio consulta con " & UserName)

240               With UserList(UserConsulta)
250                   .flags.EnConsulta = True

                      ' Pierde invi u ocu
260                   If .flags.invisible = 1 Or .flags.Oculto = 1 Then
270                       .flags.Oculto = 0
280                       .flags.invisible = 0
290                       .Counters.TiempoOculto = 0
300                       .Counters.Invisibilidad = 0

310                       Call UsUaRiOs.SetInvisible(UserConsulta, UserList(UserConsulta).Char.CharIndex, False)
320                   End If
330               End With
340           End If

350           Call UsUaRiOs.SetConsulatMode(UserConsulta)
360       End With

End Sub

''
' Handles the "Heal" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHeal(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'Se asegura que el target es un npc
30            If .flags.TargetNPC = 0 Then
40                Call WriteConsoleMsg(userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
50                Exit Sub
60            End If

70            If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor _
                  And Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie) _
                  Or .flags.Muerto <> 0 Then Exit Sub

80            If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
                  'Call WriteConsoleMsg(Userindex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
90                Call WriteShortMsj(userindex, 7, FontTypeNames.FONTTYPE_INFO)
100               Exit Sub
110           End If

120           .Stats.MinHp = .Stats.MaxHp

130           Call WriteUpdateHP(userindex)
140           Call WriteUpdateFollow(userindex)

150           Call WriteConsoleMsg(userindex, "¡¡Has sido curado!!", FontTypeNames.FONTTYPE_INFO)
160       End With
End Sub

''
' Handles the "RequestStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestStats(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call UserList(userindex).incomingData.ReadByte

20        Call SendUserStatsTxt(userindex, userindex)
End Sub

''
' Handles the "Help" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHelp(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call UserList(userindex).incomingData.ReadByte

20        Call SendHelp(userindex)
End Sub

''
' Handles the "CommerceStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
          Dim i      As Integer
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'Dead people can't commerce
30            If .flags.Muerto = 1 Then
40                Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
50                Exit Sub
60            End If

70            If Not UserList(userindex).Pos.map = 200 Then
80                WriteConsoleMsg userindex, "Para comerciar debes encontrarte en la Zona de Comercio.", FontTypeNames.FONTTYPE_INFO
90                Exit Sub
100           End If

110           If .flags.Envenenado = 1 Then
120               Call WriteConsoleMsg(userindex, "¡¡Estás envenenado!!", FontTypeNames.FONTTYPE_INFO)
130               Exit Sub
140           End If

150           If MapInfo(.Pos.map).Pk = True Then
160               Call WriteConsoleMsg(userindex, "¡Para poder comerciar debes estar en una ciudad segura!", FONTTYPE_INFO)
170               Exit Sub
180           End If

              'Is it already in commerce mode??
190           If .flags.Comerciando Then
200               Call WriteConsoleMsg(userindex, "Ya estás comerciando.", FontTypeNames.FONTTYPE_INFO)
210               Exit Sub
220           End If

              'Validate target NPC
230           If .flags.TargetNPC > 0 Then
                  'Does the NPC want to trade??
240               If Npclist(.flags.TargetNPC).Comercia = 0 Then
250                   If LenB(Npclist(.flags.TargetNPC).desc) <> 0 Then
260                       Call WriteChatOverHead(userindex, "No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
270                   End If

280                   Exit Sub
290               End If

300               If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
310                   Call WriteConsoleMsg(userindex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
320                   Exit Sub
330               End If

                  'Start commerce....
340               Call IniciarComercioNPC(userindex)
                  '[Alejo]
350           ElseIf .flags.TargetUser > 0 Then
                  'User commerce...
                  'Can he commerce??
360               If .flags.Privilegios And PlayerType.Consejero Then
370                   Call WriteConsoleMsg(userindex, "No puedes vender ítems.", FontTypeNames.FONTTYPE_WARNING)
380                   Exit Sub
390               End If

                  'Is the other one dead??
400               If UserList(.flags.TargetUser).flags.Muerto = 1 Then
410                   Call WriteConsoleMsg(userindex, "¡¡No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
420                   Exit Sub
430               End If

                  'Is it me??
440               If .flags.TargetUser = userindex Then
450                   Call WriteConsoleMsg(userindex, "¡¡No puedes comerciar con vos mismo!!", FontTypeNames.FONTTYPE_INFO)
460                   Exit Sub
470               End If

                  'Check distance
480               If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
490                   Call WriteConsoleMsg(userindex, "Estás demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
500                   Exit Sub
510               End If

                  'Is he already trading?? is it with me or someone else??
520               If UserList(.flags.TargetUser).flags.Comerciando = True And _
                     UserList(.flags.TargetUser).ComUsu.DestUsu <> userindex Then
530                   Call WriteConsoleMsg(userindex, "No puedes comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
540                   Exit Sub
550               End If

                  'Initialize some variables...
560               .ComUsu.DestUsu = .flags.TargetUser
570               .ComUsu.DestNick = UserList(.flags.TargetUser).Name
580               For i = 1 To MAX_OFFER_SLOTS
590                   .ComUsu.cant(i) = 0
600                   .ComUsu.Objeto(i) = 0
610               Next i
620               .ComUsu.GoldAmount = 0

630               .ComUsu.Acepto = False
640               .ComUsu.Confirmo = False

                  'Rutina para comerciar con otro usuario
650               Call IniciarComercioConUsuario(userindex, .flags.TargetUser)
660           Else
670               Call WriteConsoleMsg(userindex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
680           End If
690       End With
End Sub

''
' Handles the "BankStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankStart(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'Dead people can't commerce
30            If .flags.Muerto = 1 Then
40                Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
50                Exit Sub
60            End If

70            If .flags.Comerciando Then
80                Call WriteConsoleMsg(userindex, "Ya estás comerciando.", FontTypeNames.FONTTYPE_INFO)
90                Exit Sub
100           End If

              'Validate target NPC
110           If .flags.TargetNPC > 0 Then
120               If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
130                   Call WriteConsoleMsg(userindex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
140                   Exit Sub
150               End If

                  'If it's the banker....
160               If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Banquero Then
170                   Call IniciarDeposito(userindex)
180               End If
190           Else
200               Call WriteConsoleMsg(userindex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
210           End If
220       End With
End Sub

''
' Handles the "Enlist" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnlist(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'Validate target NPC
30            If .flags.TargetNPC = 0 Then
40                Call WriteConsoleMsg(userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
50                Exit Sub
60            End If

70            If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
                 Or .flags.Muerto <> 0 Then Exit Sub

80            If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
90                Call WriteConsoleMsg(userindex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)
100               Exit Sub
110           End If

120           If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
130               Call EnlistarArmadaReal(userindex)
140           Else
150               Call EnlistarCaos(userindex)
160           End If
170       End With
End Sub

''
' Handles the "Information" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInformation(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
          Dim Matados As Integer
          Dim NextRecom As Integer
          Dim Diferencia As Integer

10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'Validate target NPC
30            If .flags.TargetNPC = 0 Then
40                Call WriteConsoleMsg(userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
50                Exit Sub
60            End If

70            If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
                 Or .flags.Muerto <> 0 Then Exit Sub

80            If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
                  'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
90                Call WriteShortMsj(userindex, 8, FontTypeNames.FONTTYPE_INFO)
100               Exit Sub
110           End If


120           NextRecom = .Faccion.NextRecompensa

130           If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
140               If .Faccion.ArmadaReal = 0 Then
150                   Call WriteChatOverHead(userindex, "¡¡No perteneces a las tropas reales!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
160                   Exit Sub
170               End If

180               Matados = .Faccion.CriminalesMatados
190               Diferencia = NextRecom - Matados

200               If Diferencia > 0 Then
210                   Call WriteChatOverHead(userindex, "Tu deber es combatir criminales, mata " & Diferencia & " criminales más y te daré una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
220               Else
230                   Call WriteChatOverHead(userindex, "Tu deber es combatir criminales, y ya has matado los suficientes como para merecerte una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
240               End If
250           Else
260               If .Faccion.FuerzasCaos = 0 Then
270                   Call WriteChatOverHead(userindex, "¡¡No perteneces a la legión oscura!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
280                   Exit Sub
290               End If

300               Matados = .Faccion.CiudadanosMatados
310               Diferencia = NextRecom - Matados

320               If Diferencia > 0 Then
330                   Call WriteChatOverHead(userindex, "Tu deber es sembrar el caos y la desesperanza, mata " & Diferencia & " ciudadanos más y te daré una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
340               Else
350                   Call WriteChatOverHead(userindex, "Tu deber es sembrar el caos y la desesperanza, y creo que estás en condiciones de merecer una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
360               End If
370           End If
380       End With
End Sub

''
' Handles the "Reward" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReward(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'Validate target NPC
30            If .flags.TargetNPC = 0 Then
40                Call WriteConsoleMsg(userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
50                Exit Sub
60            End If

70            If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
                 Or .flags.Muerto <> 0 Then Exit Sub

80            If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
90                Call WriteShortMsj(userindex, 8, FontTypeNames.FONTTYPE_INFO)
100               Exit Sub
110           End If

120           If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
130               If .Faccion.ArmadaReal = 0 Then
140                   Call WriteChatOverHead(userindex, "¡¡No perteneces a las tropas reales!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
150                   Exit Sub
160               End If
170               Call RecompensaArmadaReal(userindex)
180           Else
190               If .Faccion.FuerzasCaos = 0 Then
200                   Call WriteChatOverHead(userindex, "¡¡No perteneces a la legión oscura!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
210                   Exit Sub
220               End If
230               Call RecompensaCaos(userindex)
240           End If
250       End With
End Sub

''
' Handles the "UpTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUpTime(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 01/10/08
      '01/10/2008 - Marcos Martinez (ByVal) - Automatic restart removed from the server along with all their assignments and varibles
      '***************************************************
      'Remove packet ID
10        Call UserList(userindex).incomingData.ReadByte

          Dim time   As Long
          Dim UpTimeStr As String

          'Get total time in seconds
20        time = ((GetTickCount() And &H7FFFFFFF) - tInicioServer) \ 1000

          'Get times in dd:hh:mm:ss format
30        UpTimeStr = (time Mod 60) & " segundos."
40        time = time \ 60

50        UpTimeStr = (time Mod 60) & " minutos, " & UpTimeStr
60        time = time \ 60

70        UpTimeStr = (time Mod 24) & " horas, " & UpTimeStr
80        time = time \ 24

90        If time = 1 Then
100           UpTimeStr = time & " día, " & UpTimeStr
110       Else
120           UpTimeStr = time & " días, " & UpTimeStr
130       End If

140       Call WriteConsoleMsg(userindex, "Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)
End Sub


''
' Handles the "ShareNpc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShareNpc(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 15/04/2010
      'Shares owned npcs with other user
      '***************************************************

          Dim targetUserIndex As Integer
          Dim SharingUserIndex As Integer

10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              ' Didn't target any user
30            targetUserIndex = .flags.TargetUser
40            If targetUserIndex = 0 Then Exit Sub

              ' Can't share with admins
50            If EsGm(targetUserIndex) Then
60                Call WriteConsoleMsg(userindex, "No puedes compartir npcs con administradores!!", FontTypeNames.FONTTYPE_INFO)
70                Exit Sub
80            End If

              ' Pk or Caos?
90            If criminal(userindex) Then
                  ' Caos can only share with other caos
100               If esCaos(userindex) Then
110                   If Not esCaos(targetUserIndex) Then
120                       Call WriteConsoleMsg(userindex, "Solo puedes compartir npcs con miembros de tu misma facción!!", FontTypeNames.FONTTYPE_INFO)
130                       Exit Sub
140                   End If

                      ' Pks don't need to share with anyone
150               Else
160                   Exit Sub
170               End If

                  ' Ciuda or Army?
180           Else
                  ' Can't share
190               If criminal(targetUserIndex) Then
200                   Call WriteConsoleMsg(userindex, "No puedes compartir npcs con criminales!!", FontTypeNames.FONTTYPE_INFO)
210                   Exit Sub
220               End If
230           End If

              ' Already sharing with target
240           SharingUserIndex = .flags.ShareNpcWith
250           If SharingUserIndex = targetUserIndex Then Exit Sub

              ' Aviso al usuario anterior que dejo de compartir
260           If SharingUserIndex <> 0 Then
270               Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
280               Call WriteConsoleMsg(userindex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
290           End If

300           .flags.ShareNpcWith = targetUserIndex

310           Call WriteConsoleMsg(targetUserIndex, .Name & " ahora comparte sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
320           Call WriteConsoleMsg(userindex, "Ahora compartes tus npcs con " & UserList(targetUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)

330       End With

End Sub

''
' Handles the "StopSharingNpc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleStopSharingNpc(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 15/04/2010
      'Stop Sharing owned npcs with other user
      '***************************************************

          Dim SharingUserIndex As Integer

10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            SharingUserIndex = .flags.ShareNpcWith

40            If SharingUserIndex <> 0 Then

                  ' Aviso al que compartia y al que le compartia.
50                Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
60                Call WriteConsoleMsg(SharingUserIndex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)

70                .flags.ShareNpcWith = 0
80            End If

90        End With

End Sub

''
' Handles the "Inquiry" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiry(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call UserList(userindex).incomingData.ReadByte

20        ConsultaPopular.SendInfoEncuesta (userindex)
End Sub

''
' Handles the "GuildMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMessage(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 15/07/2009
      '02/03/2009: ZaMa - Arreglado un indice mal pasado a la funcion de cartel de clanes overhead.
      '15/07/2009: ZaMa - Now invisible admins only speak by console
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim chat As String
              Dim CanTalk As Boolean
              
90            chat = buffer.ReadASCIIString()

              'If we got here then packet is complete, copy data back to original queue
100           Call .incomingData.CopyBuffer(buffer)
              
110           If LenB(chat) <> 0 Then
                  'Analize chat...
120               Call Statistics.ParseChat(chat)
                  
130               CanTalk = True
140               If .flags.SlotEvent > 0 Then
150                   If Events(.flags.SlotEvent).Modality = DeathMatch Then
160                       CanTalk = False
170                   End If
180               End If
                  
190               If CanTalk Then
200                   If .GuildIndex > 0 Then
210                       Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & "> " & chat))
          
                          'If Not (.flags.AdminInvisible = 1) Then _
                           '    Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageChatOverHead("< " & chat & " >", .Char.CharIndex, vbYellow))
220                   End If
230               End If
240           End If

250       End With

Errhandler:
          Dim error  As Long
260       error = Err.Number
270       On Error GoTo 0

          'Destroy auxiliar buffer
280       Set buffer = Nothing

290       If error <> 0 Then _
             Err.Raise error
End Sub
Private Sub HandleGroupMessage(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim chat As String

90            chat = buffer.ReadASCIIString()
              
100           If LenB(chat) <> 0 Then
                  'Analize chat...
110               Call Statistics.ParseChat(chat)
                  
                  If .GroupIndex > 0 Then
                       mGroup.SendMessageGroup .GroupIndex, .Name, chat
                  End If
130           End If

              'If we got here then packet is complete, copy data back to original queue
140           Call .incomingData.CopyBuffer(buffer)
150       End With

Errhandler:
          Dim error  As Long
160       error = Err.Number
170       On Error GoTo 0

          'Destroy auxiliar buffer
180       Set buffer = Nothing

190       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "CentinelReport" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCentinelReport(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

70            Call CentinelaCheckClave(userindex, .incomingData.ReadInteger())
80        End With
End Sub

''
' Handles the "GuildOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnline(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              Dim onlinelist As String

30            onlinelist = modGuilds.m_ListaDeMiembrosOnline(userindex, .GuildIndex)

40            If .GuildIndex <> 0 Then
50                Call WriteConsoleMsg(userindex, "Compañeros de tu clan conectados: " & onlinelist, FontTypeNames.FONTTYPE_GUILDMSG)
60            Else
70                Call WriteConsoleMsg(userindex, "No pertences a ningún clan.", FontTypeNames.FONTTYPE_GUILDMSG)
80            End If
90        End With
End Sub

''
' Handles the "CouncilMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilMessage(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler

60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim chat As String

90            chat = buffer.ReadASCIIString()

100           If LenB(chat) <> 0 Then
              
                  'Analize chat...
110               Call Statistics.ParseChat(chat)

120               If .flags.Privilegios And PlayerType.RoyalCouncil Then
130                   Call SendData(SendTarget.ToConsejo, userindex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJO))
140               ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
150                   Call SendData(SendTarget.ToConsejoCaos, userindex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))
160               End If
170           End If

              'If we got here then packet is complete, copy data back to original queue
180           Call .incomingData.CopyBuffer(buffer)
190       End With

Errhandler:
          Dim error  As Long
200       error = Err.Number
210       On Error GoTo 0

          'Destroy auxiliar buffer
220       Set buffer = Nothing

230       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "RoleMasterRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoleMasterRequest(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim request As String

90            request = buffer.ReadASCIIString()

100           If LenB(request) <> 0 Then
110               Call WriteConsoleMsg(userindex, "Su solicitud ha sido enviada.", FontTypeNames.FONTTYPE_INFO)
120               Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.Name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))
130           End If

              'If we got here then packet is complete, copy data back to original queue
140           Call .incomingData.CopyBuffer(buffer)
150       End With

Errhandler:
          Dim error  As Long
160       error = Err.Number
170       On Error GoTo 0

          'Destroy auxiliar buffer
180       Set buffer = Nothing

190       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GMRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMRequest(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If Not Ayuda.Existe(.Name) Then
40                Call WriteConsoleMsg(userindex, "El mensaje ha sido entregado, ahora sólo debes esperar que se desocupe algún GM.", FontTypeNames.FONTTYPE_INFO)
50                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(LCase$(.Name) & " GM: " & "Un usuario mando /GM por favor usa el comando /SHOW SOS", FontTypeNames.FONTTYPE_ADMIN))
60                Call Ayuda.Push(.Name)
70            Else
80                Call Ayuda.Quitar(.Name)
90                Call Ayuda.Push(.Name)
100               Call WriteConsoleMsg(userindex, "Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)
110           End If
120       End With
End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
          
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim Description As String, tmpStr As String, p As String

90            Description = buffer.ReadASCIIString()
              
              'If we got here then packet is complete, copy data back to original queue
100           Call .incomingData.CopyBuffer(buffer)
           
110               If .flags.Muerto = 1 Then
120                   Call WriteConsoleMsg(userindex, "No puedes cambiar la descripción estando muerto.", FontTypeNames.FONTTYPE_INFO)
130               Else
140                   If Not AsciiValidos(Description) Then
150                       Call WriteConsoleMsg(userindex, "La descripción tiene caracteres inválidos.", FontTypeNames.FONTTYPE_INFO)
160                   Else
170                       .desc = Trim$(Description)
180                       Call WriteConsoleMsg(userindex, "La descripción ha cambiado.", FontTypeNames.FONTTYPE_INFO)
190                   End If
200               End If

210       End With

Errhandler:
          Dim error  As Long
220       error = Err.Number
230       On Error GoTo 0

          'Destroy auxiliar buffer
240       Set buffer = Nothing

250       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GuildVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildVote(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim vote As String
              Dim errorStr As String

90            vote = buffer.ReadASCIIString()

100           If Not modGuilds.v_UsuarioVota(userindex, vote, errorStr) Then
110               Call WriteConsoleMsg(userindex, "Voto NO contabilizado: " & errorStr, FontTypeNames.FONTTYPE_GUILD)
120           Else
130               Call WriteConsoleMsg(userindex, "Voto contabilizado.", FontTypeNames.FONTTYPE_GUILD)
140           End If

              'If we got here then packet is complete, copy data back to original queue
150           Call .incomingData.CopyBuffer(buffer)
160       End With

Errhandler:
          Dim error  As Long
170       error = Err.Number
180       On Error GoTo 0

          'Destroy auxiliar buffer
190       Set buffer = Nothing

200       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "ShowGuildNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowGuildNews(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMA
      'Last Modification: 05/17/06
      '
      '***************************************************

10        With UserList(userindex)

              'Remove packet ID
20            Call .incomingData.ReadByte

30            Call modGuilds.SendGuildNews(userindex)
40        End With
End Sub

''
' Handles the "Punishments" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePunishments(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 25/08/2009
      '25/08/2009: ZaMa - Now only admins can see other admins' punishment list
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim Name As String
              Dim Count As Integer

90            Name = buffer.ReadASCIIString()

100           If LenB(Name) <> 0 Then
110               If (InStrB(Name, "\") <> 0) Then
120                   Name = Replace(Name, "\", "")
130               End If
140               If (InStrB(Name, "/") <> 0) Then
150                   Name = Replace(Name, "/", "")
160               End If
170               If (InStrB(Name, ":") <> 0) Then
180                   Name = Replace(Name, ":", "")
190               End If
200               If (InStrB(Name, "|") <> 0) Then
210                   Name = Replace(Name, "|", "")
220               End If


230               If FileExist(CharPath & Name & ".chr", vbNormal) Then
240                   Count = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
250                   If Count = 0 Then
260                       Call WriteConsoleMsg(userindex, "No tienes penas.", FontTypeNames.FONTTYPE_INFO)
270                   Else
280                       While Count > 0
290                           Call WriteConsoleMsg(userindex, Count & " - " & GetVar(CharPath & Name & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
300                           Count = Count - 1
310                       Wend
320                   End If


330               End If
340           End If

              'If we got here then packet is complete, copy data back to original queue
350           Call .incomingData.CopyBuffer(buffer)
360       End With

Errhandler:
          Dim error  As Long
370       error = Err.Number
380       On Error GoTo 0

          'Destroy auxiliar buffer
390       Set buffer = Nothing

400       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "ChangePassword" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangePassword(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Creation Date: 10/10/07
      'Last Modified By: Rapsodius
      '***************************************************
#If SeguridadAlkon Then
10        If UserList(userindex).incomingData.length < 65 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
#Else
50        If UserList(userindex).incomingData.length < 5 Then
60            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
70            Exit Sub
80        End If
#End If
          
90    On Error GoTo Errhandler
100       With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
110           Call buffer.CopyBuffer(.incomingData)
              
              Dim oldPass As String
              Dim newPass As String
              Dim oldPass2 As String
              
              'Remove packet ID
120           Call buffer.ReadByte
              
#If SeguridadAlkon Then
130           oldPass = UCase$(buffer.ReadASCIIStringFixed(32))
140           newPass = UCase$(buffer.ReadASCIIStringFixed(32))
#Else
150           oldPass = UCase$(buffer.ReadASCIIString())
160           newPass = UCase$(buffer.ReadASCIIString())
#End If
              
170           If LenB(newPass) = 0 Then
180               Call WriteConsoleMsg(userindex, "Debes especificar una contraseña nueva, inténtalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
190           Else
200               oldPass2 = UCase$(GetVar(CharPath & UserList(userindex).Name & ".chr", "INIT", "Password"))
                  
210               If oldPass2 <> oldPass Then
220                   Call WriteConsoleMsg(userindex, "La contraseña actual proporcionada no es correcta. La contraseña no ha sido cambiada, inténtalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
230               Else
240                   Call WriteVar(CharPath & UserList(userindex).Name & ".chr", "INIT", "Password", newPass)
250                   Call WriteConsoleMsg(userindex, "La contraseña fue cambiada con éxito.", FontTypeNames.FONTTYPE_INFO)
260               End If
270           End If
              
              'If we got here then packet is complete, copy data back to original queue
280           Call .incomingData.CopyBuffer(buffer)
290       End With
          
Errhandler:
          Dim error As Long
300       error = Err.Number
310   On Error GoTo 0
          
          'Destroy auxiliar buffer
320       Set buffer = Nothing
          
330       If error <> 0 Then _
              Err.Raise error
End Sub
Private Sub HandleChangePin(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Creation Date: 10/10/07
      'Last Modified By: Rapsodius
      '***************************************************
    #If SeguridadAlkon Then
10            If UserList(userindex).incomingData.length < 65 Then
20                Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30                Exit Sub
40            End If
    #Else
50            If UserList(userindex).incomingData.length < 5 Then
60                Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
70                Exit Sub
80            End If
    #End If

90        On Error GoTo Errhandler
100       With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
110           Call buffer.CopyBuffer(.incomingData)

              Dim oldPin As String
              Dim newPin As String
              Dim oldPin2 As String

              'Remove packet ID
120           Call buffer.ReadByte

        #If SeguridadAlkon Then
130               oldPin = UCase$(buffer.ReadASCIIStringFixed(32))
140               newPin = UCase$(buffer.ReadASCIIStringFixed(32))
        #Else
150               oldPin = UCase$(buffer.ReadASCIIString())
160               newPin = UCase$(buffer.ReadASCIIString())
        #End If

170           If LenB(newPin) = 0 Then
180               Call WriteConsoleMsg(userindex, "Debes especificar una nueva clave PIN, inténtalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
190           Else
200               oldPin2 = UCase$(GetVar(CharPath & UserList(userindex).Name & ".chr", "INIT", "PIN"))

210               If oldPin2 <> oldPin Then
220                   Call WriteConsoleMsg(userindex, "La clave Pin proporcionada no es correcta. La clave Pin no ha sido cambiada, inténtalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
230               Else
240                   Call WriteVar(CharPath & UserList(userindex).Name & ".chr", "INIT", "PIN", newPin)
250                   Call WriteConsoleMsg(userindex, "La clave Pin fue cambiada con éxito.", FontTypeNames.FONTTYPE_INFO)
260               End If
270           End If

              'If we got here then packet is complete, copy data back to original queue
280           Call .incomingData.CopyBuffer(buffer)
290       End With

Errhandler:
          Dim error  As Long
300       error = Err.Number
310       On Error GoTo 0

          'Destroy auxiliar buffer
320       Set buffer = Nothing

330       If error <> 0 Then _
             Err.Raise error
End Sub



''
' Handles the "Gamble" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGamble(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim Amount As Integer

70            Amount = .incomingData.ReadInteger()

80            If .flags.Muerto = 1 Then
90                Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
100           ElseIf .flags.TargetNPC = 0 Then
                  'Validate target NPC
110               Call WriteConsoleMsg(userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
120           ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
130               Call WriteShortMsj(userindex, 8, FontTypeNames.FONTTYPE_INFO)
140           ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
150               Call WriteChatOverHead(userindex, "No tengo ningún interés en apostar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
160           ElseIf Amount < 1 Then
170               Call WriteChatOverHead(userindex, "El mínimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
180           ElseIf Amount > 5000 Then
190               Call WriteChatOverHead(userindex, "El máximo de apuesta es 5000 monedas.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
200           ElseIf .Stats.Gld < Amount Then
210               Call WriteChatOverHead(userindex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
220           Else
230               If RandomNumber(1, 100) <= 47 Then
240                   .Stats.Gld = .Stats.Gld + Amount
250                   Call WriteChatOverHead(userindex, "¡Felicidades! Has ganado " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

260                   Apuestas.Perdidas = Apuestas.Perdidas + Amount
270                   Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
280               Else
290                   .Stats.Gld = .Stats.Gld - Amount
300                   Call WriteChatOverHead(userindex, "Lo siento, has perdido " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

310                   Apuestas.Ganancias = Apuestas.Ganancias + Amount
320                   Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
330               End If

340               Apuestas.Jugadas = Apuestas.Jugadas + 1

350               Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))

360               Call WriteUpdateGold(userindex)
370           End If
380       End With
End Sub

''
' Handles the "InquiryVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiryVote(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim opt As Byte

70            opt = .incomingData.ReadByte()

80            Call WriteConsoleMsg(userindex, ConsultaPopular.doVotar(userindex, opt), FontTypeNames.FONTTYPE_GUILD)
90        End With
End Sub
Private Sub HandleBankExtractGold(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim Amount As Long

70            Amount = .incomingData.ReadLong()

              'Dead people can't leave a faction.. they can't talk...
80            If .flags.Muerto = 1 Then
90                Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
100               Exit Sub
110           End If

              'Validate target NPC
120           If .flags.TargetNPC = 0 Then
130               Call WriteConsoleMsg(userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
140               Exit Sub
150           End If

160           If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub

170           If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
180               Call WriteShortMsj(userindex, 8, FontTypeNames.FONTTYPE_INFO)
190               Exit Sub
200           End If

              Dim Monea As Long
210           Monea = .Stats.Banco

220           If Amount > 0 And Amount <= .Stats.Banco Then
230               .Stats.Banco = .Stats.Banco - Amount
240               .Stats.Gld = .Stats.Gld + Amount
250               Call WriteChatOverHead(userindex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
260           Else
270               .Stats.Gld = .Stats.Gld + .Stats.Banco
280               .Stats.Banco = 0
290               Call WriteChatOverHead(userindex, "Has retirado " & Monea & " monedas de oro de tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
300           End If

310           Call WriteUpdateGold(userindex)
320           Call WriteUpdateBankGold(userindex)
330       End With
End Sub

''
' Handles the "LeaveFaction" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeaveFaction(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************

          Dim TalkToKing As Boolean
          Dim TalkToDemon As Boolean
          Dim NpcIndex As Integer

10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              'Dead people can't leave a faction.. they can't talk...
30            If .flags.Muerto = 1 Then
40                Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
50                Exit Sub
60            End If

              ' Chequea si habla con el rey o el demonio. Puede salir sin hacerlo, pero si lo hace le reponden los npcs
70            NpcIndex = .flags.TargetNPC
80            If NpcIndex <> 0 Then
                  ' Es rey o domonio?
90                If Npclist(NpcIndex).NPCtype = eNPCType.Noble Then
                      'Rey?
100                   If Npclist(NpcIndex).flags.Faccion = 0 Then
110                       TalkToKing = True
                          ' Demonio
120                   Else
130                       TalkToDemon = True
140                   End If
150               End If
160           End If

              'Quit the Royal Army?
170           If .Faccion.ArmadaReal = 1 Then
                  ' Si le pidio al demonio salir de la armada, este le responde.
180               If TalkToDemon Then
190                   Call WriteChatOverHead(userindex, "¡¡¡Sal de aquí bufón!!!", _
                                             Npclist(NpcIndex).Char.CharIndex, vbWhite)

200               Else
                      ' Si le pidio al rey salir de la armada, le responde.
210                   If TalkToKing Then
220                       Call WriteChatOverHead(userindex, "Serás bienvenido a las fuerzas imperiales si deseas regresar.", _
                                                 Npclist(NpcIndex).Char.CharIndex, vbWhite)
230                   End If

240                   Call ExpulsarFaccionReal(userindex, False)

250               End If

                  'Quit the Chaos Legion?
260           ElseIf .Faccion.FuerzasCaos = 1 Then
                  ' Si le pidio al rey salir del caos, le responde.
270               If TalkToKing Then
280                   Call WriteChatOverHead(userindex, "¡¡¡Sal de aquí maldito criminal!!!", _
                                             Npclist(NpcIndex).Char.CharIndex, vbWhite)
290               Else
                      ' Si le pidio al demonio salir del caos, este le responde.
300                   If TalkToDemon Then
310                       Call WriteChatOverHead(userindex, "Ya volverás arrastrandote.", _
                                                 Npclist(NpcIndex).Char.CharIndex, vbWhite)
320                   End If

330                   Call ExpulsarFaccionCaos(userindex, False)
340               End If
                  ' No es faccionario
350           Else

                  ' Si le hablaba al rey o demonio, le repsonden ellos
360               If NpcIndex > 0 Then
370                   Call WriteChatOverHead(userindex, "¡No perteneces a ninguna facción!", _
                                             Npclist(NpcIndex).Char.CharIndex, vbWhite)
380               Else
390                   Call WriteConsoleMsg(userindex, "¡No perteneces a ninguna facción!", FontTypeNames.FONTTYPE_FIGHT)
400               End If

410           End If

420       End With

End Sub
Private Sub HandleBankDepositGold(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim Amount As Long

70            Amount = .incomingData.ReadLong()

              'Dead people can't leave a faction.. they can't talk...
80            If .flags.Muerto = 1 Then
90                Call WriteShortMsj(userindex, 5, FontTypeNames.FONTTYPE_INFO)
100               Exit Sub
110           End If

              'Validate target NPC
120           If .flags.TargetNPC = 0 Then
130               Call WriteConsoleMsg(userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
140               Exit Sub
150           End If

160           If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
170               Call WriteShortMsj(userindex, 8, FontTypeNames.FONTTYPE_INFO)
180               Exit Sub
190           End If

200           If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub

210           If Amount > 0 And Amount <= .Stats.Gld Then
220               .Stats.Banco = .Stats.Banco + Amount
230               .Stats.Gld = .Stats.Gld - Amount
240               Call WriteChatOverHead(userindex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

250               Call WriteUpdateGold(userindex)
260               Call WriteUpdateBankGold(userindex)
270           Else
280               Call WriteChatOverHead(userindex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
290           End If
300       End With
End Sub

''
' Handles the "Denounce" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim Text As String

90            Text = buffer.ReadASCIIString()

100           If .flags.Silenciado = 0 And .Counters.Denuncia = 0 Then
                  'Analize chat...
110               Call Statistics.ParseChat(Text)

120               If UCase$(Left$(Text, 15)) = "[FOTODENUNCIAS]" Then
130                   SendData SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(Text & ". Hecha por: " & .Name, FontTypeNames.FONTTYPE_INFO)
140                   .Counters.Denuncia = 20
150               Else
160                   Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(LCase$(.Name) & " DENUNCIA: " & Text, FontTypeNames.fonttype_dios))
170                   Call WriteConsoleMsg(userindex, "Recuerda que sólo puedes enviar una denuncia cada 30 segundos.", FontTypeNames.fonttype_dios)
180                   Call WriteConsoleMsg(userindex, "Denuncia enviada, pronto será atendido por un Game Master.", FontTypeNames.FONTTYPE_INFO)
190                   .Counters.Denuncia = 30
200               End If
210           End If

              'If we got here then packet is complete, copy data back to original queue
220           Call .incomingData.CopyBuffer(buffer)
230       End With

Errhandler:
          Dim error  As Long
240       error = Err.Number
250       On Error GoTo 0

          'Destroy auxiliar buffer
260       Set buffer = Nothing

270       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GuildFundate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundate(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 14/12/2009
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 1 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
60            Call .incomingData.ReadByte

70            If HasFound(.Name) Then
80                Call WriteConsoleMsg(userindex, "¡Ya has fundado un clan, no puedes fundar otro!", FontTypeNames.FONTTYPE_INFOBOLD)
90                Exit Sub
100           End If

110           Call WriteShowGuildAlign(userindex)
120       End With
End Sub

''
' Handles the "GuildFundation" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundation(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 14/12/2009
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim clanType As eClanType
              Dim error As String

70            clanType = .incomingData.ReadByte()

80            If HasFound(.Name) Then
90                Call WriteConsoleMsg(userindex, "¡Ya has fundado un clan, no puedes fundar otro!", FontTypeNames.FONTTYPE_INFOBOLD)
100               Call LogCheating("El usuario " & .Name & " ha intentado fundar un clan ya habiendo fundado otro desde la IP " & .ip)
110               Exit Sub
120           End If

130           Select Case UCase$(Trim(clanType))
              Case eClanType.ct_RoyalArmy
140               .FundandoGuildAlineacion = ALINEACION_ARMADA
150           Case eClanType.ct_Evil
160               .FundandoGuildAlineacion = ALINEACION_LEGION
170           Case eClanType.ct_Neutral
180               .FundandoGuildAlineacion = ALINEACION_NEUTRO
190           Case eClanType.ct_GM
200               .FundandoGuildAlineacion = ALINEACION_MASTER
210           Case eClanType.ct_Legal
220               .FundandoGuildAlineacion = ALINEACION_CIUDA
230           Case eClanType.ct_Criminal
240               .FundandoGuildAlineacion = ALINEACION_CRIMINAL
250           Case Else
260               Call WriteConsoleMsg(userindex, "Alineación inválida.", FontTypeNames.FONTTYPE_GUILD)
270               Exit Sub
280           End Select

290           If modGuilds.PuedeFundarUnClan(userindex, .FundandoGuildAlineacion, error) Then
300               Call WriteShowGuildFundationForm(userindex)
310           Else
320               .FundandoGuildAlineacion = 0
330               Call WriteConsoleMsg(userindex, error, FontTypeNames.FONTTYPE_GUILD)
340           End If
350       End With
End Sub

''
' Handles the "GuildMemberList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberList(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim guild As String
              Dim memberCount As Integer
              Dim i  As Long
              Dim UserName As String

90            guild = buffer.ReadASCIIString()

100           If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
110               If (InStrB(guild, "\") <> 0) Then
120                   guild = Replace(guild, "\", "")
130               End If
140               If (InStrB(guild, "/") <> 0) Then
150                   guild = Replace(guild, "/", "")
160               End If

170               If Not FileExist(App.Path & "\guilds\" & guild & "-members.mem") Then
180                   Call WriteConsoleMsg(userindex, "No existe el clan: " & guild, FontTypeNames.FONTTYPE_INFO)
190               Else
200                   memberCount = val(GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "INIT", "NroMembers"))

210                   For i = 1 To memberCount
220                       UserName = GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "Members", "Member" & i)

230                       Call WriteConsoleMsg(userindex, UserName & "<" & guild & ">", FontTypeNames.FONTTYPE_INFO)
240                   Next i
250               End If
260           End If

              'If we got here then packet is complete, copy data back to original queue
270           Call .incomingData.CopyBuffer(buffer)
280       End With

Errhandler:
          Dim error  As Long
290       error = Err.Number
300       On Error GoTo 0

          'Destroy auxiliar buffer
310       Set buffer = Nothing

320       If error <> 0 Then _
             Err.Raise error
End Sub

Public Function SearchSlotInvasion(ByVal UserMap As Integer) As Integer
    Dim LoopC As Integer
    
    SearchSlotInvasion = -1
    
    For LoopC = 1 To NumInvasiones
        If Invasiones(LoopC).Activa Then
            If Invasiones(LoopC).map = UserMap Then
                SearchSlotInvasion = LoopC
                Exit For
            End If
        End If
    Next LoopC
End Function
Private Sub HandleTerminateInvasion(ByVal userindex As Integer)
    With UserList(userindex)
        Call .incomingData.ReadByte
        
        Dim InvasionIndex As Integer
        
        
        If EsGm(userindex) Then
            InvasionIndex = SearchSlotInvasion(.Pos.map)
        
            If Not InvasionIndex <= 0 Then
                mInvasiones.CloseInvasion InvasionIndex
            End If
        End If
        
    End With
End Sub
Private Sub HandleCreateInvasion(ByVal userindex As Integer)
      '***************************************************
      'Author: Lautaro
      'Created: 06/09/2018
      '***************************************************
      
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim Name As String
              Dim desc As String
              Dim InvasionIndex As Byte
              Dim DropIndex As Byte
              Dim map As Integer
              
              Name = buffer.ReadASCIIString
              desc = buffer.ReadASCIIString
              InvasionIndex = buffer.ReadByte
              map = buffer.ReadInteger
              
100           Call .incomingData.CopyBuffer(buffer)

               If EsGm(userindex) Then
                   mInvasiones.NewInvasion userindex, InvasionIndex, Name, desc, map
               End If

180       End With

Errhandler:
          Dim error  As Long
190       error = Err.Number
200       On Error GoTo 0

          'Destroy auxiliar buffer
210       Set buffer = Nothing

220       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GMMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMMessage(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 01/08/07
      'Last Modification by: (liquid)
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim message As String

90            message = buffer.ReadASCIIString()

100           Call .incomingData.CopyBuffer(buffer)

110           If Not .flags.Privilegios And PlayerType.User Then
120               Call LogGM(.Name, "Mensaje a Gms:" & message)

130               If LenB(message) <> 0 Then
                      'Analize chat...
140                   Call Statistics.ParseChat(message)

150                   Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & "> " & message, FontTypeNames.FONTTYPE_GMMSG))
160               End If
170           End If


180       End With

Errhandler:
          Dim error  As Long
190       error = Err.Number
200       On Error GoTo 0

          'Destroy auxiliar buffer
210       Set buffer = Nothing

220       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "ShowName" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
40                .showName = Not .showName    'Show / Hide the name

50                Call RefreshCharStatus(userindex)
60            End If
70        End With
End Sub

''
' Handles the "OnlineRoyalArmy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineRoyalArmy(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 28/05/2010
      '28/05/2010: ZaMa - Ahora solo dioses pueden ver otros dioses online.
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            .incomingData.ReadByte

30            If .flags.Privilegios And PlayerType.User Then Exit Sub

              Dim i  As Long
              Dim list As String
              Dim priv As PlayerType

40            priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoyalCouncil
              
              ' Solo dioses pueden ver otros dioses online
50            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
60                priv = priv Or PlayerType.Dios Or PlayerType.Admin
70            End If

80            For i = 1 To LastUser
90                If UserList(i).ConnID <> -1 Then
100                   If UserList(i).Faccion.ArmadaReal = 1 Then
110                       If UserList(i).flags.Privilegios And priv Then
120                           list = list & UserList(i).Name & ", "
130                       End If
140                   End If
150               End If
160           Next i
170       End With

180       If Len(list) > 0 Then
190           Call WriteConsoleMsg(userindex, "Reales conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
200       Else
210           Call WriteConsoleMsg(userindex, "No hay reales conectados.", FontTypeNames.FONTTYPE_INFO)
220       End If
End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineChaosLegion(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 28/05/2010
      '28/05/2010: ZaMa - Ahora solo dioses pueden ver otros dioses online.
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            .incomingData.ReadByte

30            If .flags.Privilegios And PlayerType.User Then Exit Sub

              Dim i  As Long
              Dim list As String
              Dim priv As PlayerType

40            priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.ChaosCouncil

              ' Solo dioses pueden ver otros dioses online
50            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
60                priv = priv Or PlayerType.Dios Or PlayerType.Admin
70            End If

80            For i = 1 To LastUser
90                If UserList(i).ConnID <> -1 Then
100                   If UserList(i).Faccion.FuerzasCaos = 1 Then
110                       If UserList(i).flags.Privilegios And priv Then
120                           list = list & UserList(i).Name & ", "
130                       End If
140                   End If
150               End If
160           Next i
170       End With

180       If Len(list) > 0 Then
190           Call WriteConsoleMsg(userindex, "Caos conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
200       Else
210           Call WriteConsoleMsg(userindex, "No hay Caos conectados.", FontTypeNames.FONTTYPE_INFO)
220       End If
End Sub

''
' Handles the "GoNearby" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoNearby(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 01/10/07
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String

90            UserName = buffer.ReadASCIIString()

              Dim tIndex As Integer
              Dim X  As Long
              Dim Y  As Long
              Dim i  As Long
              Dim Found As Boolean

100           tIndex = NameIndex(UserName)

              'Check the user has enough powers
110           If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
                  If Not Protocol.IsNickEspecial(UserName) Then
                  'If Not StrComp(UCase$(UserName), "THYRAH") = 0 And Not StrComp(UCase$(UserName), "WAITING") = 0 And Not StrComp(UCase$(UserName), "LAUTARO") = 0 Then
                      'Si es dios o Admins no podemos salvo que nosotros también lo seamos
130                   If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
140                       If tIndex <= 0 Then    'existe el usuario destino?
150                           Call WriteConsoleMsg(userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
160                       Else
170                           For i = 2 To 5    'esto for sirve ir cambiando la distancia destino
180                               For X = UserList(tIndex).Pos.X - i To UserList(tIndex).Pos.X + i
190                                   For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i
200                                       If MapData(UserList(tIndex).Pos.map, X, Y).userindex = 0 Then
210                                           If LegalPos(UserList(tIndex).Pos.map, X, Y, True, True) Then
220                                               Call WarpUserChar(userindex, UserList(tIndex).Pos.map, X, Y, True)
230                                               Call LogGM(.Name, "/IRCERCA " & UserName & " Mapa:" & UserList(tIndex).Pos.map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y)
240                                               Found = True
250                                               Exit For
260                                           End If
270                                       End If
280                                   Next Y
          
290                                   If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
300                               Next X
          
310                               If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
320                           Next i
          
                              'No space found??
330                           If Not Found Then
340                               Call WriteConsoleMsg(userindex, "Todos los lugares están ocupados.", FontTypeNames.FONTTYPE_INFO)
350                           End If
360                       End If
370                   End If
380               End If
390           End If

              'If we got here then packet is complete, copy data back to original queue
400           Call .incomingData.CopyBuffer(buffer)
410       End With

Errhandler:
          Dim error  As Long
420       error = Err.Number
430       On Error GoTo 0

          'Destroy auxiliar buffer
440       Set buffer = Nothing

450       If error <> 0 Then _
             Err.Raise error
End Sub
Private Sub Elmasbuscado(ByVal userindex As String)

10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte
              Dim UserName As String
90            UserName = buffer.ReadASCIIString()
              Dim tIndex As String
100           tIndex = NameIndex(UserName)

110           If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios) Then
120               If tIndex <= 0 Then    'usuario Offline
130                   Call WriteConsoleMsg(userindex, "Usuario Offline.", FontTypeNames.FONTTYPE_INFO)
140               Else
150                   If UserList(tIndex).flags.Muerto = 1 Then    'tu enemigo esta muerto
160                       Call WriteConsoleMsg(userindex, "El usuario que queres que sea buscado esta muerto.", FontTypeNames.FONTTYPE_INFO)
170                   Else
                      
180                       If UserList(tIndex).Pos.map = 201 Then
190                           Call WriteConsoleMsg(userindex, "Esta ocupado en un reto.", FontTypeNames.FONTTYPE_INFO)
200                       Else
210                           Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Atencion!!: Se Busca el usuario " & UserList(tIndex).Name & ", el que lo asesine tendra su recompensa.", FontTypeNames.FONTTYPE_GUILD))
220                           Call WriteConsoleMsg(tIndex, "Tu eres el usuario más buscado, ten cuidado!!.", FontTypeNames.FONTTYPE_GUILD)
230                           ElmasbuscadoFusion = UserList(tIndex).Name
240                       End If
250                   End If
              
260               End If
270           End If

              'If we got here then packet is complete, copy data back to original queue
280           Call .incomingData.CopyBuffer(buffer)
290       End With

Errhandler:
          Dim error  As Long
300       error = Err.Number
310       On Error GoTo 0

          'Destroy auxiliar buffer
320       Set buffer = Nothing

330       If error <> 0 Then _
             Err.Raise error

340       Exit Sub
End Sub


''
' Handles the "Comment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleComment(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim comment As String
90            comment = buffer.ReadASCIIString()

100           If Not .flags.Privilegios And PlayerType.User Then
110               Call LogGM(.Name, "Comentario: " & comment)
120               Call WriteConsoleMsg(userindex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)
130           End If

              'If we got here then packet is complete, copy data back to original queue
140           Call .incomingData.CopyBuffer(buffer)
150       End With

Errhandler:
          Dim error  As Long
160       error = Err.Number
170       On Error GoTo 0

          'Destroy auxiliar buffer
180       Set buffer = Nothing

190       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "ServerTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 01/08/07
      'Last Modification by: (liquid)
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And PlayerType.User Then Exit Sub

40            Call LogGM(.Name, "Hora.")
50        End With

60        Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & time & " " & Date, FontTypeNames.FONTTYPE_INFO))
End Sub

''
' Handles the "Where" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 18/11/2010
      '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
      '18/11/2010: ZaMa - Obtengo los privs del charfile antes de mostrar la posicion de un usuario offline.
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
70            Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
80            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
90            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer
              Dim miPos As String

100           UserName = buffer.ReadASCIIString()
110           tUser = NameIndex(UserName)
                  
120           If EsGm(userindex) Then
130               If tUser <= 0 Then
                      If Not Protocol.IsNickEspecial(UserName) Then
150                       If FileExist(CharPath & UserName & ".chr", vbNormal) Then
160                           miPos = GetVar(CharPath & UserName & ".chr", "INIT", "POSITION")
170                           Call WriteConsoleMsg(userindex, "Ubicación  " & UserName & " (Offline): " & ReadField(1, miPos, 45) & ", " & ReadField(2, miPos, 45) & ", " & ReadField(3, miPos, 45) & ".", FontTypeNames.FONTTYPE_INFO)
180                       End If
190                   End If
200               Else
                      If Not Protocol.IsNickEspecial(UserList(tUser).Name) Then
                      'If Not StrComp(UCase$(UserList(tUser).Name), "THYRAH") = 0 And Not StrComp(UCase$(UserList(tUser).Name), "WAITING") = 0 And Not StrComp(UCase$(UserList(tUser).Name), "LAUTARO") = 0 Then
220                       Call WriteConsoleMsg(userindex, "Ubicación  " & UserName & ": " & UserList(tUser).Pos.map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
230                   End If
240               End If
250           End If
              
              'If we got here then packet is complete, copy data back to original queue
260           Call .incomingData.CopyBuffer(buffer)
270       End With

Errhandler:
          Dim error  As Long
280       error = Err.Number
290       On Error GoTo 0

          'Destroy auxiliar buffer
300       Set buffer = Nothing

310       If error <> 0 Then _
             Err.Raise error
End Sub
''
' Handles the "CreaturesInMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 30/07/06
      'Pablo (ToxicWaste): modificaciones generales para simplificar la visualización.
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim map As Integer
              Dim i, j As Long
              Dim NPCcount1, NPCcount2 As Integer
              Dim NPCcant1() As Integer
              Dim NPCcant2() As Integer
              Dim List1() As String
              Dim List2() As String

70            map = .incomingData.ReadInteger()

80            If .flags.Privilegios And PlayerType.User Then Exit Sub

90            If MapaValido(map) Then
100               For i = 1 To LastNPC
                      'VB isn't lazzy, so we put more restrictive condition first to speed up the process
110                   If Npclist(i).Pos.map = map Then
                          '¿esta vivo?
120                       If Npclist(i).flags.NPCActive And Npclist(i).Hostile = 1 And Npclist(i).Stats.Alineacion = 2 Then
130                           If NPCcount1 = 0 Then
140                               ReDim List1(0) As String
150                               ReDim NPCcant1(0) As Integer
160                               NPCcount1 = 1
170                               List1(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
180                               NPCcant1(0) = 1
190                           Else
200                               For j = 0 To NPCcount1 - 1
210                                   If Left$(List1(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
220                                       List1(j) = List1(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
230                                       NPCcant1(j) = NPCcant1(j) + 1
240                                       Exit For
250                                   End If
260                               Next j
270                               If j = NPCcount1 Then
280                                   ReDim Preserve List1(0 To NPCcount1) As String
290                                   ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
300                                   NPCcount1 = NPCcount1 + 1
310                                   List1(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
320                                   NPCcant1(j) = 1
330                               End If
340                           End If
350                       Else
360                           If NPCcount2 = 0 Then
370                               ReDim List2(0) As String
380                               ReDim NPCcant2(0) As Integer
390                               NPCcount2 = 1
400                               List2(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
410                               NPCcant2(0) = 1
420                           Else
430                               For j = 0 To NPCcount2 - 1
440                                   If Left$(List2(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
450                                       List2(j) = List2(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
460                                       NPCcant2(j) = NPCcant2(j) + 1
470                                       Exit For
480                                   End If
490                               Next j
500                               If j = NPCcount2 Then
510                                   ReDim Preserve List2(0 To NPCcount2) As String
520                                   ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
530                                   NPCcount2 = NPCcount2 + 1
540                                   List2(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
550                                   NPCcant2(j) = 1
560                               End If
570                           End If
580                       End If
590                   End If
600               Next i

610               Call WriteConsoleMsg(userindex, "Npcs Hostiles en mapa: ", FontTypeNames.FONTTYPE_WARNING)
620               If NPCcount1 = 0 Then
630                   Call WriteConsoleMsg(userindex, "No hay NPCS Hostiles.", FontTypeNames.FONTTYPE_INFO)
640               Else
650                   For j = 0 To NPCcount1 - 1
660                       Call WriteConsoleMsg(userindex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
670                   Next j
680               End If
690               Call WriteConsoleMsg(userindex, "Otros Npcs en mapa: ", FontTypeNames.FONTTYPE_WARNING)
700               If NPCcount2 = 0 Then
710                   Call WriteConsoleMsg(userindex, "No hay más NPCS.", FontTypeNames.FONTTYPE_INFO)
720               Else
730                   For j = 0 To NPCcount2 - 1
740                       Call WriteConsoleMsg(userindex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
750                   Next j
760               End If
770               Call LogGM(.Name, "Numero enemigos en mapa " & map)
780           End If
790       End With
End Sub

''
' Handles the "WarpMeToTarget" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpMeToTarget(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 26/03/09
      '26/03/06: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              Dim X  As Integer
              Dim Y  As Integer

30            If .flags.Privilegios And PlayerType.User Then Exit Sub

40            X = .flags.TargetX
50            Y = .flags.TargetY

60            Call FindLegalPos(userindex, .flags.TargetMap, X, Y)
70            Call WarpUserChar(userindex, .flags.TargetMap, X, Y, True)
80            Call LogGM(.Name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.map)
90        End With
End Sub

''
' Handles the "WarpChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 26/03/2009
      '26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
      '***************************************************
10        If UserList(userindex).incomingData.length < 7 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim map As Integer
              Dim X  As Integer
              Dim Y  As Integer
              Dim tUser As Integer

90            UserName = buffer.ReadASCIIString()
100           map = buffer.ReadInteger()
110           X = buffer.ReadByte()
120           Y = buffer.ReadByte()

130           If Not .flags.Privilegios And PlayerType.User Then
140               If MapaValido(map) And LenB(UserName) <> 0 Then
150                   If UCase$(UserName) <> "YO" Then
160                       If Not .flags.Privilegios And PlayerType.Consejero Then
170                           tUser = NameIndex(UserName)
180                       End If
190                   Else
200                       tUser = userindex
210                   End If

220                   If tUser <= 0 Then
230                       Call WriteVar(CharPath & UserName & ".chr", "INIT", "Position", map & "-" & X & "-" & Y)
240                       Call WriteConsoleMsg(userindex, "Charfile modificado", FontTypeNames.FONTTYPE_GM)
250                   ElseIf InMapBounds(map, X, Y) Then
260                       If Not MapData(map, X, Y).TileExit.map > 0 Then
270                           Call FindLegalPos(tUser, map, X, Y)
280                           Call WarpUserChar(tUser, map, X, Y, True, True)
290                           If tUser <> userindex Then Call WriteConsoleMsg(userindex, UserList(tUser).Name & " transportado.", FontTypeNames.FONTTYPE_INFO)
300                           Call LogGM(.Name, "Transportó a " & UserList(tUser).Name & " hacia " & "Mapa" & map & " X:" & X & " Y:" & Y)
310                       Else
320                           WriteConsoleMsg userindex, "No puedes teletransportarte sobre un teleport.", FontTypeNames.FONTTYPE_INFO
330                       End If
                          
                          
340                   End If
350               End If
360           End If

              'If we got here then packet is complete, copy data back to original queue
370           Call .incomingData.CopyBuffer(buffer)
380       End With

Errhandler:
          Dim error  As Long
390       error = Err.Number
400       On Error GoTo 0

          'Destroy auxiliar buffer
410       Set buffer = Nothing

420       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "Silence" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSilence(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer

90            UserName = buffer.ReadASCIIString()

100           If Not .flags.Privilegios And PlayerType.User Then
110               tUser = NameIndex(UserName)

120               If tUser <= 0 Then
130                   Call WriteConsoleMsg(userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
140               Else
150                   If UserList(tUser).flags.Silenciado = 0 Then
160                       UserList(tUser).flags.Silenciado = 1
170                       Call WriteConsoleMsg(userindex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
180                       Call WriteShowMessageBox(tUser, "Estimado usuario, ud. ha sido silenciado por los administradores. Sus denuncias serán ignoradas por el servidor de aquí en más. Utilice /GM para contactar un administrador.")
190                       Call LogGM(.Name, "/silenciar " & UserList(tUser).Name)

                          'Flush the other user's buffer
200                       Call FlushBuffer(tUser)
210                   Else
220                       UserList(tUser).flags.Silenciado = 0
230                       Call WriteConsoleMsg(userindex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
240                       Call LogGM(.Name, "/DESsilenciar " & UserList(tUser).Name)
250                   End If
260               End If
270           End If

              'If we got here then packet is complete, copy data back to original queue
280           Call .incomingData.CopyBuffer(buffer)
290       End With

Errhandler:
          Dim error  As Long
300       error = Err.Number
310       On Error GoTo 0

          'Destroy auxiliar buffer
320       Set buffer = Nothing

330       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "SOSShowList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSShowList(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And PlayerType.User Then Exit Sub
40            Call WriteShowSOSForm(userindex)
50        End With
End Sub

''
' Handles the "ItemUpgrade" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleItemUpgrade(ByVal userindex As Integer)
      '***************************************************
      'Author: Torres Patricio
      'Last Modification: 12/09/09
      '
      '***************************************************

10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50        With UserList(userindex)
              Dim ItemIndex As Integer

              'Remove packet ID
60            Call .incomingData.ReadByte

70            ItemIndex = .incomingData.ReadInteger()

80            If ItemIndex <= 0 Then Exit Sub
90            If Not TieneObjetos(ItemIndex, 1, userindex) Then Exit Sub

100           Call DoUpgrade(userindex, ItemIndex)
110       End With
End Sub

''
' Handles the "SOSRemove" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSRemove(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
90            UserName = buffer.ReadASCIIString()

100           If Not .flags.Privilegios And PlayerType.User Then _
                 Call Ayuda.Quitar(UserName)

              'If we got here then packet is complete, copy data back to original queue
110           Call .incomingData.CopyBuffer(buffer)
120       End With

Errhandler:
          Dim error  As Long
130       error = Err.Number
140       On Error GoTo 0

          'Destroy auxiliar buffer
150       Set buffer = Nothing

160       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "GoToChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 26/03/2009
      '26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer
              Dim X  As Integer
              Dim Y  As Integer

90            UserName = buffer.ReadASCIIString()
100           tUser = NameIndex(UserName)

110           If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Consejero) Then
                  'Si es dios o Admins no podemos salvo que nosotros también lo seamos
120               If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                      If Not Protocol.IsNickEspecial(UserName) Then
140                       If tUser <= 0 Then
150                           Call WriteConsoleMsg(userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
160                       Else
170                           If Not UserList(tUser).Pos.map = 290 Then
180                               X = UserList(tUser).Pos.X
190                               Y = UserList(tUser).Pos.Y + 1
200                               Call FindLegalPos(userindex, UserList(tUser).Pos.map, X, Y)
          
210                               Call WarpUserChar(userindex, UserList(tUser).Pos.map, X, Y, True)
          
220                               If .flags.AdminInvisible = 0 Then
                                      'Call WriteConsoleMsg(tUser, " sientes una presencia cerca de ti.", FontTypeNames.FONTTYPE_INFO)
230                                   Call FlushBuffer(tUser)
240                               End If
          
250                               Call LogGM(.Name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y)
260                           End If
270                       End If
280                   End If
290               End If
300           End If

              'If we got here then packet is complete, copy data back to original queue
310           Call .incomingData.CopyBuffer(buffer)
320       End With

Errhandler:
          Dim error  As Long
330       error = Err.Number
340       On Error GoTo 0

          'Destroy auxiliar buffer
350       Set buffer = Nothing

360       If error <> 0 Then _
             Err.Raise error
End Sub
''
' Handles the "Invisible" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And PlayerType.User Then Exit Sub

40            Call DoAdminInvisible(userindex)
50            Call LogGM(.Name, "/INVISIBLE")
60        End With
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And PlayerType.User Then Exit Sub

40            Call WriteShowGMPanelForm(userindex)
50        End With
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 01/09/07
      'Last modified by: Lucas Tavolaro Ortiz (Tavo)
      'I haven`t found a solution to split, so i make an array of names
      '***************************************************
          Dim i      As Long
          Dim names() As String
          Dim Count  As Long

10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub

40            ReDim names(1 To LastUser) As String
50            Count = 1

60            For i = 1 To LastUser
70                If (LenB(UserList(i).Name) <> 0) Then
80                    If UserList(i).flags.Privilegios And PlayerType.User Then
90                        names(Count) = UserList(i).Name
100                       Count = Count + 1
110                   End If
120               End If
130           Next i

140           If Count > 1 Then Call WriteUserNameList(userindex, names(), Count - 1)
150       End With
End Sub

''
' Handles the "Working" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorking(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
          Dim i      As Long
          Dim Users  As String

10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub

40            For i = 1 To LastUser
50                If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
60                    Users = Users & ", " & UserList(i).Name

                      ' Display the user being checked by the centinel
70                    If modCentinela.Centinela.RevisandoUserIndex = i Then _
                         Users = Users & " (*)"
80                End If
90            Next i

100           If LenB(Users) <> 0 Then
110               Users = Right$(Users, Len(Users) - 2)
120               Call WriteConsoleMsg(userindex, "Usuarios trabajando: " & Users, FontTypeNames.FONTTYPE_INFO)
130           Else
140               Call WriteConsoleMsg(userindex, "No hay usuarios trabajando.", FontTypeNames.FONTTYPE_INFO)
150           End If
160       End With
End Sub

''
' Handles the "Hiding" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHiding(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
          Dim i      As Long
          Dim Users  As String

10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub

40            For i = 1 To LastUser
50                If (LenB(UserList(i).Name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
60                    Users = Users & UserList(i).Name & ", "
70                End If
80            Next i

90            If LenB(Users) <> 0 Then
100               Users = Left$(Users, Len(Users) - 2)
110               Call WriteConsoleMsg(userindex, "Usuarios ocultandose: " & Users, FontTypeNames.FONTTYPE_INFO)
120           Else
130               Call WriteConsoleMsg(userindex, "No hay usuarios ocultandose.", FontTypeNames.FONTTYPE_INFO)
140           End If
150       End With
End Sub

''
' Handles the "Jail" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleJail(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 6 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim Reason As String
              Dim jailTime As Byte
              Dim Count As Byte
              Dim tUser As Integer

90            UserName = buffer.ReadASCIIString()
100           Reason = buffer.ReadASCIIString()
110           jailTime = buffer.ReadByte()

120           If InStr(1, UserName, "+") Then
130               UserName = Replace(UserName, "+", " ")
140           End If

              '/carcel nick@motivo@<tiempo>
150           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
160               If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
170                   Call WriteConsoleMsg(userindex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
180               Else
190                   tUser = NameIndex(UserName)

200                   If tUser <= 0 Then
210                       Call WriteConsoleMsg(userindex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
220                   Else
230                       If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
240                           Call WriteConsoleMsg(userindex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
250                       ElseIf jailTime > 60 Then
260                           Call WriteConsoleMsg(userindex, "No puedés encarcelar por más de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
270                       Else
280                           If (InStrB(UserName, "\") <> 0) Then
290                               UserName = Replace(UserName, "\", "")
300                           End If
310                           If (InStrB(UserName, "/") <> 0) Then
320                               UserName = Replace(UserName, "/", "")
330                           End If

340                           If FileExist(CharPath & UserName & ".chr", vbNormal) Then
350                               Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
360                               Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
370                               Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(Reason) & " " & Date & " " & time)
380                           End If

390                           Call Encarcelar(tUser, jailTime, .Name)
400                           Call LogGM(.Name, " encarceló a " & UserName)
410                       End If
420                   End If
430               End If
440           End If

              'If we got here then packet is complete, copy data back to original queue
450           Call .incomingData.CopyBuffer(buffer)
460       End With

Errhandler:
          Dim error  As Long
470       error = Err.Number
480       On Error GoTo 0

          'Destroy auxiliar buffer
490       Set buffer = Nothing

500       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "KillNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 04/22/08 (NicoNZ)
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And PlayerType.User Then Exit Sub

              Dim tNpc As Integer
              Dim auxNPC As Npc

              'Los consejeros no pueden RMATAr a nada en el mapa pretoriano
40            If .flags.Privilegios And PlayerType.Consejero Then
50                If .Pos.map = MAPA_PRETORIANO Then
60                    Call WriteConsoleMsg(userindex, "Los consejeros no pueden usar este comando en el mapa pretoriano.", FontTypeNames.FONTTYPE_INFO)
70                    Exit Sub
80                End If
90            End If

100           tNpc = .flags.TargetNPC

110           If tNpc > 0 Then
120               Call WriteConsoleMsg(userindex, "RMatas (con posible respawn) a: " & Npclist(tNpc).Name, FontTypeNames.FONTTYPE_INFO)

130               auxNPC = Npclist(tNpc)
140               Call QuitarNPC(tNpc)
150               Call ReSpawnNpc(auxNPC)

160               .flags.TargetNPC = 0
170           Else
180               Call WriteConsoleMsg(userindex, "Antes debes hacer click sobre el NPC.", FontTypeNames.FONTTYPE_INFO)
190           End If
200       End With
End Sub

''
' Handles the "WarnUser" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/26/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim Reason As String
              Dim Privs As PlayerType
              Dim Count As Byte

90            UserName = buffer.ReadASCIIString()
100           Reason = buffer.ReadASCIIString()

110           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And _
                  (Not .flags.Privilegios And PlayerType.User) <> 0 Or _
                   (.flags.Privilegios And PlayerType.ChaosCouncil Or _
                   .flags.Privilegios And PlayerType.RoyalCouncil) Then
                  
120               If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
130                   Call WriteConsoleMsg(userindex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
140               Else
150                   Privs = UserDarPrivilegioLevel(UserName)

160                   If Not Privs And PlayerType.User Then
170                       Call WriteConsoleMsg(userindex, "No puedes advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
180                   Else
190                       If (InStrB(UserName, "\") <> 0) Then
200                           UserName = Replace(UserName, "\", "")
210                       End If
220                       If (InStrB(UserName, "/") <> 0) Then
230                           UserName = Replace(UserName, "/", "")
240                       End If

250                       If FileExist(CharPath & UserName & ".chr", vbNormal) Then
260                           Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
270                           Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
280                           Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": ADVERTENCIA por: " & LCase$(Reason) & " " & Date & " " & time)

290                           Call WriteConsoleMsg(userindex, "Has advertido a " & UCase$(UserName) & ".", FontTypeNames.FONTTYPE_INFO)
300                           Call LogGM(.Name, " advirtio a " & UserName)
310                       End If
320                   End If
330               End If
340           End If

              'If we got here then packet is complete, copy data back to original queue
350           Call .incomingData.CopyBuffer(buffer)
360       End With

Errhandler:
          Dim error  As Long
370       error = Err.Number
380       On Error GoTo 0

          'Destroy auxiliar buffer
390       Set buffer = Nothing

400       If error <> 0 Then _
             Err.Raise error
End Sub


''
' Handles the "RequestCharInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInfo(ByVal userindex As Integer)
      '***************************************************
      'Author: Fredy Horacio Treboux (liquid)
      'Last Modification: 01/08/07
      'Last Modification by: (liquid).. alto bug zapallo..
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim TargetName As String
              Dim TargetIndex As Integer

90            TargetName = Replace$(buffer.ReadASCIIString(), "+", " ")
100           TargetIndex = NameIndex(TargetName)


110           If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
                  'is the player offline?
120               If TargetIndex <= 0 Then
                      'don't allow to retrieve administrator's info
130                   If Not (EsDios(TargetName) Or EsAdmin(TargetName)) Then
140                       Call WriteConsoleMsg(userindex, "Usuario offline, buscando en charfile.", FontTypeNames.FONTTYPE_INFO)
150                       Call SendUserStatsTxtOFF(userindex, TargetName)
160                   End If
170               Else
                      'don't allow to retrieve administrator's info
180                   If UserList(TargetIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
190                       Call SendUserStatsTxt(userindex, TargetIndex)
200                   End If
210               End If
220           End If

              'If we got here then packet is complete, copy data back to original queue
230           Call .incomingData.CopyBuffer(buffer)
240       End With

Errhandler:
          Dim error  As Long
250       error = Err.Number
260       On Error GoTo 0

          'Destroy auxiliar buffer
270       Set buffer = Nothing

280       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "RequestCharStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharStats(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer
90            UserName = buffer.ReadASCIIString()

100           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
110               Call LogGM(.Name, "/STAT " & UserName)

120               tUser = NameIndex(UserName)

130               If tUser <= 0 Then
140                   Call WriteConsoleMsg(userindex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_INFO)

150                   Call SendUserMiniStatsTxtFromChar(userindex, UserName)
160               Else
170                   Call SendUserMiniStatsTxt(userindex, tUser)
180               End If
190           End If

              'If we got here then packet is complete, copy data back to original queue
200           Call .incomingData.CopyBuffer(buffer)
210       End With

Errhandler:
          Dim error  As Long
220       error = Err.Number
230       On Error GoTo 0

          'Destroy auxiliar buffer
240       Set buffer = Nothing

250       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "RequestCharGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharGold(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer

90            UserName = buffer.ReadASCIIString()
100           tUser = NameIndex(UserName)

110           If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
120               Call LogGM(.Name, "/BAL " & UserName)

130               If tUser <= 0 Then
140                   Call WriteConsoleMsg(userindex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)

150                   Call SendUserOROTxtFromChar(userindex, UserName)
160               Else
170                   Call WriteConsoleMsg(userindex, "El usuario " & UserName & " tiene " & UserList(tUser).Stats.Banco & " en el banco.", FontTypeNames.FONTTYPE_TALK)
180               End If
190           End If

              'If we got here then packet is complete, copy data back to original queue
200           Call .incomingData.CopyBuffer(buffer)
210       End With

Errhandler:
          Dim error  As Long
220       error = Err.Number
230       On Error GoTo 0

          'Destroy auxiliar buffer
240       Set buffer = Nothing

250       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "RequestCharInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInventory(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer

90            UserName = buffer.ReadASCIIString()
100           tUser = NameIndex(UserName)


110           If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
120               Call LogGM(.Name, "/INV " & UserName)

130               If tUser <= 0 Then
140                   Call WriteConsoleMsg(userindex, "Usuario offline. Leyendo del charfile...", FontTypeNames.FONTTYPE_TALK)

150                   Call SendUserInvTxtFromChar(userindex, UserName)
160               Else
170                   Call SendUserInvTxt(userindex, tUser)
180               End If
190           End If

              'If we got here then packet is complete, copy data back to original queue
200           Call .incomingData.CopyBuffer(buffer)
210       End With

Errhandler:
          Dim error  As Long
220       error = Err.Number
230       On Error GoTo 0

          'Destroy auxiliar buffer
240       Set buffer = Nothing

250       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "RequestCharBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharBank(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer

90            UserName = buffer.ReadASCIIString()
100           tUser = NameIndex(UserName)


110           If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
120               Call LogGM(.Name, "/BOV " & UserName)

130               If tUser <= 0 Then
140                   Call WriteConsoleMsg(userindex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)

150                   Call SendUserBovedaTxtFromChar(userindex, UserName)
160               Else
170                   Call SendUserBovedaTxt(userindex, tUser)
180               End If
190           End If

              'If we got here then packet is complete, copy data back to original queue
200           Call .incomingData.CopyBuffer(buffer)
210       End With

Errhandler:
          Dim error  As Long
220       error = Err.Number
230       On Error GoTo 0

          'Destroy auxiliar buffer
240       Set buffer = Nothing

250       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "RequestCharSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharSkills(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer
              Dim LoopC As Long
              Dim message As String

90            UserName = buffer.ReadASCIIString()
100           tUser = NameIndex(UserName)


110           If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
120               Call LogGM(.Name, "/STATS " & UserName)

130               If tUser <= 0 Then
140                   If (InStrB(UserName, "\") <> 0) Then
150                       UserName = Replace(UserName, "\", "")
160                   End If
170                   If (InStrB(UserName, "/") <> 0) Then
180                       UserName = Replace(UserName, "/", "")
190                   End If

200                   For LoopC = 1 To NUMSKILLS
210                       message = message & "CHAR>" & SkillsNames(LoopC) & " = " & GetVar(CharPath & UserName & ".chr", "SKILLS", "SK" & LoopC) & vbCrLf
220                   Next LoopC

230                   Call WriteConsoleMsg(userindex, message & "CHAR> Libres:" & GetVar(CharPath & UserName & ".chr", "STATS", "SKILLPTSLIBRES"), FontTypeNames.FONTTYPE_INFO)
240               Else
250                   Call SendUserSkillsTxt(userindex, tUser)
260               End If
270           End If

              'If we got here then packet is complete, copy data back to original queue
280           Call .incomingData.CopyBuffer(buffer)
290       End With

Errhandler:
          Dim error  As Long
300       error = Err.Number
310       On Error GoTo 0

          'Destroy auxiliar buffer
320       Set buffer = Nothing

330       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "ReviveChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReviveChar(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 11/03/2010
      '11/03/2010: ZaMa - Al revivir con el comando, si esta navegando le da cuerpo e barca.
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer
              Dim LoopC As Byte

90            UserName = buffer.ReadASCIIString()


100           If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
110               If UCase$(UserName) <> "YO" Then
120                   tUser = NameIndex(UserName)
130               Else
140                   tUser = userindex
150               End If

160               If tUser <= 0 Then
170                   Call WriteConsoleMsg(userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
180               Else
190                   With UserList(tUser)
                          'If dead, show him alive (naked).
200                       If .flags.Muerto = 1 Then
210                           .flags.Muerto = 0

220                           If .flags.Navegando = 1 Then
230                               Call ToggleBoatBody(tUser)
240                           Else
250                               Call DarCuerpoDesnudo(tUser)
260                           End If

270                           If .flags.Traveling = 1 Then
280                               .flags.Traveling = 0
290                               .Counters.goHome = 0
300                               Call WriteMultiMessage(tUser, eMessages.CancelHome)
310                           End If

320                           Call ChangeUserChar(tUser, .Char.body, .OrigChar.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

330                           Call WriteConsoleMsg(tUser, UserList(userindex).Name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
340                       Else
350                           Call WriteConsoleMsg(tUser, UserList(userindex).Name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)
360                       End If

370                       .Stats.MinHp = .Stats.MaxHp

380                       If .flags.Traveling = 1 Then
390                           .Counters.goHome = 0
400                           .flags.Traveling = 0
410                           Call WriteMultiMessage(tUser, eMessages.CancelHome)
420                       End If

430                   End With

440                   Call WriteUpdateHP(tUser)
450                   Call WriteUpdateFollow(tUser)

460                   Call FlushBuffer(tUser)

470                   Call LogGM(.Name, "Resucito a " & UserName)
480               End If
490           End If

              'If we got here then packet is complete, copy data back to original queue
500           Call .incomingData.CopyBuffer(buffer)
510       End With

Errhandler:
          Dim error  As Long
520       error = Err.Number
530       On Error GoTo 0

          'Destroy auxiliar buffer
540       Set buffer = Nothing

550       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "OnlineGM" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineGM(ByVal userindex As Integer)
      '***************************************************
      'Author: Fredy Horacio Treboux (liquid)
      'Last Modification: 12/28/06
      '
      '***************************************************
          Dim i      As Long
          Dim list   As String
          Dim priv   As PlayerType

10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
              
40            priv = PlayerType.Consejero Or PlayerType.SemiDios
50            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv Or PlayerType.Dios Or PlayerType.Admin
              
60            For i = 1 To LastUser
70                If UserList(i).flags.UserLogged Then
80                    If UserList(i).flags.Privilegios And priv And _
                            Not Protocol.IsNickEspecial(UserList(i).Name) Then
90                         list = list & UserList(i).Name & ", "
                      
100                   End If
110               End If
120           Next i

130           If LenB(list) <> 0 Then
140               list = Left$(list, Len(list) - 2)
150               Call WriteConsoleMsg(userindex, list & ".", FontTypeNames.FONTTYPE_INFO)
160           Else
170               Call WriteConsoleMsg(userindex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)
180           End If
190       End With
End Sub

''
' Handles the "OnlineMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineMap(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 23/03/2009
      '23/03/2009: ZaMa - Ahora no requiere estar en el mapa, sino que por defecto se toma en el que esta, pero se puede especificar otro
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

              Dim map As Integer
30            map = .incomingData.ReadInteger

40            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

              Dim LoopC As Long
              Dim list As String
              Dim priv As PlayerType

50            priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
60            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv + (PlayerType.Dios Or PlayerType.Admin)

70            For LoopC = 1 To LastUser
80                If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).Pos.map = map Then
90                    If UserList(LoopC).flags.Privilegios And priv Then _
                         list = list & UserList(LoopC).Name & ", "
100               End If
110           Next LoopC

120           If Len(list) > 2 Then list = Left$(list, Len(list) - 2)

130           Call WriteConsoleMsg(userindex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)
140       End With
End Sub

''
' Handles the "Forgive" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForgive(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer

90            UserName = buffer.ReadASCIIString()

100           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
110               tUser = NameIndex(UserName)

120               If tUser > 0 Then
130                   If EsNewbie(tUser) Then
140                       Call VolverCiudadano(tUser)
150                   Else
160                       Call LogGM(.Name, "Intento perdonar un personaje de nivel avanzado.")
170                       Call WriteConsoleMsg(userindex, "Sólo se permite perdonar newbies.", FontTypeNames.FONTTYPE_INFO)
180                   End If
190               End If
200           End If

              'If we got here then packet is complete, copy data back to original queue
210           Call .incomingData.CopyBuffer(buffer)
220       End With

Errhandler:
          Dim error  As Long
230       error = Err.Number
240       On Error GoTo 0

          'Destroy auxiliar buffer
250       Set buffer = Nothing

260       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "Kick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKick(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer
              Dim Rank As Integer

90            Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

100           UserName = buffer.ReadASCIIString()

110           If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
120               tUser = NameIndex(UserName)

130               If tUser <= 0 Then
140                   Call WriteConsoleMsg(userindex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
150               Else
160                   If (UserList(tUser).flags.Privilegios And Rank) > (.flags.Privilegios And Rank) Then
170                       Call WriteConsoleMsg(userindex, "No puedes echar a alguien con jerarquía mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
180                   Else
190                       Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " echó a " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
200                       Call CloseSocket(tUser)
210                       Call LogGM(.Name, "Echó a " & UserName)
220                   End If
230               End If
240           End If

              'If we got here then packet is complete, copy data back to original queue
250           Call .incomingData.CopyBuffer(buffer)
260       End With

Errhandler:
          Dim error  As Long
270       error = Err.Number
280       On Error GoTo 0

          'Destroy auxiliar buffer
290       Set buffer = Nothing

300       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "Execute" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExecute(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer

90            UserName = buffer.ReadASCIIString()

100           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
110               tUser = NameIndex(UserName)

120               If tUser > 0 Then
130                   If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
140                       Call WriteConsoleMsg(userindex, "¿¿Estás loco?? ¿¿Cómo vas a piñatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)
150                   Else
160                       Call UserDie(tUser)
170                       Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " ha ejecutado a " & UserName & ".", FontTypeNames.FONTTYPE_EJECUCION))
180                       Call LogGM(.Name, " ejecuto a " & UserName)
190                   End If
200               Else
210                   Call WriteConsoleMsg(userindex, "No está online.", FontTypeNames.FONTTYPE_INFO)
220               End If
230           End If

              'If we got here then packet is complete, copy data back to original queue
240           Call .incomingData.CopyBuffer(buffer)
250       End With

Errhandler:
          Dim error  As Long
260       error = Err.Number
270       On Error GoTo 0

          'Destroy auxiliar buffer
280       Set buffer = Nothing

290       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "BanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanChar(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim Reason As String

90            UserName = buffer.ReadASCIIString()
100           Reason = buffer.ReadASCIIString()

110           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
120               Call BanCharacter(userindex, UserName, Reason)
130           End If

              'If we got here then packet is complete, copy data back to original queue
140           Call .incomingData.CopyBuffer(buffer)
150       End With

Errhandler:
          Dim error  As Long
160       error = Err.Number
170       On Error GoTo 0

          'Destroy auxiliar buffer
180       Set buffer = Nothing

190       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "UnbanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanChar(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50    On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)
              
              'Remove packet ID
80            Call buffer.ReadByte
              
              Dim UserName As String
              Dim cantPenas As Byte
              
90            UserName = buffer.ReadASCIIString()
              
100           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) <> 0 Then
110               If (InStrB(UserName, "\") <> 0) Then
120                   UserName = Replace(UserName, "\", "")
130               End If
140               If (InStrB(UserName, "/") <> 0) Then
150                   UserName = Replace(UserName, "/", "")
160               End If
                  
170               If Not FileExist(CharPath & UserName & ".chr", vbNormal) Then
180                   Call WriteConsoleMsg(userindex, "Charfile inexistente (no use +).", FontTypeNames.FONTTYPE_INFO)
190               Else
200                   If (val(GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban")) = 1) Then
210                       Call UnBan(UserName)
                      
                          'penas
220                       cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
230                       Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
240                       Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": UNBAN. " & Date & " " & time)
                      
250                       Call LogGM(.Name, "/UNBAN a " & UserName)
260                       Call WriteConsoleMsg(userindex, UserName & " unbanned.", FontTypeNames.FONTTYPE_INFO)
270                   Else
280                       Call WriteConsoleMsg(userindex, UserName & " no está baneado. Imposible unbanear.", FontTypeNames.FONTTYPE_INFO)
290                   End If
300               End If
310           End If
              
              'If we got here then packet is complete, copy data back to original queue
320           Call .incomingData.CopyBuffer(buffer)
330       End With

Errhandler:
          Dim error As Long
340       error = Err.Number
350   On Error GoTo 0
          
          'Destroy auxiliar buffer
360       Set buffer = Nothing
          
370       If error <> 0 Then _
              Err.Raise error
End Sub

''
' Handles the "NPCFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNPCFollow(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

40            If .flags.TargetNPC > 0 Then
50                Call DoFollow(.flags.TargetNPC, .Name)
60                Npclist(.flags.TargetNPC).flags.Inmovilizado = 0
70                Npclist(.flags.TargetNPC).flags.Paralizado = 0
80                Npclist(.flags.TargetNPC).Contadores.Paralisis = 0
90            End If
100       End With
End Sub

''
' Handles the "SummonChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSummonChar(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 26/03/2009
      '26/03/2009: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer
              Dim X  As Integer
              Dim Y  As Integer

90            UserName = buffer.ReadASCIIString()

100           If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) Then
110               tUser = NameIndex(UserName)

                  If Not Protocol.IsNickEspecial(UserName) Then
130                   If tUser <= 0 Then
140                       Call WriteConsoleMsg(userindex, "El jugador no está online.", FontTypeNames.FONTTYPE_INFO)
150                   Else
160                       If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or _
                             (UserList(tUser).flags.Privilegios And (PlayerType.Consejero Or PlayerType.User)) <> 0 Then
                              
170                           If UserList(tUser).flags.SlotEvent > 0 Or UserList(tUser).flags.SlotReto > 0 Then
180                               Call WriteConsoleMsg(userindex, "El personaje esta evento. Tene mayor cuidado para la proxima que me vas a buguear el evento " & .Name & ".", FontTypeNames.FONTTYPE_ADMIN)
190                           Else
200                               If Not UserList(tUser).Counters.Pena >= 1 Then
210                                   Call WriteConsoleMsg(tUser, .Name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
220                                   X = .Pos.X
230                                   Y = .Pos.Y + 1
240                                   Call FindLegalPos(tUser, .Pos.map, X, Y)
250                                   Call WarpUserChar(tUser, .Pos.map, X, Y, True, True)
260                                   Call LogGM(.Name, "/SUM " & UserName & " Map:" & .Pos.map & " X:" & .Pos.X & " Y:" & .Pos.Y)
270                               Else
280                                   Call WriteConsoleMsg(userindex, "Está en la carcel", FontTypeNames.FONTTYPE_INFO)
290                               End If
300                           End If
310                       End If
320                   End If
330               End If
340           End If

              'If we got here then packet is complete, copy data back to original queue
350           Call .incomingData.CopyBuffer(buffer)
360       End With

Errhandler:
          Dim error  As Long
370       error = Err.Number
380       On Error GoTo 0

          'Destroy auxiliar buffer
390       Set buffer = Nothing

400       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "SpawnListRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnListRequest(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

40            Call EnviarSpawnList(userindex)
50        End With
End Sub

''
' Handles the "SpawnCreature" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnCreature(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim Npc As Integer
70            Npc = .incomingData.ReadInteger()

80            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
90                If Npc > 0 And Npc <= UBound(Declaraciones.SpawnList()) Then _
                     Call SpawnNpc(Declaraciones.SpawnList(Npc).NpcIndex, .Pos, True, False)

100               Call LogGM(.Name, "Sumoneo " & Declaraciones.SpawnList(Npc).NpcName)
110           End If
120       End With
End Sub

''
' Handles the "ResetNPCInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResetNPCInventory(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
40            If .flags.TargetNPC = 0 Then Exit Sub

50            Call ResetNpcInv(.flags.TargetNPC)
60            Call LogGM(.Name, "/RESETINV " & Npclist(.flags.TargetNPC).Name)
70        End With
End Sub

''
' Handles the "CleanWorld" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCleanWorld(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub

40            CountDownLimpieza = 10
50            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Limpieza del mundo en 10 segundos. Recojan sus objetos para no perderlos.", FontTypeNames.FONTTYPE_SERVER)
              
              'Call LimpiarMundo
60        End With
End Sub

''
' Handles the "ServerMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerMessage(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 28/05/2010
      '28/05/2010: ZaMa - Ahora no dice el nombre del gm que lo dice.
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
70            Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
80            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
90            Call buffer.ReadByte

              Dim message As String
100           message = buffer.ReadASCIIString()

110           If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) Then
120               If LenB(message) <> 0 Then
130                   Call LogGM(.Name, "Mensaje Broadcast:" & message)
140                   Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(message, FontTypeNames.FONTTYPE_GUILD))
                      ''''''''''''''''SOLO PARA EL TESTEO'''''''
                      ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
                      'frmMain.txtChat.Text = frmMain.txtChat.Text & vbNewLine & UserList(UserIndex).name & " > " & message
150               End If
160           End If

              'If we got here then packet is complete, copy data back to original queue
170           Call .incomingData.CopyBuffer(buffer)
180       End With

Errhandler:
          Dim error  As Long
190       error = Err.Number
200       On Error GoTo 0

          'Destroy auxiliar buffer
210       Set buffer = Nothing

220       If error <> 0 Then _
             Err.Raise error
End Sub
Private Sub HandleRolMensaje(ByVal userindex As Integer)
      '***************************************************
      'Author: Tomás (Nibf ~) Para Gs zone y Servers Argentum
      'Last Modification: 20/09/13
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim message As String
90            message = buffer.ReadASCIIString()

100           If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
110               If LenB(message) <> 0 Then
120                   Call LogGM(.Name, "Mensaje Broadcast:" & message)
130                   Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("" & .Name & "> " & message, FontTypeNames.FONTTYPE_GUILD))
                      ''''''''''''''''SOLO PARA EL TESTEO'''''''
                      ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
140                   frmMain.txtChat.Text = frmMain.txtChat.Text & vbNewLine & message
150               End If
160           End If

              'If we got here then packet is complete, copy data back to original queue
170           Call .incomingData.CopyBuffer(buffer)
180       End With

Errhandler:
          Dim error  As Long
190       error = Err.Number
200       On Error GoTo 0

          'Destroy auxiliar buffer
210       Set buffer = Nothing

220       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "NickToIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNickToIP(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 07/06/2010
      'Pablo (ToxicWaste): Agrego para que el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
      '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
70            Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
80            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
90            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer
              Dim priv As PlayerType
              Dim IsAdmin As Boolean

100           UserName = buffer.ReadASCIIString()

110           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
120               tUser = NameIndex(UserName)
130               Call LogGM(.Name, "NICK2IP Solicito la IP de " & UserName)

140               IsAdmin = (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0
150               If IsAdmin Then
160                   priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
170               Else
180                   priv = PlayerType.User
190               End If

200               If tUser > 0 Then
210                   If UserList(tUser).flags.Privilegios And priv Then
220                       Call WriteConsoleMsg(userindex, "El ip de " & UserName & " es " & UserList(tUser).ip, FontTypeNames.FONTTYPE_INFO)
                          Dim ip As String
                          Dim lista As String
                          Dim LoopC As Long
230                       ip = UserList(tUser).ip
240                       For LoopC = 1 To LastUser
250                           If UserList(LoopC).ip = ip Then
260                               If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
270                                   If UserList(LoopC).flags.Privilegios And priv Then
280                                       lista = lista & UserList(LoopC).Name & ", "
290                                   End If
300                               End If
310                           End If
320                       Next LoopC
330                       If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
340                       Call WriteConsoleMsg(userindex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
350                   End If
360               Else
370                   If Not (EsDios(UserName) Or EsAdmin(UserName)) Or IsAdmin Then
380                       Call WriteConsoleMsg(userindex, "No hay ningún personaje con ese nick.", FontTypeNames.FONTTYPE_INFO)
390                   End If
400               End If
410           End If

              'If we got here then packet is complete, copy data back to original queue
420           Call .incomingData.CopyBuffer(buffer)
430       End With

Errhandler:
          Dim error  As Long
440       error = Err.Number
450       On Error GoTo 0

          'Destroy auxiliar buffer
460       Set buffer = Nothing

470       If error <> 0 Then _
             Err.Raise error
End Sub
''
' Handles the "IPToNick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleIPToNick(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim ip As String
              Dim LoopC As Long
              Dim lista As String
              Dim priv As PlayerType

70            ip = .incomingData.ReadByte() & "."
80            ip = ip & .incomingData.ReadByte() & "."
90            ip = ip & .incomingData.ReadByte() & "."
100           ip = ip & .incomingData.ReadByte()

110           If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub

120           Call LogGM(.Name, "IP2NICK Solicito los Nicks de IP " & ip)

130           If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
140               priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
150           Else
160               priv = PlayerType.User
170           End If

180           For LoopC = 1 To LastUser
190               If UserList(LoopC).ip = ip Then
200                   If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
210                       If UserList(LoopC).flags.Privilegios And priv Then
220                           lista = lista & UserList(LoopC).Name & ", "
230                       End If
240                   End If
250               End If
260           Next LoopC

270           If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
280           Call WriteConsoleMsg(userindex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
290       End With
End Sub

''
' Handles the "GuildOnlineMembers" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnlineMembers(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim GuildName As String
              Dim tGuild As Integer

90            GuildName = buffer.ReadASCIIString()

100           If (InStrB(GuildName, "+") <> 0) Then
110               GuildName = Replace(GuildName, "+", " ")
120           End If

130           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
140               tGuild = GuildIndex(GuildName)

150               If tGuild > 0 Then
160                   Call WriteConsoleMsg(userindex, "Clan " & UCase(GuildName) & ": " & _
                                                      modGuilds.m_ListaDeMiembrosOnline(userindex, tGuild), FontTypeNames.FONTTYPE_GUILDMSG)
170               End If
180           End If

              'If we got here then packet is complete, copy data back to original queue
190           Call .incomingData.CopyBuffer(buffer)
200       End With

Errhandler:
          Dim error  As Long
210       error = Err.Number
220       On Error GoTo 0

          'Destroy auxiliar buffer
230       Set buffer = Nothing

240       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "TeleportCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportCreate(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 22/03/2010
      '15/11/2009: ZaMa - Ahora se crea un teleport con un radio especificado.
      '22/03/2010: ZaMa - Harcodeo los teleps y radios en el dat, para evitar mapas bugueados.
      '***************************************************
10        If UserList(userindex).incomingData.length < 6 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim Mapa As Integer
              Dim X  As Byte
              Dim Y  As Byte
              Dim Radio As Byte

70            Mapa = .incomingData.ReadInteger()
80            X = .incomingData.ReadByte()
90            Y = .incomingData.ReadByte()
100           Radio = .incomingData.ReadByte()

110           Radio = MinimoInt(Radio, 6)

120           If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

130           Call LogGM(.Name, "/CT " & Mapa & "," & X & "," & Y & "," & Radio)

140           If Not MapaValido(Mapa) Or Not InMapBounds(Mapa, X, Y) Then _
                 Exit Sub

150           If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then _
                 Exit Sub

160           If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).TileExit.map > 0 Then _
                 Exit Sub

170           If MapData(Mapa, X, Y).ObjInfo.ObjIndex > 0 Then
180               Call WriteConsoleMsg(userindex, "Hay un objeto en el piso en ese lugar.", FontTypeNames.FONTTYPE_INFO)
190               Exit Sub
200           End If

210           If MapData(Mapa, X, Y).TileExit.map > 0 Then
220               Call WriteConsoleMsg(userindex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
230               Exit Sub
240           End If

              Dim ET As Obj
250           ET.Amount = 1
              ' Es el numero en el dat. El indice es el comienzo + el radio, todo harcodeado :(.
260           ET.ObjIndex = 378

270           With MapData(.Pos.map, .Pos.X, .Pos.Y - 1)
280               .TileExit.map = Mapa
290               .TileExit.X = X
300               .TileExit.Y = Y
310           End With

320           Call MakeObj(ET, .Pos.map, .Pos.X, .Pos.Y - 1)
330       End With
End Sub

''
' Handles the "TeleportDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportDestroy(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        With UserList(userindex)
              Dim Mapa As Integer
              Dim X  As Byte
              Dim Y  As Byte

              'Remove packet ID
20            Call .incomingData.ReadByte

              '/dt
30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

40            Mapa = .flags.TargetMap
50            X = .flags.TargetX
60            Y = .flags.TargetY

70            If Not InMapBounds(Mapa, X, Y) Then Exit Sub

80            With MapData(Mapa, X, Y)
90                If .ObjInfo.ObjIndex = 0 Then Exit Sub

100               If ObjData(.ObjInfo.ObjIndex).ObjType = eOBJType.otTeleport And .TileExit.map > 0 Then
110                   Call LogGM(UserList(userindex).Name, "/DT: " & Mapa & "," & X & "," & Y)

120                   Call EraseObj(.ObjInfo.Amount, Mapa, X, Y)

130                   If MapData(.TileExit.map, .TileExit.X, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
140                       Call EraseObj(1, .TileExit.map, .TileExit.X, .TileExit.Y)
150                   End If

160                   .TileExit.map = 0
170                   .TileExit.X = 0
180                   .TileExit.Y = 0
190               End If
200           End With
210       End With
End Sub

''
' Handles the "RainToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRainToggle(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

40            Call LogGM(.Name, "/LLUVIA")
50            Lloviendo = Not Lloviendo

              'Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
60        End With
End Sub

''
' Handles the "SetCharDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetCharDescription(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim tUser As Integer
              Dim desc As String

90            desc = buffer.ReadASCIIString()

100           If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (.flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
110               tUser = .flags.TargetUser
120               If tUser > 0 Then
130                   UserList(tUser).DescRM = desc
140               Else
150                   Call WriteConsoleMsg(userindex, "Haz click sobre un personaje antes.", FontTypeNames.FONTTYPE_INFO)
160               End If
170           End If

              'If we got here then packet is complete, copy data back to original queue
180           Call .incomingData.CopyBuffer(buffer)
190       End With

Errhandler:
          Dim error  As Long
200       error = Err.Number
210       On Error GoTo 0

          'Destroy auxiliar buffer
220       Set buffer = Nothing

230       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "ForceMIDIToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HanldeForceMIDIToMap(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 4 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim midiID As Byte
              Dim Mapa As Integer

70            midiID = .incomingData.ReadByte
80            Mapa = .incomingData.ReadInteger

              'Solo dioses, admins y RMS
90            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
                  'Si el mapa no fue enviado tomo el actual
100               If Not InMapBounds(Mapa, 50, 50) Then
110                   Mapa = .Pos.map
120               End If

130               If midiID = 0 Then
                      'Ponemos el default del mapa
140                   Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(MapInfo(.Pos.map).Music))
150               Else
                      'Ponemos el pedido por el GM
160                   Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(midiID))
170               End If
180           End If
190       End With
End Sub

''
' Handles the "ForceWAVEToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEToMap(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 6 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim waveID As Byte
              Dim Mapa As Integer
              Dim X  As Byte
              Dim Y  As Byte

70            waveID = .incomingData.ReadByte()
80            Mapa = .incomingData.ReadInteger()
90            X = .incomingData.ReadByte()
100           Y = .incomingData.ReadByte()

              'Solo dioses, admins y RMS
110           If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
                  'Si el mapa no fue enviado tomo el actual
120               If Not InMapBounds(Mapa, X, Y) Then
130                   Mapa = .Pos.map
140                   X = .Pos.X
150                   Y = .Pos.Y
160               End If

                  'Ponemos el pedido por el GM
170               Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayWave(waveID, X, Y))
180           End If
190       End With
End Sub

''
' Handles the "RoyalArmyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyMessage(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim message As String
90            message = buffer.ReadASCIIString()

              'Solo dioses, admins y RMS
100           If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster Or PlayerType.RoyalCouncil) Then
110               Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Consejo de Banderbill] " & .Name & "> " & message, FontTypeNames.FONTTYPE_CONSEJOVesA))
120           End If

              'If we got here then packet is complete, copy data back to original queue
130           Call .incomingData.CopyBuffer(buffer)
140       End With

Errhandler:
          Dim error  As Long
150       error = Err.Number
160       On Error GoTo 0

          'Destroy auxiliar buffer
170       Set buffer = Nothing

180       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "ChaosLegionMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionMessage(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim message As String
90            message = buffer.ReadASCIIString()

              'Solo dioses, admins y RMS
100           If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster Or PlayerType.ChaosCouncil) Then
110               Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Concilio de las Sombras] " & .Name & "> " & message, FontTypeNames.FONTTYPE_EJECUCION))
120           End If

              'If we got here then packet is complete, copy data back to original queue
130           Call .incomingData.CopyBuffer(buffer)
140       End With

Errhandler:
          Dim error  As Long
150       error = Err.Number
160       On Error GoTo 0

          'Destroy auxiliar buffer
170       Set buffer = Nothing

180       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "CitizenMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCitizenMessage(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim message As String
90            message = buffer.ReadASCIIString()

              'Solo dioses, admins y RMS
100           If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
110               Call SendData(SendTarget.ToCiudadanosYRMs, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & message, FontTypeNames.FONTTYPE_TALK))
120           End If

              'If we got here then packet is complete, copy data back to original queue
130           Call .incomingData.CopyBuffer(buffer)
140       End With

Errhandler:
          Dim error  As Long
150       error = Err.Number
160       On Error GoTo 0

          'Destroy auxiliar buffer
170       Set buffer = Nothing

180       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "CriminalMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCriminalMessage(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim message As String
90            message = buffer.ReadASCIIString()

              'Solo dioses, admins y RMS
100           If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
110               Call SendData(SendTarget.ToCriminalesYRMs, 0, PrepareMessageConsoleMsg("CRIMINALES> " & message, FontTypeNames.FONTTYPE_TALK))
120           End If

              'If we got here then packet is complete, copy data back to original queue
130           Call .incomingData.CopyBuffer(buffer)
140       End With

Errhandler:
          Dim error  As Long
150       error = Err.Number
160       On Error GoTo 0

          'Destroy auxiliar buffer
170       Set buffer = Nothing

180       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "TalkAsNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalkAsNPC(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/29/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim message As String
90            message = buffer.ReadASCIIString()

              'Solo dioses, admins y RMS
100           If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
                  'Asegurarse haya un NPC seleccionado
110               If .flags.TargetNPC > 0 Then
120                   Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(message, Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
130               Else
140                   Call WriteConsoleMsg(userindex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando.", FontTypeNames.FONTTYPE_INFO)
150               End If
160           End If

              'If we got here then packet is complete, copy data back to original queue
170           Call .incomingData.CopyBuffer(buffer)
180       End With

Errhandler:
          Dim error  As Long
190       error = Err.Number
200       On Error GoTo 0

          'Destroy auxiliar buffer
210       Set buffer = Nothing

220       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "DestroyAllItemsInArea" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyAllItemsInArea(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

              Dim X  As Long
              Dim Y  As Long
              Dim bIsExit As Boolean

40            For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
50                For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
60                    If X > 0 And Y > 0 And X < 101 And Y < 101 Then
70                        If MapData(.Pos.map, X, Y).ObjInfo.ObjIndex > 0 Then
80                            bIsExit = MapData(.Pos.map, X, Y).TileExit.map > 0
90                            If ItemNoEsDeMapa(MapData(.Pos.map, X, Y).ObjInfo.ObjIndex, bIsExit) Then
100                               Call EraseObj(MAX_INVENTORY_OBJS, .Pos.map, X, Y)
110                           End If
120                       End If
130                   End If
140               Next X
150           Next Y

160           Call LogGM(UserList(userindex).Name, "/MASSDEST")
170       End With
End Sub

''
' Handles the "AcceptRoyalCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptRoyalCouncilMember(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer
              Dim LoopC As Byte

90            UserName = buffer.ReadASCIIString()

100           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
110               tUser = NameIndex(UserName)
120               If tUser <= 0 Then
130                   Call WriteConsoleMsg(userindex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
140               Else
150                   Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))
160                   With UserList(tUser)
170                       If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
180                       If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil

190                       Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False)
200                   End With
210               End If
220           End If

              'If we got here then packet is complete, copy data back to original queue
230           Call .incomingData.CopyBuffer(buffer)
240       End With

Errhandler:
          Dim error  As Long
250       error = Err.Number
260       On Error GoTo 0

          'Destroy auxiliar buffer
270       Set buffer = Nothing

280       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "ChaosCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptChaosCouncilMember(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer
              Dim LoopC As Byte

90            UserName = buffer.ReadASCIIString()

100           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
110               tUser = NameIndex(UserName)
120               If tUser <= 0 Then
130                   Call WriteConsoleMsg(userindex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
140               Else
150                   Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))

160                   With UserList(tUser)
170                       If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
180                       If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil

190                       Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False)
200                   End With
210               End If
220           End If

              'If we got here then packet is complete, copy data back to original queue
230           Call .incomingData.CopyBuffer(buffer)
240       End With

Errhandler:
          Dim error  As Long
250       error = Err.Number
260       On Error GoTo 0

          'Destroy auxiliar buffer
270       Set buffer = Nothing

280       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "ItemsInTheFloor" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleItemsInTheFloor(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

              Dim tObj As Integer
              Dim lista As String
              Dim X  As Long
              Dim Y  As Long

40            For X = 5 To 95
50                For Y = 5 To 95
60                    tObj = MapData(.Pos.map, X, Y).ObjInfo.ObjIndex
70                    If tObj > 0 Then
80                        If ObjData(tObj).ObjType <> eOBJType.otarboles Then
90                            Call WriteConsoleMsg(userindex, "(" & X & "," & Y & ") " & ObjData(tObj).Name, FontTypeNames.FONTTYPE_INFO)
100                       End If
110                   End If
120               Next Y
130           Next X
140       End With
End Sub

''
' Handles the "MakeDumb" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumb(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer

90            UserName = buffer.ReadASCIIString()

100           If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
110               tUser = NameIndex(UserName)
                  'para deteccion de aoice
120               If tUser <= 0 Then
130                   Call WriteConsoleMsg(userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
140               Else
150                   Call WriteDumb(tUser)
160               End If
170           End If

              'If we got here then packet is complete, copy data back to original queue
180           Call .incomingData.CopyBuffer(buffer)
190       End With

Errhandler:
          Dim error  As Long
200       error = Err.Number
210       On Error GoTo 0

          'Destroy auxiliar buffer
220       Set buffer = Nothing

230       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "MakeDumbNoMore" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumbNoMore(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer

90            UserName = buffer.ReadASCIIString()

100           If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
110               tUser = NameIndex(UserName)
                  'para deteccion de aoice
120               If tUser <= 0 Then
130                   Call WriteConsoleMsg(userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
140               Else
150                   Call WriteDumbNoMore(tUser)
160                   Call FlushBuffer(tUser)
170               End If
180           End If

              'If we got here then packet is complete, copy data back to original queue
190           Call .incomingData.CopyBuffer(buffer)
200       End With

Errhandler:
          Dim error  As Long
210       error = Err.Number
220       On Error GoTo 0

          'Destroy auxiliar buffer
230       Set buffer = Nothing

240       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "DumpIPTables" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDumpIPTables(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

40            Call SecurityIp.DumpTables
50        End With
End Sub

''
' Handles the "CouncilKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilKick(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer

90            UserName = buffer.ReadASCIIString()

100           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
110               tUser = NameIndex(UserName)
120               If tUser <= 0 Then
130                   If FileExist(CharPath & UserName & ".chr") Then
140                       Call WriteConsoleMsg(userindex, "Usuario offline, echando de los consejos.", FontTypeNames.FONTTYPE_INFO)
150                       Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECE", 0)
160                       Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECECAOS", 0)
170                   Else
180                       Call WriteConsoleMsg(userindex, "No se encuentra el charfile " & CharPath & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
190                   End If
200               Else
210                   With UserList(tUser)
220                       If .flags.Privilegios And PlayerType.RoyalCouncil Then
230                           Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Banderbill.", FontTypeNames.FONTTYPE_GUILD)
240                           .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil

250                           Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False)
260                           Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill.", FontTypeNames.FONTTYPE_INFO))
270                       End If

280                       If .flags.Privilegios And PlayerType.ChaosCouncil Then
290                           Call WriteConsoleMsg(tUser, "Has sido echado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_GUILD)
300                           .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil

310                           Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False)
320                           Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_INFO))
330                       End If
340                   End With
350               End If
360           End If

              'If we got here then packet is complete, copy data back to original queue
370           Call .incomingData.CopyBuffer(buffer)
380       End With

Errhandler:
          Dim error  As Long
390       error = Err.Number
400       On Error GoTo 0

          'Destroy auxiliar buffer
410       Set buffer = Nothing

420       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "SetTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetTrigger(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim tTrigger As Byte
              Dim tLog As String

70            tTrigger = .incomingData.ReadByte()

80            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

90            If tTrigger >= 0 Then
100               MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = tTrigger
110               tLog = "Trigger " & tTrigger & " en mapa " & .Pos.map & " " & .Pos.X & "," & .Pos.Y

120               Call LogGM(.Name, tLog)
130               Call WriteConsoleMsg(userindex, tLog, FontTypeNames.FONTTYPE_INFO)
140           End If
150       End With
End Sub

''
' Handles the "AskTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAskTrigger(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 04/13/07
      '
      '***************************************************
          Dim tTrigger As Byte

10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

40            tTrigger = MapData(.Pos.map, .Pos.X, .Pos.Y).trigger

50            Call LogGM(.Name, "Miro el trigger en " & .Pos.map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)

60            Call WriteConsoleMsg(userindex, _
                                   "Trigger " & tTrigger & " en mapa " & .Pos.map & " " & .Pos.X & ", " & .Pos.Y _
                                   , FontTypeNames.FONTTYPE_INFO)
70        End With
End Sub

''
' Handles the "BannedIPList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

              Dim lista As String
              Dim LoopC As Long

40            Call LogGM(.Name, "/BANIPLIST")

50            For LoopC = 1 To BanIps.Count
60                lista = lista & BanIps.Item(LoopC) & ", "
70            Next LoopC

80            If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)

90            Call WriteConsoleMsg(userindex, lista, FontTypeNames.FONTTYPE_INFO)
100       End With
End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

40            Call BanIpGuardar
50            Call BanIpCargar
60        End With
End Sub

''
' Handles the "GuildBan" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildBan(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim GuildName As String
              Dim cantMembers As Integer
              Dim LoopC As Long
              Dim member As String
              Dim Count As Byte
              Dim tIndex As Integer
              Dim tFile As String

90            GuildName = buffer.ReadASCIIString()

100           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
110               tFile = App.Path & "\guilds\" & GuildName & "-members.mem"

120               If Not FileExist(tFile) Then
130                   Call WriteConsoleMsg(userindex, "No existe el clan: " & GuildName, FontTypeNames.FONTTYPE_INFO)
140               Else
150                   Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " baneó al clan " & UCase$(GuildName), FontTypeNames.FONTTYPE_FIGHT))

                      'baneamos a los miembros
160                   Call LogGM(.Name, "BANCLAN a " & UCase$(GuildName))

170                   cantMembers = val(GetVar(tFile, "INIT", "NroMembers"))

180                   For LoopC = 1 To cantMembers
190                       member = GetVar(tFile, "Members", "Member" & LoopC)
                          'member es la victima
200                       Call Ban(member, "Administracion del servidor", "Clan Banned")

210                       Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("   " & member & "<" & GuildName & "> ha sido expulsado del servidor.", FontTypeNames.FONTTYPE_FIGHT))

220                       tIndex = NameIndex(member)
230                       If tIndex > 0 Then
                              'esta online
240                           UserList(tIndex).flags.Ban = 1
250                           Call CloseSocket(tIndex)
260                       End If

                          'ponemos el flag de ban a 1
270                       Call WriteVar(CharPath & member & ".chr", "FLAGS", "Ban", "1")
                          'ponemos la pena
280                       Count = val(GetVar(CharPath & member & ".chr", "PENAS", "Cant"))
290                       Call WriteVar(CharPath & member & ".chr", "PENAS", "Cant", Count + 1)
300                       Call WriteVar(CharPath & member & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": BAN AL CLAN: " & GuildName & " " & Date & " " & time)
310                   Next LoopC
320               End If
330           End If

              'If we got here then packet is complete, copy data back to original queue
340           Call .incomingData.CopyBuffer(buffer)
350       End With

Errhandler:
          Dim error  As Long
360       error = Err.Number
370       On Error GoTo 0

          'Destroy auxiliar buffer
380       Set buffer = Nothing

390       If error <> 0 Then _
             Err.Raise error
End Sub
''
' Handles the "CheckHD" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleCheckHD(ByVal userindex As Integer)
      '***************************************************
      'Author: ArzenaTh
      'Last Modification: 01/09/10
      'Verifica el HD del usuario.
      '***************************************************

10        If UserList(userindex).incomingData.length < 6 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler


60        With UserList(userindex)
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)
80            Call buffer.ReadByte

              Dim Usuario As Integer
              Dim nickUsuario As String
90            nickUsuario = buffer.ReadASCIIString()
100           Usuario = NameIndex(nickUsuario)

110           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
120               If Usuario = 0 Then
130                   Call WriteConsoleMsg(userindex, "Usuario offline.", FONTTYPE_INFO)
140               Else
150                   Call WriteConsoleMsg(userindex, "El disco del usuario " & UserList(Usuario).Name & " es " & UserList(Usuario).HD, FONTTYPE_INFOBOLD)
160               End If
170           End If

180           Call .incomingData.CopyBuffer(buffer)

190       End With

Errhandler:
          Dim error  As Long
200       error = Err.Number
210       On Error GoTo 0

220       Set buffer = Nothing

230       If error <> 0 Then Err.Raise error

End Sub
''
' Handles the "BanHD" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleBanHD(ByVal userindex As Integer)
      '***************************************************
      'Author: ArzenaTh
      'Last Modification: 02/09/10
      'Maneja el baneo del serial del HD de un usuario.
      '***************************************************

10        If UserList(userindex).incomingData.length < 6 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)
80            Call buffer.ReadByte

              Dim Usuario As Integer
90            Usuario = NameIndex(buffer.ReadASCIIString())
              Dim bannedHD As String
100           If Usuario > 0 Then bannedHD = UserList(Usuario).HD
              Dim i  As Long    'El mandamás dijo Long, Long será.
110           If .flags.Privilegios And (PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Dios) Then
120               If LenB(bannedHD) > 0 Then
130                   If BuscarRegistroHD(bannedHD) > 0 Then
140                       Call WriteConsoleMsg(userindex, "El usuario ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
150                   Else
160                       Call AgregarRegistroHD(bannedHD)
170                       Call WriteConsoleMsg(userindex, "Has baneado el root " & bannedHD & " del usuario " & UserList(Usuario).Name, FontTypeNames.FONTTYPE_INFO)
                          'Call CloseSocket(Usuario)
180                       For i = 1 To LastUser
190                           If UserList(i).ConnIDValida Then
200                               If UserList(i).HD = bannedHD Then
210                                   Call BanCharacter(userindex, UserList(i).Name, "t0 en el servidor")
220                               End If
230                           End If
240                       Next i
250                   End If
260               ElseIf Usuario <= 0 Then
270                   Call WriteConsoleMsg(userindex, "El usuario no se encuentra online.", FontTypeNames.FONTTYPE_INFO)
280               End If
290           End If

300           Call .incomingData.CopyBuffer(buffer)
310       End With

Errhandler:
          Dim error  As Long
320       error = Err.Number
330       On Error GoTo 0

340       Set buffer = Nothing

350       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the "UnBanHD" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUnbanHD(ByVal userindex As Integer)
      '***************************************************
      'Author: ArzenaTh
      'Last Modification: 02/09/10
      'Maneja el unbaneo del serial del HD de un usuario.
      '***************************************************

10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50    On Error GoTo Errhandler
60        With UserList(userindex)
          
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
70            Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
80            Call buffer.CopyBuffer(.incomingData)
              
              'Remove packet ID
90            Call buffer.ReadByte
              
              Dim HD As String
100           HD = buffer.ReadASCIIString()
              
110           If .flags.Privilegios And (PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Dios) Then
          
120               If (RemoverRegistroHD(HD)) Then
130                   Call WriteConsoleMsg(userindex, "El root n°" & HD & " se ha quitado de la lista de baneados.", FONTTYPE_INFOBOLD)
140               Else
150                   Call WriteConsoleMsg(userindex, "El root n°" & HD & " no se encuentra en la lista de baneados.", FONTTYPE_INFO)
160               End If
170           End If
              
180           Call .incomingData.CopyBuffer(buffer)
190       End With
Errhandler:
          Dim error  As Long
200       error = Err.Number
210       On Error GoTo 0

220       Set buffer = Nothing

230       If error <> 0 Then Err.Raise error

End Sub


''
' Handles the "BanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 07/02/09
      'Agregado un CopyBuffer porque se producia un bucle
      'inifito al intentar banear una ip ya baneada. (NicoNZ)
      '07/02/09 Pato - Ahora no es posible saber si un gm está o no online.
      '***************************************************
10        If UserList(userindex).incomingData.length < 6 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim bannedIP As String
              Dim tUser As Integer
              Dim Reason As String
              Dim i  As Long

              ' Is it by ip??
90            If buffer.ReadBoolean() Then
100               bannedIP = buffer.ReadByte() & "."
110               bannedIP = bannedIP & buffer.ReadByte() & "."
120               bannedIP = bannedIP & buffer.ReadByte() & "."
130               bannedIP = bannedIP & buffer.ReadByte()
140           Else
150               tUser = NameIndex(buffer.ReadASCIIString())

160               If tUser > 0 Then bannedIP = UserList(tUser).ip
170           End If

180           Reason = buffer.ReadASCIIString()


190           If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
200               If LenB(bannedIP) > 0 Then
210                   Call LogGM(.Name, "/BanIP " & bannedIP & " por " & Reason)

220                   If BanIpBuscar(bannedIP) > 0 Then
230                       Call WriteConsoleMsg(userindex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
240                   Else
250                       Call BanIpAgrega(bannedIP)
260                       Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " baneó la IP " & bannedIP & " por " & Reason, FontTypeNames.FONTTYPE_FIGHT))

                          'Find every player with that ip and ban him!
270                       For i = 1 To LastUser
280                           If UserList(i).ConnIDValida Then
290                               If UserList(i).ip = bannedIP Then
300                                   Call BanCharacter(userindex, UserList(i).Name, "IP POR " & Reason)
310                               End If
320                           End If
330                       Next i
340                   End If
350               ElseIf tUser <= 0 Then
360                   Call WriteConsoleMsg(userindex, "El personaje no está online.", FontTypeNames.FONTTYPE_INFO)
370               End If
380           End If

              'If we got here then packet is complete, copy data back to original queue
390           Call .incomingData.CopyBuffer(buffer)
400       End With

Errhandler:
          Dim error  As Long
410       error = Err.Number
420       On Error GoTo 0

          'Destroy auxiliar buffer
430       Set buffer = Nothing

440       If error <> 0 Then _
             Err.Raise error
End Sub

Private Sub HandleUnbanIP(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte
              
              Dim bannedIP As String
              
70            bannedIP = .incomingData.ReadByte() & "."
80            bannedIP = bannedIP & .incomingData.ReadByte() & "."
90            bannedIP = bannedIP & .incomingData.ReadByte() & "."
100           bannedIP = bannedIP & .incomingData.ReadByte()
              
110           If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
              
120           If BanIpQuita(bannedIP) Then
130               Call WriteConsoleMsg(userindex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
140           Else
150               Call WriteConsoleMsg(userindex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
160           End If
170       End With
End Sub

''
' Handles the "CreateItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateItem(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 11/02/2011
      'maTih.- : Ahora se puede elegir, la cantidad a crear.
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim tObj As Integer
              Dim Cuantos As Integer
              Dim tStr As String
70            tObj = .incomingData.ReadInteger()
80            Cuantos = .incomingData.ReadInteger()

90            If Not Protocol.IsNickEspecial(UCase$(.Name)) Then Exit Sub

100           Call LogGM(.Name, "/CI: " & tObj & " Cantidad : " & Cuantos)

110           If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then _
                 Exit Sub

120           If Cuantos > 9999 Then Call WriteConsoleMsg(userindex, "Demasiados, máximo para crear : 10.000", FontTypeNames.FONTTYPE_TALK): Exit Sub

130           If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).TileExit.map > 0 Then _
                 Exit Sub

140           If tObj < 1 Or tObj > NumObjDatas Then _
                 Exit Sub

              'Is the object not null?
150           If LenB(ObjData(tObj).Name) = 0 Then Exit Sub

              Dim Objeto As Obj
160           Call WriteConsoleMsg(userindex, "¡¡ATENCIÓN: FUERON CREADOS ***" & Cuantos & "*** ÍTEMS, TIRE Y /DEST LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)

170           Objeto.Amount = Cuantos
180           Objeto.ObjIndex = tObj
190           Call MakeObj(Objeto, .Pos.map, .Pos.X, .Pos.Y - 1)

              'Agrega a la lista.
              Dim tmpPos As WorldPos

200           tmpPos = .Pos
210           tmpPos.Y = .Pos.X - 1


220       End With
End Sub

''
' Handles the "DestroyItems" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyItems(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

40            If MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex = 0 Then Exit Sub

50            Call LogGM(.Name, "/DEST")

60            If ObjData(MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).ObjType = eOBJType.otTeleport And _
                 MapData(.Pos.map, .Pos.X, .Pos.Y).TileExit.map > 0 Then

70                Call WriteConsoleMsg(userindex, "No puede destruir teleports así. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
80                Exit Sub
90            End If

100           Call EraseObj(10000, .Pos.map, .Pos.X, .Pos.Y)
110       End With
End Sub

''
' Handles the "ChaosLegionKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionKick(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer

90            UserName = buffer.ReadASCIIString()

100           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And _
                  (.flags.Privilegios And PlayerType.ChaosCouncil) Or _
                  (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
110               If (InStrB(UserName, "\") <> 0) Then
120                   UserName = Replace(UserName, "\", "")
130               End If
140               If (InStrB(UserName, "/") <> 0) Then
150                   UserName = Replace(UserName, "/", "")
160               End If
170               tUser = NameIndex(UserName)

180               Call LogGM(.Name, "ECHO DEL CAOS A: " & UserName)

190               If tUser > 0 Then
200                   Call ExpulsarFaccionCaos(tUser, True)
210                   UserList(tUser).Faccion.Reenlistadas = 200
220                   Call WriteConsoleMsg(userindex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
230                   Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas del caos.", FontTypeNames.FONTTYPE_FIGHT)
240                   Call FlushBuffer(tUser)
250               Else
260                   If FileExist(CharPath & UserName & ".chr") Then
270                       Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoCaos", 0)
280                       Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas", 200)
290                       Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & .Name)
300                       Call WriteConsoleMsg(userindex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
310                   Else
320                       Call WriteConsoleMsg(userindex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)
330                   End If
340               End If
350           End If

              'If we got here then packet is complete, copy data back to original queue
360           Call .incomingData.CopyBuffer(buffer)
370       End With

Errhandler:
          Dim error  As Long
380       error = Err.Number
390       On Error GoTo 0

          'Destroy auxiliar buffer
400       Set buffer = Nothing

410       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "RoyalArmyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyKick(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer
              Dim Privs As PlayerType
              
90            UserName = buffer.ReadASCIIString()

100           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And _
                   .flags.Privilegios And PlayerType.RoyalCouncil Or _
                   (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
110               If (InStrB(UserName, "\") <> 0) Then
120                   UserName = Replace(UserName, "\", "")
130               End If
140               If (InStrB(UserName, "/") <> 0) Then
150                   UserName = Replace(UserName, "/", "")
160               End If
170               tUser = NameIndex(UserName)

180               Call LogGM(.Name, "ECHÓ DE LA REAL A: " & UserName)

190               If tUser > 0 Then
200                   Call ExpulsarFaccionReal(tUser, True)
210                   UserList(tUser).Faccion.Reenlistadas = 200
220                   Call WriteConsoleMsg(userindex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
230                   Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas reales.", FontTypeNames.FONTTYPE_FIGHT)
240                   Call FlushBuffer(tUser)
250               Else
260                   If FileExist(CharPath & UserName & ".chr") Then
270                       Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoReal", 0)
280                       Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas", 200)
290                       Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & .Name)
300                       Call WriteConsoleMsg(userindex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
310                   Else
320                       Call WriteConsoleMsg(userindex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)
330                   End If
340               End If
350           End If

              'If we got here then packet is complete, copy data back to original queue
360           Call .incomingData.CopyBuffer(buffer)
370       End With

Errhandler:
          Dim error  As Long
380       error = Err.Number
390       On Error GoTo 0

          'Destroy auxiliar buffer
400       Set buffer = Nothing

410       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "ForceMIDIAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceMIDIAll(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim midiID As Byte
70            midiID = .incomingData.ReadByte()

80            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

90            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " broadcast música: " & midiID, FontTypeNames.FONTTYPE_SERVER))

100           Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMidi(midiID))
110       End With
End Sub

''
' Handles the "ForceWAVEAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEAll(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim waveID As Byte
70            waveID = .incomingData.ReadByte()

80            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

90            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))
100       End With
End Sub

''
' Handles the "RemovePunishment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRemovePunishment(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 1/05/07
      'Pablo (ToxicWaste): 1/05/07, You can now edit the punishment.
      '***************************************************
10        If UserList(userindex).incomingData.length < 6 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim punishment As Byte
              Dim NewText As String

90            UserName = buffer.ReadASCIIString()
100           punishment = buffer.ReadByte
110           NewText = buffer.ReadASCIIString()

120           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
130               If LenB(UserName) = 0 Then
140                   Call WriteConsoleMsg(userindex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", FontTypeNames.FONTTYPE_INFO)
150               Else
160                   If (InStrB(UserName, "\") <> 0) Then
170                       UserName = Replace(UserName, "\", "")
180                   End If
190                   If (InStrB(UserName, "/") <> 0) Then
200                       UserName = Replace(UserName, "/", "")
210                   End If

220                   If FileExist(CharPath & UserName & ".chr", vbNormal) Then
230                       Call LogGM(.Name, " borro la pena: " & punishment & "-" & _
                                            GetVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment) _
                                            & " de " & UserName & " y la cambió por: " & NewText)

240                       Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment, LCase$(.Name) & ": <" & NewText & "> " & Date & " " & time)

250                       Call WriteConsoleMsg(userindex, "Pena modificada.", FontTypeNames.FONTTYPE_INFO)
260                   End If
270               End If
280           End If

              'If we got here then packet is complete, copy data back to original queue
290           Call .incomingData.CopyBuffer(buffer)
300       End With

Errhandler:
          Dim error  As Long
310       error = Err.Number
320       On Error GoTo 0

          'Destroy auxiliar buffer
330       Set buffer = Nothing

340       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "TileBlockedToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTileBlockedToggle(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

40            Call LogGM(.Name, "/BLOQ")

50            If MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked = 0 Then
60                MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked = 1
70            Else
80                MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked = 0
90            End If

100           Call Bloquear(True, .Pos.map, .Pos.X, .Pos.Y, MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked)
110       End With
End Sub

''
' Handles the "KillNPCNoRespawn" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPCNoRespawn(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

40            If .flags.TargetNPC = 0 Then Exit Sub

50            Call QuitarNPC(.flags.TargetNPC)
60            Call LogGM(.Name, "/MATA " & Npclist(.flags.TargetNPC).Name)
70        End With
End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

              Dim X  As Long
              Dim Y  As Long

40            For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
50                For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
60                    If X > 0 And Y > 0 And X < 101 And Y < 101 Then
70                        If MapData(.Pos.map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(.Pos.map, X, Y).NpcIndex)
80                    End If
90                Next X
100           Next Y
110           Call LogGM(.Name, "/MASSKILL")
120       End With
End Sub

''
' Handles the "LastIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLastIP(ByVal userindex As Integer)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 12/30/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim lista As String
              Dim LoopC As Byte
              Dim priv As Integer
              Dim validCheck As Boolean

90            priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
100           UserName = buffer.ReadASCIIString()

110           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
                  'Handle special chars
120               If (InStrB(UserName, "\") <> 0) Then
130                   UserName = Replace(UserName, "\", "")
140               End If
150               If (InStrB(UserName, "\") <> 0) Then
160                   UserName = Replace(UserName, "/", "")
170               End If
180               If (InStrB(UserName, "+") <> 0) Then
190                   UserName = Replace(UserName, "+", " ")
200               End If

                  'Only Gods and Admins can see the ips of adminsitrative characters. All others can be seen by every adminsitrative char.
210               If NameIndex(UserName) > 0 Then
220                   validCheck = (UserList(NameIndex(UserName)).flags.Privilegios And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
230               Else
240                   validCheck = (UserDarPrivilegioLevel(UserName) And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
250               End If

260               If validCheck Then
270                   Call LogGM(.Name, "/LASTIP " & UserName)

280                   If FileExist(CharPath & UserName & ".chr", vbNormal) Then
290                       lista = "Las ultimas IPs con las que " & UserName & " se conectó son:"
300                       For LoopC = 1 To 5
310                           lista = lista & vbCrLf & LoopC & " - " & GetVar(CharPath & UserName & ".chr", "INIT", "LastIP" & LoopC)
320                       Next LoopC
330                       Call WriteConsoleMsg(userindex, lista, FontTypeNames.FONTTYPE_INFO)
340                   Else
350                       Call WriteConsoleMsg(userindex, "Charfile """ & UserName & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
360                   End If
370               Else
380                   Call WriteConsoleMsg(userindex, UserName & " es de mayor jerarquía que vos.", FontTypeNames.FONTTYPE_INFO)
390               End If
400           End If

              'If we got here then packet is complete, copy data back to original queue
410           Call .incomingData.CopyBuffer(buffer)
420       End With

Errhandler:
          Dim error  As Long
430       error = Err.Number
440       On Error GoTo 0

          'Destroy auxiliar buffer
450       Set buffer = Nothing

460       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "ChatColor" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleChatColor(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/23/06
      'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
      'Change the user`s chat color
      '***************************************************
10        If UserList(userindex).incomingData.length < 4 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove packet ID
60            Call .incomingData.ReadByte

              Dim color As Long

70            color = RGB(.incomingData.ReadByte(), .incomingData.ReadByte(), .incomingData.ReadByte())

80            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
90                .flags.ChatColor = color
100           End If
110       End With
End Sub

''
' Handles the "Ignored" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIgnored(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/23/06
      'Ignore the user
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
40                .flags.AdminPerseguible = Not .flags.AdminPerseguible
50            End If
60        End With
End Sub
Public Sub HandleUserOro(ByVal userindex As Integer)

10        With UserList(userindex)
20            Call .incomingData.ReadByte


              'Lo hace vip
30            If .flags.Oro = 0 Then
40                .flags.Oro = 1
50            End If

              'Le da el Random de vida entre 5 y 13 cambiar a su gusto
60            If .Stats.MaxHp + RandomNumber(1, 3) Then
70            End If
80        End With
End Sub
Public Sub HandleUserPlata(ByVal userindex As Integer)

10        With UserList(userindex)
20            Call .incomingData.ReadByte


              'Lo hace vip
30            If .flags.Plata = 0 Then
40                .flags.Plata = 1
50            End If

              'Le da el Random de vida entre 5 y 13 cambiar a su gusto
60            If .Stats.MaxHp + RandomNumber(1, 2) Then
70            End If
80        End With
End Sub
Public Sub HandleUserBronce(ByVal userindex As Integer)

10        With UserList(userindex)
20            Call .incomingData.ReadByte


              'Lo hace vip
30            If .flags.Bronce = 0 Then
40                .flags.Bronce = 1
50            End If

              'Le da el Random de vida entre 5 y 13 cambiar a su gusto
60            If .Stats.MaxHp + RandomNumber(1, 1) Then
70            End If
80        End With
End Sub

''
' Handles the "CheckSlot" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleCheckSlot(ByVal userindex As Integer)
      '***************************************************
      'Author: Pablo (ToxicWaste)
      'Last Modification: 09/09/2008 (NicoNZ)
      'Check one Users Slot in Particular from Inventory
      '***************************************************
10        If UserList(userindex).incomingData.length < 4 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              'Reads the UserName and Slot Packets
              Dim UserName As String
              Dim Slot As Byte
              Dim tIndex As Integer

90            UserName = buffer.ReadASCIIString()    'Que UserName?
100           Slot = buffer.ReadByte()    'Que Slot?

110           If .flags.Privilegios And (PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Dios) Then
120               tIndex = NameIndex(UserName)  'Que user index?

130               Call LogGM(.Name, .Name & " Checkeó el slot " & Slot & " de " & UserName)

140               If tIndex > 0 Then
150                   If Slot > 0 And Slot <= UserList(tIndex).CurrentInventorySlots Then
160                       If UserList(tIndex).Invent.Object(Slot).ObjIndex > 0 Then
170                           Call WriteConsoleMsg(userindex, " Objeto " & Slot & ") " & ObjData(UserList(tIndex).Invent.Object(Slot).ObjIndex).Name & " Cantidad:" & UserList(tIndex).Invent.Object(Slot).Amount, FontTypeNames.FONTTYPE_INFO)
180                       Else
190                           Call WriteConsoleMsg(userindex, "No hay ningún objeto en slot seleccionado.", FontTypeNames.FONTTYPE_INFO)
200                       End If
210                   Else
220                       Call WriteConsoleMsg(userindex, "Slot Inválido.", FontTypeNames.FONTTYPE_TALK)
230                   End If
240               Else
250                   Call WriteConsoleMsg(userindex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
260               End If
270           End If

              'If we got here then packet is complete, copy data back to original queue
280           Call .incomingData.CopyBuffer(buffer)
290       End With

Errhandler:
          Dim error  As Long
300       error = Err.Number
310       On Error GoTo 0

          'Destroy auxiliar buffer
320       Set buffer = Nothing

330       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handles the "ResetAutoUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleResetAutoUpdate(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/23/06
      'Reset the AutoUpdate
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
              If Not Protocol.IsNickEspecial(.Name) Then Exit Sub

50            Call WriteConsoleMsg(userindex, "TID: " & CStr(ReiniciarAutoUpdate()), FontTypeNames.FONTTYPE_INFO)
60        End With
End Sub

''
' Handles the "Restart" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleRestart(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/23/06
      'Restart the game
      '***************************************************
10        With UserList(userindex)
              'Remove Packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
40            If Not Protocol.IsNickEspecial(.Name) Then Exit Sub
              
              'time and Time BUG!
50            Call LogGM(.Name, .Name & " reinició el mundo.")

60            Call ReiniciarServidor(True)
70        End With
End Sub

''
' Handles the "ReloadObjects" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadObjects(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/23/06
      'Reload the objects
      '***************************************************
10        With UserList(userindex)
              'Remove Packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

40            Call LogGM(.Name, .Name & " ha recargado los objetos.")

50            Call LoadOBJData
60        End With
End Sub

''
' Handles the "ReloadSpells" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadSpells(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/23/06
      'Reload the spells
      '***************************************************
10        With UserList(userindex)
              'Remove Packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

40            Call LogGM(.Name, .Name & " ha recargado los hechizos.")

50            Call CargarHechizos
60        End With
End Sub

''
' Handle the "ReloadServerIni" message.
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadServerIni(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/23/06
      'Reload the Server`s INI
      '***************************************************
10        With UserList(userindex)
              'Remove Packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

40            Call LogGM(.Name, .Name & " ha recargado los INITs.")

50            Call LoadSini
60        End With
End Sub

''
' Handle the "ReloadNPCs" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadNPCs(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/23/06
      'Reload the Server`s NPC
      '***************************************************
10        With UserList(userindex)
              'Remove Packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

40            Call LogGM(.Name, .Name & " ha recargado los NPCs.")

50            Call CargaNpcsDat

60            Call WriteConsoleMsg(userindex, "Npcs.dat recargado.", FontTypeNames.FONTTYPE_INFO)
70        End With
End Sub

''
' Handle the "KickAllChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleKickAllChars(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/23/06
      'Kick all the chars that are online
      '***************************************************
10        With UserList(userindex)
              'Remove Packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

40            Call LogGM(.Name, .Name & " ha echado a todos los personajes.")

50            Call EcharPjsNoPrivilegiados
60        End With
End Sub

''
' Handle the "Night" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNight(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/23/06
      'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
      '
      '***************************************************
10        With UserList(userindex)
              'Remove Packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
40            If Not Protocol.IsNickEspecial(.Name) Then Exit Sub

50            DeNoche = Not DeNoche

              Dim i  As Long

60            For i = 1 To NumUsers
70                If UserList(i).flags.UserLogged And UserList(i).ConnID > -1 Then
80                    Call EnviarNoche(i)
90                End If
100           Next i
110       End With
End Sub

''
' Handle the "ShowServerForm" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowServerForm(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/23/06
      'Show the server form
      '***************************************************
10        With UserList(userindex)
              'Remove Packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

40            Call LogGM(.Name, .Name & " ha solicitado mostrar el formulario del servidor.")
50            Call frmMain.mnuMostrar_Click
60        End With
End Sub

''
' Handle the "CleanSOS" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCleanSOS(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/23/06
      'Clean the SOS
      '***************************************************
10        With UserList(userindex)
              'Remove Packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

40            Call LogGM(.Name, .Name & " ha borrado los SOS.")

50            Call Ayuda.Reset
60        End With
End Sub

''
' Handle the "SaveChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/23/06
      'Save the characters
      '***************************************************
10        With UserList(userindex)
              'Remove Packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

40            Call LogGM(.Name, .Name & " ha guardado todos los chars.")

50            Call mGroup.DistributeExpAndGldGroups
60            Call GuardarUsuarios
70        End With
End Sub

''
' Handle the "ChangeMapInfoBackup" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/24/06
      'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
      'Change the backup`s info of the map
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove Packet ID
60            Call .incomingData.ReadByte

              Dim doTheBackUp As Boolean

70            doTheBackUp = .incomingData.ReadBoolean()

80            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub

90            Call LogGM(.Name, .Name & " ha cambiado la información sobre el BackUp.")

              'Change the boolean to byte in a fast way
100           If doTheBackUp Then
110               MapInfo(.Pos.map).BackUp = 1
120           Else
130               MapInfo(.Pos.map).BackUp = 0
140           End If

              'Change the boolean to string in a fast way
150           Call WriteVar(App.Path & MapPath & "mapa" & .Pos.map & ".dat", "Mapa" & .Pos.map, "backup", MapInfo(.Pos.map).BackUp)

160           Call WriteConsoleMsg(userindex, "Mapa " & .Pos.map & " Backup: " & MapInfo(.Pos.map).BackUp, FontTypeNames.FONTTYPE_INFO)
170       End With
End Sub

''
' Handle the "ChangeMapInfoPK" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoPK(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/24/06
      'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
      'Change the pk`s info of the  map
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove Packet ID
60            Call .incomingData.ReadByte

              Dim isMapPk As Boolean

70            isMapPk = .incomingData.ReadBoolean()

80            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub

90            Call LogGM(.Name, .Name & " ha cambiado la información sobre si es PK el mapa.")

100           MapInfo(.Pos.map).Pk = isMapPk

              'Change the boolean to string in a fast way
110           Call WriteVar(App.Path & MapPath & "mapa" & .Pos.map & ".dat", "Mapa" & .Pos.map, "Pk", IIf(isMapPk, "1", "0"))

120           Call WriteConsoleMsg(userindex, "Mapa " & .Pos.map & " PK: " & MapInfo(.Pos.map).Pk, FontTypeNames.FONTTYPE_INFO)
130       End With
End Sub

''
' Handle the "ChangeMapInfoRestricted" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal userindex As Integer)
      '***************************************************
      'Author: Pablo (ToxicWaste)
      'Last Modification: 26/01/2007
      'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS", "FACCION".
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
          Dim tStr   As String

60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove Packet ID
80            Call buffer.ReadByte

90            tStr = buffer.ReadASCIIString()

100           If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
110               If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Or tStr = "QUINCE" Or tStr = "VEINTE" Or tStr = "VEINTICINCO" Or tStr = "CUARENTA" Or tStr = "SEIS" Or tStr = "SIETE" Or tStr = "OCHO" Or tStr = "NUEVE" Or tStr = "CINCO" Or tStr = "MENOSCINCO" Or tStr = "MENOSCUATRO" Or tStr = "NOESUM" Or tStr = "VIPP" Or tStr = "VIP" Then
120                   Call LogGM(.Name, .Name & " ha cambiado la información sobre si es restringido el mapa.")
130                   MapInfo(UserList(userindex).Pos.map).Restringir = tStr
140                   Call WriteVar(App.Path & MapPath & "mapa" & UserList(userindex).Pos.map & ".dat", "Mapa" & UserList(userindex).Pos.map, "Restringir", tStr)
150                   Call WriteConsoleMsg(userindex, "Mapa " & .Pos.map & " Restringido: " & MapInfo(.Pos.map).Restringir, FontTypeNames.FONTTYPE_INFO)
160               Else
170                   Call WriteConsoleMsg(userindex, "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION', 'QUINCE',  'VENTE', 'VEINTICINCO', 'CUARENTA', 'CINCO', 'SEIS', 'SIETE', 'OCHO', 'NUEVE', 'MENOSCINCO', 'MENOSCUATRO', 'NOESUM', 'VIPP', 'VIPP'.", FontTypeNames.FONTTYPE_INFO)
180               End If
190           End If

              'If we got here then packet is complete, copy data back to original queue
200           Call .incomingData.CopyBuffer(buffer)
210       End With

Errhandler:
          Dim error  As Long
220       error = Err.Number
230       On Error GoTo 0

          'Destroy auxiliar buffer
240       Set buffer = Nothing

250       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handle the "ChangeMapInfoNoMagic" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoMagic(ByVal userindex As Integer)
      '***************************************************
      'Author: Pablo (ToxicWaste)
      'Last Modification: 26/01/2007
      'MagiaSinEfecto -> Options: "1" , "0".
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          Dim nomagic As Boolean

50        With UserList(userindex)
              'Remove Packet ID
60            Call .incomingData.ReadByte

70            nomagic = .incomingData.ReadBoolean

80            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
90                Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido usar la magia el mapa.")
100               MapInfo(UserList(userindex).Pos.map).MagiaSinEfecto = nomagic
110               Call WriteVar(App.Path & MapPath & "mapa" & UserList(userindex).Pos.map & ".dat", "Mapa" & UserList(userindex).Pos.map, "MagiaSinEfecto", nomagic)
120               Call WriteConsoleMsg(userindex, "Mapa " & .Pos.map & " MagiaSinEfecto: " & MapInfo(.Pos.map).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)
130           End If
140       End With
End Sub

''
' Handle the "ChangeMapInfoNoInvi" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvi(ByVal userindex As Integer)
      '***************************************************
      'Author: Pablo (ToxicWaste)
      'Last Modification: 26/01/2007
      'InviSinEfecto -> Options: "1", "0"
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          Dim noinvi As Boolean

50        With UserList(userindex)
              'Remove Packet ID
60            Call .incomingData.ReadByte

70            noinvi = .incomingData.ReadBoolean()

80            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
90                Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido usar la invisibilidad en el mapa.")
100               MapInfo(UserList(userindex).Pos.map).InviSinEfecto = noinvi
110               Call WriteVar(App.Path & MapPath & "mapa" & UserList(userindex).Pos.map & ".dat", "Mapa" & UserList(userindex).Pos.map, "InviSinEfecto", noinvi)
120               Call WriteConsoleMsg(userindex, "Mapa " & .Pos.map & " InviSinEfecto: " & MapInfo(.Pos.map).InviSinEfecto, FontTypeNames.FONTTYPE_INFO)
130           End If
140       End With
End Sub

''
' Handle the "ChangeMapInfoNoResu" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoResu(ByVal userindex As Integer)
      '***************************************************
      'Author: Pablo (ToxicWaste)
      'Last Modification: 26/01/2007
      'ResuSinEfecto -> Options: "1", "0"
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          Dim noresu As Boolean

50        With UserList(userindex)
              'Remove Packet ID
60            Call .incomingData.ReadByte

70            noresu = .incomingData.ReadBoolean()

80            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
90                Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido usar el resucitar en el mapa.")
100               MapInfo(UserList(userindex).Pos.map).ResuSinEfecto = noresu
110               Call WriteVar(App.Path & MapPath & "mapa" & UserList(userindex).Pos.map & ".dat", "Mapa" & UserList(userindex).Pos.map, "ResuSinEfecto", noresu)
120               Call WriteConsoleMsg(userindex, "Mapa " & .Pos.map & " ResuSinEfecto: " & MapInfo(.Pos.map).ResuSinEfecto, FontTypeNames.FONTTYPE_INFO)
130           End If
140       End With
End Sub

''
' Handle the "ChangeMapInfoLand" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoLand(ByVal userindex As Integer)
      '***************************************************
      'Author: Pablo (ToxicWaste)
      'Last Modification: 26/01/2007
      'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
          Dim tStr   As String

60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove Packet ID
80            Call buffer.ReadByte

90            tStr = buffer.ReadASCIIString()

100           If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
110               If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
120                   Call LogGM(.Name, .Name & " ha cambiado la información del terreno del mapa.")
130                   MapInfo(UserList(userindex).Pos.map).Terreno = tStr
140                   Call WriteVar(App.Path & MapPath & "mapa" & UserList(userindex).Pos.map & ".dat", "Mapa" & UserList(userindex).Pos.map, "Terreno", tStr)
150                   Call WriteConsoleMsg(userindex, "Mapa " & .Pos.map & " Terreno: " & MapInfo(.Pos.map).Terreno, FontTypeNames.FONTTYPE_INFO)
160               Else
170                   Call WriteConsoleMsg(userindex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
180                   Call WriteConsoleMsg(userindex, "Igualmente, el único útil es 'NIEVE' ya que al ingresarlo, la gente muere de frío en el mapa.", FontTypeNames.FONTTYPE_INFO)
190               End If
200           End If

              'If we got here then packet is complete, copy data back to original queue
210           Call .incomingData.CopyBuffer(buffer)
220       End With

Errhandler:
          Dim error  As Long
230       error = Err.Number
240       On Error GoTo 0

          'Destroy auxiliar buffer
250       Set buffer = Nothing

260       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handle the "ChangeMapInfoZone" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoZone(ByVal userindex As Integer)
      '***************************************************
      'Author: Pablo (ToxicWaste)
      'Last Modification: 26/01/2007
      'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
          Dim tStr   As String

60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove Packet ID
80            Call buffer.ReadByte

90            tStr = buffer.ReadASCIIString()

100           If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
110               If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
120                   Call LogGM(.Name, .Name & " ha cambiado la información de la zona del mapa.")
130                   MapInfo(UserList(userindex).Pos.map).Zona = tStr
140                   Call WriteVar(App.Path & MapPath & "mapa" & UserList(userindex).Pos.map & ".dat", "Mapa" & UserList(userindex).Pos.map, "Zona", tStr)
150                   Call WriteConsoleMsg(userindex, "Mapa " & .Pos.map & " Zona: " & MapInfo(.Pos.map).Zona, FontTypeNames.FONTTYPE_INFO)
160               Else
170                   Call WriteConsoleMsg(userindex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
180                   Call WriteConsoleMsg(userindex, "Igualmente, el único útil es 'DUNGEON' ya que al ingresarlo, NO se sentirá el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)
190               End If
200           End If

              'If we got here then packet is complete, copy data back to original queue
210           Call .incomingData.CopyBuffer(buffer)
220       End With

Errhandler:
          Dim error  As Long
230       error = Err.Number
240       On Error GoTo 0

          'Destroy auxiliar buffer
250       Set buffer = Nothing

260       If error <> 0 Then _
             Err.Raise error
End Sub
''
' Handle the "ChangeMapInfoStealNp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoStealNpc(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 25/07/2010
      'RoboNpcsPermitido -> Options: "1", "0"
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          Dim RoboNpc As Byte

50        With UserList(userindex)
              'Remove Packet ID
60            Call .incomingData.ReadByte

70            RoboNpc = val(IIf(.incomingData.ReadBoolean(), 1, 0))

80            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
90                Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido robar npcs en el mapa.")

100               MapInfo(UserList(userindex).Pos.map).RoboNpcsPermitido = RoboNpc

110               Call WriteVar(App.Path & MapPath & "mapa" & UserList(userindex).Pos.map & ".dat", "Mapa" & UserList(userindex).Pos.map, "RoboNpcsPermitido", RoboNpc)
120               Call WriteConsoleMsg(userindex, "Mapa " & .Pos.map & " RoboNpcsPermitido: " & MapInfo(.Pos.map).RoboNpcsPermitido, FontTypeNames.FONTTYPE_INFO)
130           End If
140       End With
End Sub

''
' Handle the "ChangeMapInfoNoOcultar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoOcultar(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 18/09/2010
      'OcultarSinEfecto -> Options: "1", "0"
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          Dim NoOcultar As Byte
          Dim Mapa   As Integer

50        With UserList(userindex)

              'Remove Packet ID
60            Call .incomingData.ReadByte

70            NoOcultar = val(IIf(.incomingData.ReadBoolean(), 1, 0))

80            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then

90                Mapa = .Pos.map

100               Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido ocultarse en el mapa " & Mapa & ".")

110               MapInfo(Mapa).OcultarSinEfecto = NoOcultar

120               Call WriteVar(App.Path & MapPath & "mapa" & Mapa & ".dat", "Mapa" & Mapa, "OcultarSinEfecto", NoOcultar)
130               Call WriteConsoleMsg(userindex, "Mapa " & Mapa & " OcultarSinEfecto: " & NoOcultar, FontTypeNames.FONTTYPE_INFO)
140           End If

150       End With

End Sub

''
' Handle the "ChangeMapInfoNoInvocar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvocar(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 18/09/2010
      'InvocarSinEfecto -> Options: "1", "0"
      '***************************************************
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          Dim NoInvocar As Byte
          Dim Mapa   As Integer

50        With UserList(userindex)

              'Remove Packet ID
60            Call .incomingData.ReadByte

70            NoInvocar = val(IIf(.incomingData.ReadBoolean(), 1, 0))

80            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then

90                Mapa = .Pos.map

100               Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido invocar en el mapa " & Mapa & ".")

110               MapInfo(Mapa).InvocarSinEfecto = NoInvocar

120               Call WriteVar(App.Path & MapPath & "mapa" & Mapa & ".dat", "Mapa" & Mapa, "InvocarSinEfecto", NoInvocar)
130               Call WriteConsoleMsg(userindex, "Mapa " & Mapa & " InvocarSinEfecto: " & NoInvocar, FontTypeNames.FONTTYPE_INFO)
140           End If

150       End With

End Sub

''
' Handle the "SaveMap" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/24/06
      'Saves the map
      '***************************************************
10        With UserList(userindex)
              'Remove Packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

40            Call LogGM(.Name, .Name & " ha guardado el mapa " & CStr(.Pos.map))

50            Call GrabarMapa(.Pos.map, App.Path & "\WorldBackUp\Mapa" & .Pos.map)

60            Call WriteConsoleMsg(userindex, "Mapa Guardado.", FontTypeNames.FONTTYPE_INFO)
70        End With
End Sub

''
' Handle the "ShowGuildMessages" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowGuildMessages(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/24/06
      'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
      'Allows admins to read guild messages
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim guild As String

90            guild = buffer.ReadASCIIString()

100           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
110               Call modGuilds.GMEscuchaClan(userindex, guild)
120           End If

              'If we got here then packet is complete, copy data back to original queue
130           Call .incomingData.CopyBuffer(buffer)
140       End With

Errhandler:
          Dim error  As Long
150       error = Err.Number
160       On Error GoTo 0

          'Destroy auxiliar buffer
170       Set buffer = Nothing

180       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handle the "DoBackUp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/24/06
      'Show guilds messages
      '***************************************************
10        With UserList(userindex)
              'Remove Packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

40            Call LogGM(.Name, .Name & " ha hecho un backup.")

50            Call ES.DoBackUp    'Sino lo confunde con la id del paquete
60        End With
End Sub

''
' Handle the "ToggleCentinelActivated" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleToggleCentinelActivated(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/26/06
      'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
      'Activate or desactivate the Centinel
      '***************************************************
10        With UserList(userindex)
              'Remove Packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

40            centinelaActivado = Not centinelaActivado

50            With Centinela
60                .RevisandoUserIndex = 0
70                .Clave = 0
80                .TiempoRestante = 0
90            End With

100           If CentinelaNPCIndex Then
110               Call QuitarNPC(CentinelaNPCIndex)
120               CentinelaNPCIndex = 0
130           End If

140           If centinelaActivado Then
150               Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido activado.", FontTypeNames.FONTTYPE_SERVER))
160           Else
170               Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido desactivado.", FontTypeNames.FONTTYPE_SERVER))
180           End If
190       End With
End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterName(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 12/26/06
      'Change user name
      '***************************************************
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              'Reads the userName and newUser Packets
              Dim UserName As String
              Dim newName As String
              Dim changeNameUI As Integer
              Dim GuildIndex As Integer

90            UserName = buffer.ReadASCIIString()
100           newName = buffer.ReadASCIIString()

110           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin)) Then
120               If LenB(UserName) = 0 Or LenB(newName) = 0 Then
130                   Call WriteConsoleMsg(userindex, "Usar: /ANAME origen@destino", FontTypeNames.FONTTYPE_INFO)
140               Else
150                   changeNameUI = NameIndex(UserName)

160                   If changeNameUI > 0 Then
170                       Call WriteConsoleMsg(userindex, "El Pj está online, debe salir para hacer el cambio.", FontTypeNames.FONTTYPE_WARNING)
180                   Else
190                       If Not FileExist(CharPath & UserName & ".chr") Then
200                           Call WriteConsoleMsg(userindex, "El pj " & UserName & " es inexistente.", FontTypeNames.FONTTYPE_INFO)
210                       Else
220                           GuildIndex = val(GetVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX"))

230                           If GuildIndex > 0 Then
240                               Call WriteConsoleMsg(userindex, "El pj " & UserName & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", FontTypeNames.FONTTYPE_INFO)
250                           Else
260                               If Not FileExist(CharPath & newName & ".chr") Then
270                                   Call FileCopy(CharPath & UserName & ".chr", CharPath & UCase$(newName) & ".chr")

280                                   Call WriteConsoleMsg(userindex, "Transferencia exitosa.", FontTypeNames.FONTTYPE_INFO)

290                                   Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")

                                      Dim cantPenas As Byte

300                                   cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))

310                                   Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", CStr(cantPenas + 1))

320                                   Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & CStr(cantPenas + 1), LCase$(.Name) & ": BAN POR Cambio de nick a " & UCase$(newName) & " " & Date & " " & time)

330                                   Call LogGM(.Name, "Ha cambiado de nombre al usuario " & UserName & ". Ahora se llama " & newName)
340                               Else
350                                   Call WriteConsoleMsg(userindex, "El nick solicitado ya existe.", FontTypeNames.FONTTYPE_INFO)
360                               End If
370                           End If
380                       End If
390                   End If
400               End If
410           End If

              'If we got here then packet is complete, copy data back to original queue
420           Call .incomingData.CopyBuffer(buffer)
430       End With

Errhandler:
          Dim error  As Long
440       error = Err.Number
450       On Error GoTo 0

          'Destroy auxiliar buffer
460       Set buffer = Nothing

470       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterMail(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 12/26/06
      'Change user password
      '***************************************************
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim newMail As String

90            UserName = buffer.ReadASCIIString()
100           newMail = buffer.ReadASCIIString()

110           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin)) Then
120               If LenB(UserName) = 0 Or LenB(newMail) = 0 Then
130                   Call WriteConsoleMsg(userindex, "usar /AEMAIL <pj>-<nuevomail>", FontTypeNames.FONTTYPE_INFO)
140               Else
150                   If Not FileExist(CharPath & UserName & ".chr") Then
160                       Call WriteConsoleMsg(userindex, "No existe el charfile " & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
170                   Else
180                       Call WriteVar(CharPath & UserName & ".chr", "CONTACTO", "Email", newMail)
190                       Call WriteConsoleMsg(userindex, "Email de " & UserName & " cambiado a: " & newMail, FontTypeNames.FONTTYPE_INFO)
200                   End If

210                   Call LogGM(.Name, "Le ha cambiado el mail a " & UserName)
220               End If
230           End If

              'If we got here then packet is complete, copy data back to original queue
240           Call .incomingData.CopyBuffer(buffer)
250       End With

Errhandler:
          Dim error  As Long
260       error = Err.Number
270       On Error GoTo 0

          'Destroy auxiliar buffer
280       Set buffer = Nothing

290       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handle the "AlterPassword" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterPassword(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 12/26/06
      'Change user password
      '***************************************************
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50    On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)
              
              'Remove packet ID
80            Call buffer.ReadByte
              
              Dim UserName As String
              Dim copyFrom As String
              Dim Password As String
              
90            UserName = Replace(buffer.ReadASCIIString(), "+", " ")
100           copyFrom = Replace(buffer.ReadASCIIString(), "+", " ")
              
110           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin)) Then
120               Call LogGM(.Name, "Ha alterado la contraseña de " & UserName)
                  
130               If LenB(UserName) = 0 Or LenB(copyFrom) = 0 Then
140                   Call WriteConsoleMsg(userindex, "usar /APASS <pjsinpass>@<pjconpass>", FontTypeNames.FONTTYPE_INFO)
150               Else
160                   If Not FileExist(CharPath & UserName & ".chr") Or Not FileExist(CharPath & copyFrom & ".chr") Then
170                       Call WriteConsoleMsg(userindex, "Alguno de los PJs no existe " & UserName & "@" & copyFrom, FontTypeNames.FONTTYPE_INFO)
180                   Else
190                       Password = GetVar(CharPath & copyFrom & ".chr", "INIT", "Password")
200                       Call WriteVar(CharPath & UserName & ".chr", "INIT", "Password", Password)
                          
210                       Call WriteConsoleMsg(userindex, "Password de " & UserName & " ha cambiado por la de " & copyFrom, FontTypeNames.FONTTYPE_INFO)
220                   End If
230               End If
240           End If
              
              'If we got here then packet is complete, copy data back to original queue
250           Call .incomingData.CopyBuffer(buffer)
260       End With

Errhandler:
          Dim error As Long
270       error = Err.Number
280   On Error GoTo 0
          
          'Destroy auxiliar buffer
290       Set buffer = Nothing
          
300       If error <> 0 Then _
              Err.Raise error
End Sub

''
' Handle the "HandleCreateNPC" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 12/24/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove Packet ID
60            Call .incomingData.ReadByte

              Dim NpcIndex As Integer

70            NpcIndex = .incomingData.ReadInteger()

80            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

90            NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)

100           If NpcIndex <> 0 Then
110               Call LogGM(.Name, "Sumoneó a " & Npclist(NpcIndex).Name & " en mapa " & .Pos.map)
120           End If
130       End With
End Sub


''
' Handle the "CreateNPCWithRespawn" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPCWithRespawn(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 12/24/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove Packet ID
60            Call .incomingData.ReadByte

              Dim NpcIndex As Integer

70            NpcIndex = .incomingData.ReadInteger()

80            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

90            NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True)

100           If NpcIndex <> 0 Then
110               Call LogGM(.Name, "Sumoneó con respawn " & Npclist(NpcIndex).Name & " en mapa " & .Pos.map)
120           End If
130       End With
End Sub

''
' Handle the "ImperialArmour" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleImperialArmour(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 12/24/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 4 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove Packet ID
60            Call .incomingData.ReadByte

              Dim Index As Byte
              Dim ObjIndex As Integer

70            Index = .incomingData.ReadByte()
80            ObjIndex = .incomingData.ReadInteger()

90            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

100           Select Case Index
              Case 1
110               ArmaduraImperial1 = ObjIndex

120           Case 2
130               ArmaduraImperial2 = ObjIndex

140           Case 3
150               ArmaduraImperial3 = ObjIndex

160           Case 4
170               TunicaMagoImperial = ObjIndex
180           End Select
190       End With
End Sub

''
' Handle the "ChaosArmour" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChaosArmour(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 12/24/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 4 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        With UserList(userindex)
              'Remove Packet ID
60            Call .incomingData.ReadByte

              Dim Index As Byte
              Dim ObjIndex As Integer

70            Index = .incomingData.ReadByte()
80            ObjIndex = .incomingData.ReadInteger()

90            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

100           Select Case Index
              Case 1
110               ArmaduraCaos1 = ObjIndex

120           Case 2
130               ArmaduraCaos2 = ObjIndex

140           Case 3
150               ArmaduraCaos3 = ObjIndex

160           Case 4
170               TunicaMagoCaos = ObjIndex
180           End Select
190       End With
End Sub

''
' Handle the "NavigateToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNavigateToggle(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 01/12/07
      '
      '***************************************************
10        With UserList(userindex)
              'Remove Packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

40            If .flags.Navegando = 1 Then
50                .flags.Navegando = 0
60            Else
70                .flags.Navegando = 1
80            End If

              'Tell the client that we are navigating.
90            Call WriteNavigateToggle(userindex)
100       End With
End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 12/24/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove Packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

40            If ServerSoloGMs > 0 Then
50                Call WriteConsoleMsg(userindex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
60                ServerSoloGMs = 0
70            Else
80                Call WriteConsoleMsg(userindex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
90                ServerSoloGMs = 1
100           End If
110       End With
End Sub

''
' Handle the "TurnOffServer" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnOffServer(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 12/24/06
      'Turns off the server
      '***************************************************
          Dim handle As Integer

10        With UserList(userindex)
              'Remove Packet ID
20            Call .incomingData.ReadByte

30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

40            Call LogGM(.Name, "/APAGAR")
50            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡¡¡" & .Name & " VA A APAGAR EL SERVIDOR!!!", FontTypeNames.FONTTYPE_FIGHT))

              'Log
60            handle = FreeFile
70            Open App.Path & "\logs\Main.log" For Append Shared As #handle

80            Print #handle, Date & " " & time & " server apagado por " & .Name & ". "

90            Close #handle

100           Unload frmMain
110       End With
End Sub

''
' Handle the "TurnCriminal" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnCriminal(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 12/26/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer

90            UserName = buffer.ReadASCIIString()

100           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
110               Call LogGM(.Name, "/CONDEN " & UserName)

120               tUser = NameIndex(UserName)
130               If tUser > 0 Then _
                     Call VolverCriminal(tUser)
140           End If

              'If we got here then packet is complete, copy data back to original queue
150           Call .incomingData.CopyBuffer(buffer)
160       End With

Errhandler:
          Dim error  As Long
170       error = Err.Number
180       On Error GoTo 0

          'Destroy auxiliar buffer
190       Set buffer = Nothing

200       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handle the "ResetFactions" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleResetFactionCaos(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 06/09/09
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer
              Dim Char As String

90            UserName = buffer.ReadASCIIString()

100           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And _
                  (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.ChaosCouncil)) Then
110               Call LogGM(.Name, "/PERDONARCAOS " & UserName)

120               tUser = NameIndex(UserName)

130               If tUser > 0 Then
140                   Call ResetFaccionCaos(tUser)
150               Else
160                   Char = CharPath & UserName & ".chr"

170                   If FileExist(Char, vbNormal) Then
180                       Call WriteVar(Char, "FACCIONES", "EjercitoReal", 0)
190                       Call WriteVar(Char, "FACCIONES", "CrimMatados", 0)
200                       Call WriteVar(Char, "FACCIONES", "EjercitoCaos", 0)
210                       Call WriteVar(Char, "FACCIONES", "FechaIngreso", "No ingresó a ninguna Facción")
220                       Call WriteVar(Char, "FACCIONES", "rArCaos", 0)
230                       Call WriteVar(Char, "FACCIONES", "rArReal", 0)
240                       Call WriteVar(Char, "FACCIONES", "rExCaos", 0)
250                       Call WriteVar(Char, "FACCIONES", "rExReal", 0)
260                       Call WriteVar(Char, "FACCIONES", "recCaos", 0)
270                       Call WriteVar(Char, "FACCIONES", "recReal", 0)
280                       Call WriteVar(Char, "FACCIONES", "Reenlistadas", 0)
290                       Call WriteVar(Char, "FACCIONES", "NivelIngreso", 0)
300                       Call WriteVar(Char, "FACCIONES", "MatadosIngreso", 0)
310                       Call WriteVar(Char, "FACCIONES", "NextRecompensa", 0)
320                   Else
330                       Call WriteConsoleMsg(userindex, "El personaje " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
340                   End If
350               End If
360           End If

              'If we got here then packet is complete, copy data back to original queue
370           Call .incomingData.CopyBuffer(buffer)
380       End With

Errhandler:
          Dim error  As Long
390       error = Err.Number
400       On Error GoTo 0

          'Destroy auxiliar buffer
410       Set buffer = Nothing

420       If error <> 0 Then _
             Err.Raise error
End Sub
''
' Handle the "ResetFactions" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleResetFactionReal(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 06/09/09
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim tUser As Integer
              Dim Char As String

90            UserName = buffer.ReadASCIIString()

100           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoyalCouncil)) Then
110               Call LogGM(.Name, "/PERDONARREAL " & UserName)

120               tUser = NameIndex(UserName)

130               If tUser > 0 Then
140                   Call ResetFaccionReal(tUser)
150               Else
160                   Char = CharPath & UserName & ".chr"

170                   If FileExist(Char, vbNormal) Then
180                       Call WriteVar(Char, "FACCIONES", "EjercitoReal", 0)
190                       Call WriteVar(Char, "FACCIONES", "CiudMatados", 0)
200                       Call WriteVar(Char, "FACCIONES", "EjercitoCaos", 0)
210                       Call WriteVar(Char, "FACCIONES", "FechaIngreso", "No ingresó a ninguna Facción")
220                       Call WriteVar(Char, "FACCIONES", "rArCaos", 0)
230                       Call WriteVar(Char, "FACCIONES", "rArReal", 0)
240                       Call WriteVar(Char, "FACCIONES", "rExCaos", 0)
250                       Call WriteVar(Char, "FACCIONES", "rExReal", 0)
260                       Call WriteVar(Char, "FACCIONES", "recCaos", 0)
270                       Call WriteVar(Char, "FACCIONES", "recReal", 0)
280                       Call WriteVar(Char, "FACCIONES", "Reenlistadas", 0)
290                       Call WriteVar(Char, "FACCIONES", "NivelIngreso", 0)
300                       Call WriteVar(Char, "FACCIONES", "MatadosIngreso", 0)
310                       Call WriteVar(Char, "FACCIONES", "NextRecompensa", 0)
320                   Else
330                       Call WriteConsoleMsg(userindex, "El personaje " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
340                   End If
350               End If
360           End If

              'If we got here then packet is complete, copy data back to original queue
370           Call .incomingData.CopyBuffer(buffer)
380       End With

Errhandler:
          Dim error  As Long
390       error = Err.Number
400       On Error GoTo 0

          'Destroy auxiliar buffer
410       Set buffer = Nothing

420       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handle the "RemoveCharFromGuild" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRemoveCharFromGuild(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 12/26/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim GuildIndex As Integer

90            UserName = buffer.ReadASCIIString()

100           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
110               Call LogGM(.Name, "/RAJARCLAN " & UserName)

120               GuildIndex = modGuilds.m_EcharMiembroDeClan(userindex, UserName)

130               If GuildIndex = 0 Then
140                   Call WriteConsoleMsg(userindex, "No pertenece a ningún clan o es fundador.", FontTypeNames.FONTTYPE_INFO)
150               Else
160                   Call WriteConsoleMsg(userindex, "Expulsado.", FontTypeNames.FONTTYPE_INFO)
170                   Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido expulsado del clan por los administradores del servidor.", FontTypeNames.FONTTYPE_GUILD))
180               End If
190           End If

              'If we got here then packet is complete, copy data back to original queue
200           Call .incomingData.CopyBuffer(buffer)
210       End With

Errhandler:
          Dim error  As Long
220       error = Err.Number
230       On Error GoTo 0

          'Destroy auxiliar buffer
240       Set buffer = Nothing

250       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handle the "RequestCharMail" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRequestCharMail(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 12/26/06
      'Request user mail
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String
              Dim mail As String

90            UserName = buffer.ReadASCIIString()

100           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
110               If FileExist(CharPath & UserName & ".chr") Then
120                   mail = GetVar(CharPath & UserName & ".chr", "CONTACTO", "email")

130                   Call WriteConsoleMsg(userindex, "Last email de " & UserName & ":" & mail, FontTypeNames.FONTTYPE_INFO)
140               End If
150           End If

              'If we got here then packet is complete, copy data back to original queue
160           Call .incomingData.CopyBuffer(buffer)
170       End With

Errhandler:
          Dim error  As Long
180       error = Err.Number
190       On Error GoTo 0

          'Destroy auxiliar buffer
200       Set buffer = Nothing

210       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handle the "SystemMessage" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/29/06
      'Send a message to all the users
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim message As String
90            message = buffer.ReadASCIIString()

100           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
110               Call LogGM(.Name, "Mensaje de sistema:" & message)

120               Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(message))
130           End If

              'If we got here then packet is complete, copy data back to original queue
140           Call .incomingData.CopyBuffer(buffer)
150       End With

Errhandler:
          Dim error  As Long
160       error = Err.Number
170       On Error GoTo 0

          'Destroy auxiliar buffer
180       Set buffer = Nothing

190       If error <> 0 Then _
             Err.Raise error
End Sub

''
' Handle the "Ping" message
'
' @param userIndex The index of the user sending the message

Public Sub HandlePing(ByVal userindex As Integer)
      '***************************************************
      'Author: Lucas Tavolaro Ortiz (Tavo)
      'Last Modification: 12/24/06
      'Show guilds messages
      '***************************************************
10        With UserList(userindex)
              'Remove Packet ID
20            Call .incomingData.ReadByte
              
30            If .Counters.TimePin > 0 Then Exit Sub
              
40            .Counters.TimePin = 4
50            Call WritePong(userindex)
60        End With
End Sub

''
' Handle the "SetIniVar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetIniVar(ByVal userindex As Integer)
      '***************************************************
      'Author: Brian Chaia (BrianPr)
      'Last Modification: 01/23/10 (Marco)
      'Modify server.ini
      '***************************************************
10        If UserList(userindex).incomingData.length < 6 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50    On Error GoTo Errhandler

60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
              
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim sLlave As String
              Dim sClave As String
              Dim sValor As String

              'Obtengo los parámetros
90            sLlave = buffer.ReadASCIIString()
100           sClave = buffer.ReadASCIIString()
110           sValor = buffer.ReadASCIIString()

120           If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
                  Dim sTmp As String

                  'No podemos modificar [INIT]Dioses ni [Dioses]*
130               If (UCase$(sLlave) = "INIT" And UCase$(sClave) = "ADMINES") Or UCase$(sLlave) = "ADMINES" Then
140                   Call WriteConsoleMsg(userindex, "¡No puedes modificar esa información desde aquí!", FontTypeNames.FONTTYPE_INFO)
150               If (UCase$(sLlave) = "INIT" And UCase$(sClave) = "DIOSES") Or UCase$(sLlave) = "DIOSES" Then
160                   Call WriteConsoleMsg(userindex, "¡No puedes modificar esa información desde aquí!", FontTypeNames.FONTTYPE_INFO)
170               Else
                      'Obtengo el valor según llave y clave
180                   sTmp = GetVar(IniPath & "Server.ini", sLlave, sClave)

                      'Si obtengo un valor escribo en el server.ini
190                   If LenB(sTmp) Then
200                       Call WriteVar(IniPath & "Server.ini", sLlave, sClave, sValor)
210                       Call LogGM(.Name, "Modificó en server.ini (" & sLlave & " " & sClave & ") el valor " & sTmp & " por " & sValor)
220                       Call WriteConsoleMsg(userindex, "Modificó " & sLlave & " " & sClave & " a " & sValor & ". Valor anterior " & sTmp, FontTypeNames.FONTTYPE_INFO)
230                   Else
240                       Call WriteConsoleMsg(userindex, "No existe la llave y/o clave", FontTypeNames.FONTTYPE_INFO)
250                   End If
260               End If
270               End If
280           End If

              'If we got here then packet is complete, copy data back to original queue
290           Call .incomingData.CopyBuffer(buffer)
300       End With

Errhandler:
          Dim error As Long

310       error = Err.Number

320   On Error GoTo 0
          'Destroy auxiliar buffer
330       Set buffer = Nothing

340       If error <> 0 Then _
              Err.Raise error
End Sub

''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal userindex As Integer)
   On Error GoTo WriteLoggedMessage_Error

10        On Error GoTo Errhandler
20          Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.Logged)
            Call UserList(userindex).outgoingData.WriteByte(UserList(userindex).ClaveAC)
30        Exit Sub
Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteLoggedMessage_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteLoggedMessage of Módulo Protocol in line " & Erl
End Sub


''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteRemoveAllDialogs_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.RemoveDialogs)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteRemoveAllDialogs_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteRemoveAllDialogs of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal userindex As Integer, ByVal CharIndex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteRemoveCharDialog_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageRemoveCharDialog(CharIndex))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteRemoveCharDialog_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteRemoveCharDialog of Módulo Protocol in line " & Erl
End Sub
Public Function PrepareMessageCreateDamage(ByVal X As Byte, ByVal Y As Byte, ByVal DamageValue As Long, ByVal DamageType As Byte)

      ' @ Envia el paquete para crear daño (Y)

10        With auxiliarBuffer
20            .WriteByte ServerPacketID.CreateDamage
30            .WriteByte X
40            .WriteByte Y
50            .WriteLong DamageValue
60            .WriteByte DamageType

70            PrepareMessageCreateDamage = .ReadASCIIStringFixed(.length)

80        End With

End Function

''
' Writes the "NavigateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "NavigateToggle" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteNavigateToggle_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.NavigateToggle)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteNavigateToggle_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteNavigateToggle of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Disconnect" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteDisconnect_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.Disconnect)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteDisconnect_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteDisconnect of Módulo Protocol in line " & Erl
End Sub
Public Sub WriteMontateToggle(ByVal userindex As Integer)

          Dim obData As ObjData
   On Error GoTo WriteMontateToggle_Error
          
          
          If UserList(userindex).Invent.MonturaObjIndex > 0 Then
20            obData = ObjData(UserList(userindex).Invent.MonturaObjIndex)
              
              
30            Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.MontateToggle)
40            Call UserList(userindex).outgoingData.WriteByte(obData.Velocidad)

          Else
          
             LogError "El personaje " & UserList(userindex).Name & " envió en WriteMontateToggle una monturaindex<0"
          End If

   On Error GoTo 0
   Exit Sub

WriteMontateToggle_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteMontateToggle of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "UserOfferConfirm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserOfferConfirm(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 14/12/2009
      'Writes the "UserOfferConfirm" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteUserOfferConfirm_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.UserOfferConfirm)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteUserOfferConfirm_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUserOfferConfirm of Módulo Protocol in line " & Erl
End Sub


''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CommerceEnd" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteCommerceEnd_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.CommerceEnd)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteCommerceEnd_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteCommerceEnd of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "BankEnd" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteBankEnd_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.BankEnd)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteBankEnd_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteBankEnd of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CommerceInit" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteCommerceInit_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.CommerceInit)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteCommerceInit_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteCommerceInit of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "BankInit" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteBankInit_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.BankInit)
30        Call UserList(userindex).outgoingData.WriteLong(UserList(userindex).Stats.Banco)
40        Exit Sub

Errhandler:
50        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
60            Call FlushBuffer(userindex)
70            Resume
80        End If

   On Error GoTo 0
   Exit Sub

WriteBankInit_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteBankInit of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteUserCommerceInit_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.UserCommerceInit)
30        Call UserList(userindex).outgoingData.WriteASCIIString(UserList(userindex).ComUsu.DestNick)
40        Exit Sub

Errhandler:
50        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
60            Call FlushBuffer(userindex)
70            Resume
80        End If

   On Error GoTo 0
   Exit Sub

WriteUserCommerceInit_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUserCommerceInit of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteUserCommerceEnd_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.UserCommerceEnd)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteUserCommerceEnd_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUserCommerceEnd of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowBlacksmithForm(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteShowBlacksmithForm_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.ShowBlacksmithForm)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteShowBlacksmithForm_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteShowBlacksmithForm of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowCarpenterForm(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteShowCarpenterForm_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.ShowCarpenterForm)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteShowCarpenterForm_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteShowCarpenterForm of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UpdateMana" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteUpdateSta_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.UpdateSta)
40            Call .WriteInteger(UserList(userindex).Stats.MinSta)
50        End With
60        Exit Sub

Errhandler:
70        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
80            Call FlushBuffer(userindex)
90            Resume
100       End If

   On Error GoTo 0
   Exit Sub

WriteUpdateSta_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUpdateSta of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UpdateMana" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteUpdateMana_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.UpdateMana)
40            Call .WriteInteger(UserList(userindex).Stats.MinMAN)
50        End With
60        Exit Sub

Errhandler:
70        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
80            Call FlushBuffer(userindex)
90            Resume
100       End If

   On Error GoTo 0
   Exit Sub

WriteUpdateMana_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUpdateMana of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UpdateMana" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteUpdateHP_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.UpdateHP)
40            Call .WriteInteger(UserList(userindex).Stats.MinHp)
50        End With
60        Exit Sub

Errhandler:
70        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
80            Call FlushBuffer(userindex)
90            Resume
100       End If

   On Error GoTo 0
   Exit Sub

WriteUpdateHP_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUpdateHP of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UpdateGold" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteUpdateGold_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.UpdateGold)
40            Call .WriteLong(UserList(userindex).Stats.Gld)
50        End With
60        Exit Sub

Errhandler:
70        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
80            Call FlushBuffer(userindex)
90            Resume
100       End If

   On Error GoTo 0
   Exit Sub

WriteUpdateGold_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUpdateGold of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "UpdateBankGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateBankGold(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 14/12/2009
      'Writes the "UpdateBankGold" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteUpdateBankGold_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.UpdateBankGold)
40            Call .WriteLong(UserList(userindex).Stats.Banco)
50        End With
60        Exit Sub

Errhandler:
70        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
80            Call FlushBuffer(userindex)
90            Resume
100       End If

   On Error GoTo 0
   Exit Sub

WriteUpdateBankGold_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUpdateBankGold of Módulo Protocol in line " & Erl
End Sub


''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UpdateExp" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteUpdateExp_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.UpdateExp)
40            Call .WriteLong(UserList(userindex).Stats.Exp)
50        End With
60        Exit Sub

Errhandler:
70        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
80            Call FlushBuffer(userindex)
90            Resume
100       End If

   On Error GoTo 0
   Exit Sub

WriteUpdateExp_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUpdateExp of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenghtAndDexterity(ByVal userindex As Integer)
      '***************************************************
      'Author: Budi
      'Last Modification: 11/26/09
      'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteUpdateStrenghtAndDexterity_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.UpdateStrenghtAndDexterity)
40            Call .WriteByte(UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza))
50            Call .WriteByte(UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad))
60        End With
70        Exit Sub

Errhandler:
80        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
90            Call FlushBuffer(userindex)
100           Resume
110       End If

   On Error GoTo 0
   Exit Sub

WriteUpdateStrenghtAndDexterity_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUpdateStrenghtAndDexterity of Módulo Protocol in line " & Erl
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateDexterity(ByVal userindex As Integer)
      '***************************************************
      'Author: Budi
      'Last Modification: 11/26/09
      'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteUpdateDexterity_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.UpdateDexterity)
40            Call .WriteByte(UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad))
50        End With
60        Exit Sub

Errhandler:
70        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
80            Call FlushBuffer(userindex)
90            Resume
100       End If

   On Error GoTo 0
   Exit Sub

WriteUpdateDexterity_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUpdateDexterity of Módulo Protocol in line " & Erl
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenght(ByVal userindex As Integer)
      '***************************************************
      'Author: Budi
      'Last Modification: 11/26/09
      'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteUpdateStrenght_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.UpdateStrenght)
40            Call .WriteByte(UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza))
50        End With
60        Exit Sub

Errhandler:
70        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
80            Call FlushBuffer(userindex)
90            Resume
100       End If

   On Error GoTo 0
   Exit Sub

WriteUpdateStrenght_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUpdateStrenght of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal userindex As Integer, ByVal map As Integer, ByVal Version As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ChangeMap" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteChangeMap_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.ChangeMap)
40            Call .WriteInteger(map)
50            Call .WriteASCIIString(MapInfo(map).Name)
60            Call .WriteInteger(Version)
70        End With
80        Exit Sub

Errhandler:
90        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
100           Call FlushBuffer(userindex)
110           Resume
120       End If

   On Error GoTo 0
   Exit Sub

WriteChangeMap_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteChangeMap of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "PosUpdate" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WritePosUpdate_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.PosUpdate)
40            Call .WriteByte(UserList(userindex).Pos.X)
50            Call .WriteByte(UserList(userindex).Pos.Y)
            
60        End With
70        Exit Sub

Errhandler:
80        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
90            Call FlushBuffer(userindex)
100           Resume
110       End If

   On Error GoTo 0
   Exit Sub

WritePosUpdate_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WritePosUpdate of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatOverHead(ByVal userindex As Integer, ByVal chat As String, ByVal CharIndex As Integer, ByVal color As Long)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ChatOverHead" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteChatOverHead_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHead(chat, CharIndex, color))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteChatOverHead_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteChatOverHead of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal userindex As Integer, ByVal chat As String, ByVal FontIndex As FontTypeNames, Optional ByVal Crlf As Boolean = True)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteConsoleMsg_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(chat & IIf(Crlf = False, "`", ""), FontIndex))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteConsoleMsg_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteConsoleMsg of Módulo Protocol in line " & Erl
End Sub

Public Sub WriteCommerceChat(ByVal userindex As Integer, ByVal chat As String, ByVal FontIndex As FontTypeNames)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 05/17/06
      'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteCommerceChat_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareCommerceConsoleMsg(chat, FontIndex))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteCommerceChat_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteCommerceChat of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "GuildChat" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildChat(ByVal userindex As Integer, ByVal chat As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildChat" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteGuildChat_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageGuildChat(chat))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteGuildChat_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteGuildChat of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMessageBox(ByVal userindex As Integer, ByVal message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteShowMessageBox_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.ShowMessageBox)
40            Call .WriteASCIIString(message)
50        End With
60        Exit Sub

Errhandler:
70        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
80            Call FlushBuffer(userindex)
90            Resume
100       End If

   On Error GoTo 0
   Exit Sub

WriteShowMessageBox_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteShowMessageBox of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteUserIndexInServer_Error

10        On Error GoTo Errhandler
          
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.UserIndexInServer)
40            Call .WriteInteger(userindex)
50        End With
60        Exit Sub

Errhandler:
70        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
80            Call FlushBuffer(userindex)
90            Resume
100       End If

   On Error GoTo 0
   Exit Sub

WriteUserIndexInServer_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUserIndexInServer of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteUserCharIndexInServer_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.UserCharIndexInServer)
40            Call .WriteInteger(UserList(userindex).Char.CharIndex)
50        End With
60        Exit Sub

Errhandler:
70        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
80            Call FlushBuffer(userindex)
90            Resume
100       End If

   On Error GoTo 0
   Exit Sub

WriteUserCharIndexInServer_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUserCharIndexInServer of Módulo Protocol in line " & Erl
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

Public Sub WriteCharacterCreate(ByVal userindex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal Heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Name As String, ByVal NickColor As Byte, _
                                ByVal Privileges As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CharacterCreate" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteCharacterCreate_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterCreate(body, Head, Heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, _
                                                                                                    helmet, Name, NickColor, Privileges))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteCharacterCreate_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteCharacterCreate of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal userindex As Integer, ByVal CharIndex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CharacterRemove" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteCharacterRemove_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(CharIndex))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteCharacterRemove_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteCharacterRemove of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal userindex As Integer, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CharacterMove" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteCharacterMove_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterMove(CharIndex, X, Y))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteCharacterMove_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteCharacterMove of Módulo Protocol in line " & Erl
End Sub

Public Sub WriteForceCharMove(ByVal userindex, ByVal Direccion As eHeading)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 26/03/2009
      'Writes the "ForceCharMove" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteForceCharMove_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageForceCharMove(Direccion))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteForceCharMove_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteForceCharMove of Módulo Protocol in line " & Erl
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

Public Sub WriteCharacterChange(ByVal userindex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal Heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CharacterChange" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteCharacterChange_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChange(body, Head, Heading, CharIndex, weapon, shield, FX, FXLoops, helmet))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteCharacterChange_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteCharacterChange of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectCreate(ByVal userindex As Integer, ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ObjectCreate" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteObjectCreate_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectCreate(GrhIndex, X, Y))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteObjectCreate_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteObjectCreate of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal userindex As Integer, ByVal X As Byte, ByVal Y As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ObjectDelete" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteObjectDelete_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectDelete(X, Y))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteObjectDelete_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteObjectDelete of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal userindex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "BlockPosition" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteBlockPosition_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.BlockPosition)
40            Call .WriteByte(X)
50            Call .WriteByte(Y)
60            Call .WriteBoolean(Blocked)
70        End With
80        Exit Sub

Errhandler:
90        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
100           Call FlushBuffer(userindex)
110           Resume
120       End If

   On Error GoTo 0
   Exit Sub

WriteBlockPosition_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteBlockPosition of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "PlayMidi" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMidi(ByVal userindex As Integer, ByVal midi As Byte, Optional ByVal loops As Integer = -1)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "PlayMidi" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WritePlayMidi_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayMidi(midi, loops))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WritePlayMidi_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WritePlayMidi of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "PlayWave" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayWave(ByVal userindex As Integer, ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 08/08/07
      'Last Modified by: Rapsodius
      'Added X and Y positions for 3D Sounds
      '***************************************************
   On Error GoTo WritePlayWave_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave, X, Y))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WritePlayWave_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WritePlayWave of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "GuildList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildList(ByVal userindex As Integer, ByRef guildList() As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildList" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteGuildList_Error

10        On Error GoTo Errhandler
          Dim tmp    As String
          Dim i      As Long

20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.guildList)

              ' Prepare guild name's list
40            For i = LBound(guildList()) To UBound(guildList())
50                tmp = tmp & guildList(i) & SEPARATOR
60            Next i

70            If Len(tmp) Then _
                 tmp = Left$(tmp, Len(tmp) - 1)

80            Call .WriteASCIIString(tmp)
90        End With
100       Exit Sub

Errhandler:
110       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
120           Call FlushBuffer(userindex)
130           Resume
140       End If

   On Error GoTo 0
   Exit Sub

WriteGuildList_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteGuildList of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAreaChanged(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "AreaChanged" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteAreaChanged_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.AreaChanged)
40            Call .WriteByte(UserList(userindex).Pos.X)
50            Call .WriteByte(UserList(userindex).Pos.Y)
60        End With
70        Exit Sub

Errhandler:
80        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
90            Call FlushBuffer(userindex)
100           Resume
110       End If

   On Error GoTo 0
   Exit Sub

WriteAreaChanged_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteAreaChanged of Módulo Protocol in line " & Erl
End Sub


Public Sub writeClave(ByVal userindex As Integer)
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Manda la clave de encriptacion al cliente
      '***************************************************
   On Error GoTo errh

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.SendClave)
        
            Call UserList(userindex).outgoingData.WriteByte(UserList(userindex).ClaveAC)
            Call FlushBuffer(userindex)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

errh:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure writeClave of Módulo Protocol in line " & Erl
End Sub


''
' Writes the "PauseToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "PauseToggle" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WritePauseToggle_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle())
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WritePauseToggle_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WritePauseToggle of Módulo Protocol in line " & Erl
End Sub


''
' Writes the "CreateFX" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal userindex As Integer, ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CreateFX" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteCreateFX_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(CharIndex, FX, FXLoops))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteCreateFX_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteCreateFX of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteUpdateUserStats_Error

10        On Error GoTo Errhandler
          
20            With UserList(userindex).outgoingData
30                Call .WriteByte(ServerPacketID.UpdateUserStats)
40                Call .WriteInteger(UserList(userindex).Stats.MaxHp)
50                Call .WriteInteger(UserList(userindex).Stats.MinHp)
60                Call .WriteInteger(UserList(userindex).Stats.MaxMAN)
70                Call .WriteInteger(UserList(userindex).Stats.MinMAN)
80                Call .WriteInteger(UserList(userindex).Stats.MaxSta)
90                Call .WriteInteger(UserList(userindex).Stats.MinSta)
100               Call .WriteLong(UserList(userindex).Stats.Gld)
110               Call .WriteByte(UserList(userindex).Stats.ELV)
120               Call .WriteLong(UserList(userindex).Stats.ELU)
130               Call .WriteLong(UserList(userindex).Stats.Exp)
140               .WriteByte (UserList(userindex).flags.Oculto)
150               Call .WriteBoolean(UserList(userindex).flags.ModoCombate)
160               .WriteInteger UserList(userindex).Char.CharIndex
170           End With
              
180       Exit Sub

Errhandler:
190       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
200           Call FlushBuffer(userindex)
210           Resume
220       End If

   On Error GoTo 0
   Exit Sub

WriteUpdateUserStats_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUpdateUserStats of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkRequestTarget(ByVal userindex As Integer, ByVal Skill As eSkill)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "WorkRequestTarget" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteWorkRequestTarget_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.WorkRequestTarget)
40            Call .WriteByte(Skill)
50        End With
60        Exit Sub

Errhandler:
70        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
80            Call FlushBuffer(userindex)
90            Resume
100       End If

   On Error GoTo 0
   Exit Sub

WriteWorkRequestTarget_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteWorkRequestTarget of Módulo Protocol in line " & Erl
End Sub

Public Sub WriteChangeInventorySlot(ByVal userindex As Integer, ByVal Slot As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 3/12/09
      'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
      '3/12/09: Budi - Ahora se envia MaxDef y MinDef en lugar de Def
      '***************************************************
   On Error GoTo WriteChangeInventorySlot_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.ChangeInventorySlot)
40            Call .WriteByte(Slot)

              Dim ObjIndex As Integer
              Dim obData As ObjData

50            ObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
              
60            Call .WriteInteger(ObjIndex)
70            Call .WriteInteger(UserList(userindex).Invent.Object(Slot).Amount)
80            Call .WriteBoolean(UserList(userindex).Invent.Object(Slot).Equipped)
              
90            If ObjIndex > 0 Then
100               obData = ObjData(ObjIndex)
110               Call .WriteInteger(obData.GrhIndex)
120               Call .WriteByte(obData.ObjType)
130               Call .WriteInteger(obData.MaxHIT)
140               Call .WriteInteger(obData.MinHIT)
150               Call .WriteInteger(obData.MaxDef)
160               Call .WriteInteger(obData.MinDef)
170               Call .WriteSingle(SalePrice(ObjIndex))
180           End If
190       End With
          
200       Exit Sub

Errhandler:
210       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
220           Call FlushBuffer(userindex)
230           Resume
240       End If

   On Error GoTo 0
   Exit Sub

WriteChangeInventorySlot_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteChangeInventorySlot of Módulo Protocol in line " & Erl
End Sub



''
' Writes the "ChangeBankSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal userindex As Integer, ByVal Slot As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 12/03/09
      'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
      '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
      '***************************************************
   On Error GoTo WriteChangeBankSlot_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.ChangeBankSlot)
40            Call .WriteByte(Slot)

              Dim ObjIndex As Integer
              Dim obData As ObjData

50            ObjIndex = UserList(userindex).BancoInvent.Object(Slot).ObjIndex

60            Call .WriteInteger(ObjIndex)
70            Call .WriteInteger(UserList(userindex).BancoInvent.Object(Slot).Amount)

80            If ObjIndex > 0 Then
90                obData = ObjData(ObjIndex)
100               Call .WriteInteger(obData.GrhIndex)
110               Call .WriteByte(obData.ObjType)
120               Call .WriteInteger(obData.MaxHIT)
130               Call .WriteInteger(obData.MinHIT)
140               Call .WriteInteger(obData.MaxDef)
150               Call .WriteInteger(obData.MinDef)
160               Call .WriteLong(obData.valor)
170           End If
180       End With
190       Exit Sub

Errhandler:
200       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
210           Call FlushBuffer(userindex)
220           Resume
230       End If

   On Error GoTo 0
   Exit Sub

WriteChangeBankSlot_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteChangeBankSlot of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal userindex As Integer, ByVal Slot As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
      ' /ver paquete mañana by lautaro
      '***************************************************
   On Error GoTo WriteChangeSpellSlot_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.ChangeSpellSlot)
40            Call .WriteByte(Slot)
50            Call .WriteInteger(UserList(userindex).Stats.UserHechizos(Slot))

60            If UserList(userindex).Stats.UserHechizos(Slot) > 0 Then
70                Call .WriteASCIIString(Hechizos(UserList(userindex).Stats.UserHechizos(Slot)).Nombre)
80            Else
90                Call .WriteASCIIString("(Vacio)")
100           End If
110       End With
120       Exit Sub

Errhandler:
130       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
140           Call FlushBuffer(userindex)
150           Resume
160       End If

   On Error GoTo 0
   Exit Sub

WriteChangeSpellSlot_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteChangeSpellSlot of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Atributes" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteAttributes_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.Atributes)
40            Call .WriteByte(UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza))
50            Call .WriteByte(UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad))
60            Call .WriteByte(UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia))
70            Call .WriteByte(UserList(userindex).Stats.UserAtributos(eAtributos.Carisma))
80            Call .WriteByte(UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion))
90        End With
100       Exit Sub

Errhandler:
110       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
120           Call FlushBuffer(userindex)
130           Resume
140       End If

   On Error GoTo 0
   Exit Sub

WriteAttributes_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteAttributes of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithWeapons(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
      'Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteBlacksmithWeapons_Error

10        On Error GoTo Errhandler
          Dim i      As Long
          Dim Obj    As ObjData
          Dim validIndexes() As Integer
          Dim Count  As Integer

20        ReDim validIndexes(1 To UBound(ArmasHerrero()))

30        With UserList(userindex).outgoingData
40            Call .WriteByte(ServerPacketID.BlacksmithWeapons)

50            For i = 1 To UBound(ArmasHerrero())
                  ' Can the user create this object? If so add it to the list....
60                If ObjData(ArmasHerrero(i)).SkHerreria <= Round(UserList(userindex).Stats.UserSkills(eSkill.herreria) / ModHerreriA(UserList(userindex).clase), 0) Then
70                    Count = Count + 1
80                    validIndexes(Count) = i
90                End If
100           Next i

              ' Write the number of objects in the list
110           Call .WriteInteger(Count)

              ' Write the needed data of each object
120           For i = 1 To Count
130               Obj = ObjData(ArmasHerrero(validIndexes(i)))
140               Call .WriteASCIIString(Obj.Name)
150               Call .WriteInteger(Obj.LingH)
160               Call .WriteInteger(Obj.LingP)
170               Call .WriteInteger(Obj.LingO)
180               Call .WriteInteger(ArmasHerrero(validIndexes(i)))
190           Next i
200       End With
210       Exit Sub

Errhandler:
220       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
230           Call FlushBuffer(userindex)
240           Resume
250       End If

   On Error GoTo 0
   Exit Sub

WriteBlacksmithWeapons_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteBlacksmithWeapons of Módulo Protocol in line " & Erl
End Sub


''
' Writes the "BlacksmithArmors" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Sub WriteBlacksmithArmors(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
      'Writes the "BlacksmithArmors" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteBlacksmithArmors_Error

10        On Error GoTo Errhandler
          Dim i      As Long
          Dim Obj    As ObjData
          Dim validIndexes() As Integer
          Dim Count  As Integer

20        ReDim validIndexes(1 To UBound(ArmadurasHerrero()))

30        With UserList(userindex).outgoingData
40            Call .WriteByte(ServerPacketID.BlacksmithArmors)

50            For i = 1 To UBound(ArmadurasHerrero())
                  ' Can the user create this object? If so add it to the list....
60                If ObjData(ArmadurasHerrero(i)).SkHerreria <= Round(UserList(userindex).Stats.UserSkills(eSkill.herreria) / ModHerreriA(UserList(userindex).clase), 0) Then
70                    Count = Count + 1
80                    validIndexes(Count) = i
90                End If
100           Next i

              ' Write the number of objects in the list
110           Call .WriteInteger(Count)

              ' Write the needed data of each object
120           For i = 1 To Count
130               Obj = ObjData(ArmadurasHerrero(validIndexes(i)))
140               Call .WriteASCIIString(Obj.Name)
150               Call .WriteInteger(Obj.LingH)
160               Call .WriteInteger(Obj.LingP)
170               Call .WriteInteger(Obj.LingO)
180               Call .WriteInteger(ArmadurasHerrero(validIndexes(i)))
190           Next i
200       End With
210       Exit Sub

Errhandler:
220       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
230           Call FlushBuffer(userindex)
240           Resume
250       End If

   On Error GoTo 0
   Exit Sub

WriteBlacksmithArmors_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteBlacksmithArmors of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCarpenterObjects(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteCarpenterObjects_Error

10        On Error GoTo Errhandler
          Dim i      As Long
          Dim Obj    As ObjData
          Dim validIndexes() As Integer
          Dim Count  As Integer

20        ReDim validIndexes(1 To UBound(ObjCarpintero()))

30        With UserList(userindex).outgoingData
40            Call .WriteByte(ServerPacketID.CarpenterObjects)

50            For i = 1 To UBound(ObjCarpintero())
                  ' Can the user create this object? If so add it to the list....
60                If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(userindex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(userindex).clase) Then
70                    Count = Count + 1
80                    validIndexes(Count) = i
90                End If
100           Next i

              ' Write the number of objects in the list
110           Call .WriteInteger(Count)

              ' Write the needed data of each object
120           For i = 1 To Count
130               Obj = ObjData(ObjCarpintero(validIndexes(i)))
140               Call .WriteASCIIString(Obj.Name)
150               Call .WriteInteger(Obj.Madera)
160               Call .WriteInteger(ObjCarpintero(validIndexes(i)))
170           Next i
180       End With
190       Exit Sub

Errhandler:
200       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
210           Call FlushBuffer(userindex)
220           Resume
230       End If

   On Error GoTo 0
   Exit Sub

WriteCarpenterObjects_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteCarpenterObjects of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "RestOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RestOK" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteRestOK_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.RestOK)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteRestOK_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteRestOK of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteErrorMsg(ByVal userindex As Integer, ByVal message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ErrorMsg" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteErrorMsg_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageErrorMsg(message))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteErrorMsg_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteErrorMsg of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "Blind" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Blind" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteBlind_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.Blind)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteBlind_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteBlind of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "Dumb" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Dumb" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteDumb_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.Dumb)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteDumb_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteDumb of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSignal(ByVal userindex As Integer, ByVal ObjIndex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ShowSignal" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteShowSignal_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.ShowSignal)
40            Call .WriteASCIIString(ObjData(ObjIndex).Texto)
50            Call .WriteInteger(ObjData(ObjIndex).GrhSecundario)
60        End With
70        Exit Sub

Errhandler:
80        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
90            Call FlushBuffer(userindex)
100           Resume
110       End If

   On Error GoTo 0
   Exit Sub

WriteShowSignal_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteShowSignal of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal userindex As Integer, ByVal Slot As Byte, ByRef Obj As Obj, ByVal price As Single)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 12/03/09
      'Last Modified by: Budi
      'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
      '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
      '***************************************************
   On Error GoTo WriteChangeNPCInventorySlot_Error

10        On Error GoTo Errhandler
          Dim ObjInfo As ObjData

20        If Obj.ObjIndex >= LBound(ObjData()) And Obj.ObjIndex <= UBound(ObjData()) Then
30            ObjInfo = ObjData(Obj.ObjIndex)
40        End If

50        With UserList(userindex).outgoingData
60            Call .WriteByte(ServerPacketID.ChangeNPCInventorySlot)
70            Call .WriteByte(Slot)
80            Call .WriteInteger(Obj.ObjIndex)
              
90            If Obj.ObjIndex > 0 Then
100               Call .WriteInteger(Obj.Amount)
110               Call .WriteSingle(price)
120               Call .WriteByte(ObjInfo.copaS)
130               Call .WriteByte(ObjInfo.Eldhir)
140               Call .WriteInteger(ObjInfo.GrhIndex)
                  
150               Call .WriteByte(ObjInfo.ObjType)
160               Call .WriteInteger(ObjInfo.MaxHIT)
170               Call .WriteInteger(ObjInfo.MinHIT)
180               Call .WriteInteger(ObjInfo.MaxDef)
190               Call .WriteInteger(ObjInfo.MinDef)
200           End If
210       End With
220       Exit Sub

Errhandler:
230       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
240           Call FlushBuffer(userindex)
250           Resume
260       End If

   On Error GoTo 0
   Exit Sub

WriteChangeNPCInventorySlot_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteChangeNPCInventorySlot of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteUpdateHungerAndThirst_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.UpdateHungerAndThirst)
40            Call .WriteByte(UserList(userindex).Stats.MaxAGU)
50            Call .WriteByte(UserList(userindex).Stats.MinAGU)
60            Call .WriteByte(UserList(userindex).Stats.MaxHam)
70            Call .WriteByte(UserList(userindex).Stats.MinHam)
80        End With
90        Exit Sub

Errhandler:
100       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
110           Call FlushBuffer(userindex)
120           Resume
130       End If

   On Error GoTo 0
   Exit Sub

WriteUpdateHungerAndThirst_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUpdateHungerAndThirst of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "Fame" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteFame(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Fame" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteFame_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.Fame)

40            Call .WriteLong(UserList(userindex).Reputacion.AsesinoRep)
50            Call .WriteLong(UserList(userindex).Reputacion.BandidoRep)
60            Call .WriteLong(UserList(userindex).Reputacion.BurguesRep)
70            Call .WriteLong(UserList(userindex).Reputacion.LadronesRep)
80            Call .WriteLong(UserList(userindex).Reputacion.NobleRep)
90            Call .WriteLong(UserList(userindex).Reputacion.PlebeRep)
100           Call .WriteLong(UserList(userindex).Reputacion.Promedio)
110       End With
120       Exit Sub

Errhandler:
130       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
140           Call FlushBuffer(userindex)
150           Resume
160       End If

   On Error GoTo 0
   Exit Sub

WriteFame_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteFame of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "MiniStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "MiniStats" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteMiniStats_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.MiniStats)

40            Call .WriteLong(UserList(userindex).Faccion.CiudadanosMatados)
50            Call .WriteLong(UserList(userindex).Faccion.CriminalesMatados)

              'TODO : Este valor es calculable, no debería NI EXISTIR, ya sea en el servidor ni en el cliente!!!
60            Call .WriteLong(UserList(userindex).Stats.UsuariosMatados)

70            Call .WriteInteger(UserList(userindex).Stats.NPCsMuertos)

80            Call .WriteByte(UserList(userindex).clase)
90            Call .WriteLong(UserList(userindex).Counters.Pena)
100       End With
110       Exit Sub

Errhandler:
120       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
130           Call FlushBuffer(userindex)
140           Resume
150       End If

   On Error GoTo 0
   Exit Sub

WriteMiniStats_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteMiniStats of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data buffer.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLevelUp(ByVal userindex As Integer, ByVal skillPoints As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "LevelUp" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteLevelUp_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.LevelUp)
40            Call .WriteInteger(skillPoints)
50        End With
60        Exit Sub

Errhandler:
70        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
80            Call FlushBuffer(userindex)
90            Resume
100       End If

   On Error GoTo 0
   Exit Sub

WriteLevelUp_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteLevelUp of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "AddForumMsg" message to the given user's outgoing data buffer.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAddForumMsg(ByVal userindex As Integer, ByVal ForumType As eForumType, _
                            ByRef Title As String, ByRef Author As String, ByRef message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 02/01/2010
      'Writes the "AddForumMsg" message to the given user's outgoing data buffer
      '02/01/2010: ZaMa - Now sends Author and forum type
      '***************************************************
   On Error GoTo WriteAddForumMsg_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.AddForumMsg)
40            Call .WriteByte(ForumType)
50            Call .WriteASCIIString(Title)
60            Call .WriteASCIIString(Author)
70            Call .WriteASCIIString(message)
80        End With
90        Exit Sub

Errhandler:
100       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
110           Call FlushBuffer(userindex)
120           Resume
130       End If

   On Error GoTo 0
   Exit Sub

WriteAddForumMsg_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteAddForumMsg of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowForumForm(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ShowForumForm" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteShowForumForm_Error

10        On Error GoTo Errhandler

          Dim Visibilidad As Byte
          Dim CanMakeSticky As Byte

20        With UserList(userindex)
30            Call .outgoingData.WriteByte(ServerPacketID.ShowForumForm)

40            Visibilidad = eForumVisibility.ieGENERAL_MEMBER

50            If esCaos(userindex) Or EsGm(userindex) Then
60                Visibilidad = Visibilidad Or eForumVisibility.ieCAOS_MEMBER
70            End If

80            If esArmada(userindex) Or EsGm(userindex) Then
90                Visibilidad = Visibilidad Or eForumVisibility.ieREAL_MEMBER
100           End If

110           Call .outgoingData.WriteByte(Visibilidad)

              ' Pueden mandar sticky los gms o los del consejo de armada/caos
120           If EsGm(userindex) Then
130               CanMakeSticky = 2
140           ElseIf (.flags.Privilegios And PlayerType.ChaosCouncil) <> 0 Then
150               CanMakeSticky = 1
160           ElseIf (.flags.Privilegios And PlayerType.RoyalCouncil) <> 0 Then
170               CanMakeSticky = 1
180           End If

190           Call .outgoingData.WriteByte(CanMakeSticky)
200       End With
210       Exit Sub

Errhandler:
220       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
230           Call FlushBuffer(userindex)
240           Resume
250       End If

   On Error GoTo 0
   Exit Sub

WriteShowForumForm_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteShowForumForm of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal userindex As Integer, ByVal CharIndex As Integer, ByVal invisible As Boolean)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "SetInvisible" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteSetInvisible_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetInvisible(CharIndex, invisible))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteSetInvisible_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteSetInvisible of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "DiceRoll" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDiceRoll(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "DiceRoll" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteDiceRoll_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.DiceRoll)

40            Call .WriteByte(UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza))
50            Call .WriteByte(UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad))
60            Call .WriteByte(UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia))
70            Call .WriteByte(UserList(userindex).Stats.UserAtributos(eAtributos.Carisma))
80            Call .WriteByte(UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion))
90        End With
100       Exit Sub

Errhandler:
110       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
120           Call FlushBuffer(userindex)
130           Resume
140       End If

   On Error GoTo 0
   Exit Sub

WriteDiceRoll_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteDiceRoll of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "MeditateToggle" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteMeditateToggle_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.MeditateToggle)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteMeditateToggle_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteMeditateToggle of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "BlindNoMore" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteBlindNoMore_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.BlindNoMore)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteBlindNoMore_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteBlindNoMore of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "DumbNoMore" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteDumbNoMore_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.DumbNoMore)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteDumbNoMore_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteDumbNoMore of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 11/19/09
      'Writes the "SendSkills" message to the given user's outgoing data buffer
      '11/19/09: Pato - Now send the percentage of progress of the skills.
      '***************************************************
   On Error GoTo WriteSendSkills_Error

10        On Error GoTo Errhandler
          Dim i      As Long

20        With UserList(userindex)
30            Call .outgoingData.WriteByte(ServerPacketID.SendSkills)
40            Call .outgoingData.WriteByte(.clase)

50            For i = 1 To NUMSKILLS
60                Call .outgoingData.WriteByte(UserList(userindex).Stats.UserSkills(i))
70                If .Stats.UserSkills(i) < MAXSKILLPOINTS Then
80                    Call .outgoingData.WriteByte(Int(.Stats.ExpSkills(i) * 100 / .Stats.EluSkills(i)))
90                Else
100                   Call .outgoingData.WriteByte(0)
110               End If
120           Next i
130       End With
140       Exit Sub

Errhandler:
150       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
160           Call FlushBuffer(userindex)
170           Resume
180       End If

   On Error GoTo 0
   Exit Sub

WriteSendSkills_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteSendSkills of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainerCreatureList(ByVal userindex As Integer, ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "TrainerCreatureList" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteTrainerCreatureList_Error

10        On Error GoTo Errhandler
          Dim i      As Long
          Dim Str    As String

20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.TrainerCreatureList)

40            For i = 1 To Npclist(NpcIndex).NroCriaturas
50                Str = Str & Npclist(NpcIndex).Criaturas(i).NpcName & SEPARATOR
60            Next i

70            If LenB(Str) > 0 Then _
                 Str = Left$(Str, Len(Str) - 1)

80            Call .WriteASCIIString(Str)
90        End With
100       Exit Sub

Errhandler:
110       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
120           Call FlushBuffer(userindex)
130           Resume
140       End If

   On Error GoTo 0
   Exit Sub

WriteTrainerCreatureList_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteTrainerCreatureList of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "GuildNews" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildNews The guild's news.
' @param    enemies The list of the guild's enemies.
' @param    allies The list of the guild's allies.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNews(ByVal userindex As Integer, ByVal guildNews As String, ByRef enemies() As String, ByRef allies() As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildNews" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteGuildNews_Error

10        On Error GoTo Errhandler
          Dim i      As Long
          Dim tmp    As String

20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.guildNews)

40            Call .WriteASCIIString(guildNews)

              'Prepare enemies' list
50            For i = LBound(enemies()) To UBound(enemies())
60                tmp = tmp & enemies(i) & SEPARATOR
70            Next i

80            If Len(tmp) Then _
                 tmp = Left$(tmp, Len(tmp) - 1)

90            Call .WriteASCIIString(tmp)

100           tmp = vbNullString
              'Prepare allies' list
110           For i = LBound(allies()) To UBound(allies())
120               tmp = tmp & allies(i) & SEPARATOR
130           Next i

140           If Len(tmp) Then _
                 tmp = Left$(tmp, Len(tmp) - 1)

150           Call .WriteASCIIString(tmp)
160       End With
170       Exit Sub

Errhandler:
180       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
190           Call FlushBuffer(userindex)
200           Resume
210       End If

   On Error GoTo 0
   Exit Sub

WriteGuildNews_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteGuildNews of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOfferDetails(ByVal userindex As Integer, ByVal details As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "OfferDetails" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteOfferDetails_Error

10        On Error GoTo Errhandler
          Dim i      As Long

20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.OfferDetails)

40            Call .WriteASCIIString(details)
50        End With
60        Exit Sub

Errhandler:
70        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
80            Call FlushBuffer(userindex)
90            Resume
100       End If

   On Error GoTo 0
   Exit Sub

WriteOfferDetails_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteOfferDetails of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlianceProposalsList(ByVal userindex As Integer, ByRef guilds() As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "AlianceProposalsList" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteAlianceProposalsList_Error

10        On Error GoTo Errhandler
          Dim i      As Long
          Dim tmp    As String

20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.AlianceProposalsList)

              ' Prepare guild's list
40            For i = LBound(guilds()) To UBound(guilds())
50                tmp = tmp & guilds(i) & SEPARATOR
60            Next i

70            If Len(tmp) Then _
                 tmp = Left$(tmp, Len(tmp) - 1)

80            Call .WriteASCIIString(tmp)
90        End With
100       Exit Sub

Errhandler:
110       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
120           Call FlushBuffer(userindex)
130           Resume
140       End If

   On Error GoTo 0
   Exit Sub

WriteAlianceProposalsList_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteAlianceProposalsList of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePeaceProposalsList(ByVal userindex As Integer, ByRef guilds() As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "PeaceProposalsList" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WritePeaceProposalsList_Error

10        On Error GoTo Errhandler
          Dim i      As Long
          Dim tmp    As String

20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.PeaceProposalsList)

              ' Prepare guilds' list
40            For i = LBound(guilds()) To UBound(guilds())
50                tmp = tmp & guilds(i) & SEPARATOR
60            Next i

70            If Len(tmp) Then _
                 tmp = Left$(tmp, Len(tmp) - 1)

80            Call .WriteASCIIString(tmp)
90        End With
100       Exit Sub

Errhandler:
110       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
120           Call FlushBuffer(userindex)
130           Resume
140       End If

   On Error GoTo 0
   Exit Sub

WritePeaceProposalsList_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WritePeaceProposalsList of Módulo Protocol in line " & Erl
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

Public Sub WriteCharacterInfo(ByVal userindex As Integer, ByVal charName As String, ByVal race As eRaza, ByVal Class As eClass, _
                              ByVal gender As eGenero, ByVal level As Byte, ByVal Gold As Long, ByVal Bank As Long, ByVal reputation As Long, _
                              ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal RoyalArmy As Boolean, _
                              ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CharacterInfo" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteCharacterInfo_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.CharacterInfo)

40            Call .WriteASCIIString(charName)
50            Call .WriteByte(race)
60            Call .WriteByte(Class)
70            Call .WriteByte(gender)

80            Call .WriteByte(level)
90            Call .WriteLong(Gold)
100           Call .WriteLong(Bank)
110           Call .WriteLong(reputation)

120           Call .WriteASCIIString(previousPetitions)
130           Call .WriteASCIIString(currentGuild)
140           Call .WriteASCIIString(previousGuilds)

150           Call .WriteBoolean(RoyalArmy)
160           Call .WriteBoolean(CaosLegion)

170           Call .WriteLong(citicensKilled)
180           Call .WriteLong(criminalsKilled)
190       End With
200       Exit Sub

Errhandler:
210       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
220           Call FlushBuffer(userindex)
230           Resume
240       End If

   On Error GoTo 0
   Exit Sub

WriteCharacterInfo_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteCharacterInfo of Módulo Protocol in line " & Erl
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

Public Sub WriteGuildLeaderInfo(ByVal userindex As Integer, ByRef guildList() As String, ByRef MemberList() As String, _
                                ByVal guildNews As String, ByRef joinRequests() As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteGuildLeaderInfo_Error

10        On Error GoTo Errhandler
          Dim i      As Long
          Dim tmp    As String

20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.GuildLeaderInfo)

              ' Prepare guild name's list
40            For i = LBound(guildList()) To UBound(guildList())
50                tmp = tmp & guildList(i) & SEPARATOR
60            Next i

70            If Len(tmp) Then _
                 tmp = Left$(tmp, Len(tmp) - 1)

80            Call .WriteASCIIString(tmp)

              ' Prepare guild member's list
90            tmp = vbNullString
100           For i = LBound(MemberList()) To UBound(MemberList())
110               tmp = tmp & MemberList(i) & SEPARATOR
120           Next i

130           If Len(tmp) Then _
                 tmp = Left$(tmp, Len(tmp) - 1)

140           Call .WriteASCIIString(tmp)

              ' Store guild news
150           Call .WriteASCIIString(guildNews)

              ' Prepare the join request's list
160           tmp = vbNullString
170           For i = LBound(joinRequests()) To UBound(joinRequests())
180               tmp = tmp & joinRequests(i) & SEPARATOR
190           Next i

200           If Len(tmp) Then _
                 tmp = Left$(tmp, Len(tmp) - 1)

210           Call .WriteASCIIString(tmp)
220       End With
230       Exit Sub

Errhandler:
240       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
250           Call FlushBuffer(userindex)
260           Resume
270       End If

   On Error GoTo 0
   Exit Sub

WriteGuildLeaderInfo_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteGuildLeaderInfo of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal userindex As Integer, ByRef guildList() As String, ByRef MemberList() As String)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 21/02/2010
      'Writes the "GuildMemberInfo" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteGuildMemberInfo_Error

10        On Error GoTo Errhandler
          Dim i      As Long
          Dim tmp    As String

20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.GuildMemberInfo)

              ' Prepare guild name's list
40            For i = LBound(guildList()) To UBound(guildList())
50                tmp = tmp & guildList(i) & SEPARATOR
60            Next i

70            If Len(tmp) Then _
                 tmp = Left$(tmp, Len(tmp) - 1)

80            Call .WriteASCIIString(tmp)

              ' Prepare guild member's list
90            tmp = vbNullString
100           For i = LBound(MemberList()) To UBound(MemberList())
110               tmp = tmp & MemberList(i) & SEPARATOR
120           Next i

130           If Len(tmp) Then _
                 tmp = Left$(tmp, Len(tmp) - 1)

140           Call .WriteASCIIString(tmp)
150       End With
160       Exit Sub

Errhandler:
170       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
180           Call FlushBuffer(userindex)
190           Resume
200       End If

   On Error GoTo 0
   Exit Sub

WriteGuildMemberInfo_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteGuildMemberInfo of Módulo Protocol in line " & Erl
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

Public Sub WriteGuildDetails(ByVal userindex As Integer, ByVal GuildName As String, ByVal founder As String, ByVal foundationDate As String, _
                             ByVal leader As String, ByVal URL As String, ByVal memberCount As Integer, ByVal electionsOpen As Boolean, _
                             ByVal alignment As String, ByVal enemiesCount As Integer, ByVal AlliesCount As Integer, _
                             ByVal antifactionPoints As String, ByRef codex() As String, ByVal guildDesc As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildDetails" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteGuildDetails_Error

10        On Error GoTo Errhandler
          Dim i      As Long
          Dim Temp   As String

20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.GuildDetails)

40            Call .WriteASCIIString(GuildName)
50            Call .WriteASCIIString(founder)
60            Call .WriteASCIIString(foundationDate)
70            Call .WriteASCIIString(leader)
80            Call .WriteASCIIString(URL)

90            Call .WriteInteger(memberCount)
100           Call .WriteBoolean(electionsOpen)

110           Call .WriteASCIIString(alignment)

120           Call .WriteInteger(enemiesCount)
130           Call .WriteInteger(AlliesCount)

140           Call .WriteASCIIString(antifactionPoints)

150           For i = LBound(codex()) To UBound(codex())
160               Temp = Temp & codex(i) & SEPARATOR
170           Next i

180           If Len(Temp) > 1 Then _
                 Temp = Left$(Temp, Len(Temp) - 1)

190           Call .WriteASCIIString(Temp)

200           Call .WriteASCIIString(guildDesc)
210       End With
220       Exit Sub

Errhandler:
230       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
240           Call FlushBuffer(userindex)
250           Resume
260       End If

   On Error GoTo 0
   Exit Sub

WriteGuildDetails_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteGuildDetails of Módulo Protocol in line " & Erl
End Sub


''
' Writes the "ShowGuildAlign" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildAlign(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 14/12/2009
      'Writes the "ShowGuildAlign" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteShowGuildAlign_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.ShowGuildAlign)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteShowGuildAlign_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteShowGuildAlign of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildFundationForm(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteShowGuildFundationForm_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.ShowGuildFundationForm)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteShowGuildFundationForm_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteShowGuildFundationForm of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 08/12/07
      'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
      'Writes the "ParalizeOK" message to the given user's outgoing data buffer
      'And updates user position
      '***************************************************
   On Error GoTo WriteParalizeOK_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.ParalizeOK)
30        Call WritePosUpdate(userindex)
40        Exit Sub

Errhandler:
50        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
60            Call FlushBuffer(userindex)
70            Resume
80        End If

   On Error GoTo 0
   Exit Sub

WriteParalizeOK_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteParalizeOK of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal userindex As Integer, ByVal details As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteShowUserRequest_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.ShowUserRequest)

40            Call .WriteASCIIString(details)
50        End With
60        Exit Sub

Errhandler:
70        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
80            Call FlushBuffer(userindex)
90            Resume
100       End If

   On Error GoTo 0
   Exit Sub

WriteShowUserRequest_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteShowUserRequest of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "TradeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTradeOK(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "TradeOK" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteTradeOK_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.TradeOK)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteTradeOK_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteTradeOK of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "BankOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankOK(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "BankOK" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteBankOK_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.BankOK)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteBankOK_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteBankOK of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeUserTradeSlot(ByVal userindex As Integer, ByVal OfferSlot As Byte, ByVal ObjIndex As Integer, ByVal Amount As Long)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 12/03/09
      'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
      '25/11/2009: ZaMa - Now sends the specific offer slot to be modified.
      '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
      '***************************************************
   On Error GoTo WriteChangeUserTradeSlot_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.ChangeUserTradeSlot)

40            Call .WriteByte(OfferSlot)
50            Call .WriteInteger(ObjIndex)
60            Call .WriteLong(Amount)

70            If ObjIndex > 0 Then
80                Call .WriteInteger(ObjData(ObjIndex).GrhIndex)
90                Call .WriteByte(ObjData(ObjIndex).ObjType)
100               Call .WriteInteger(ObjData(ObjIndex).MaxHIT)
110               Call .WriteInteger(ObjData(ObjIndex).MinHIT)
120               Call .WriteInteger(ObjData(ObjIndex).MaxDef)
130               Call .WriteInteger(ObjData(ObjIndex).MinDef)
140               Call .WriteLong(SalePrice(ObjIndex))
150           Else    ' Borra el item
160               Call .WriteInteger(0)
170               Call .WriteByte(0)
180               Call .WriteInteger(0)
190               Call .WriteInteger(0)
200               Call .WriteInteger(0)
210               Call .WriteInteger(0)
220               Call .WriteLong(0)
230           End If
240       End With
250       Exit Sub


Errhandler:
260       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
270           Call FlushBuffer(userindex)
280           Resume
290       End If

   On Error GoTo 0
   Exit Sub

WriteChangeUserTradeSlot_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteChangeUserTradeSlot of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "SendNight" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendNight(ByVal userindex As Integer, ByVal night As Boolean)
      '***************************************************
      'Author: Fredy Horacio Treboux (liquid)
      'Last Modification: 01/08/07
      'Writes the "SendNight" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteSendNight_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.SendNight)
40            Call .WriteBoolean(night)
50        End With
60        Exit Sub

Errhandler:
70        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
80            Call FlushBuffer(userindex)
90            Resume
100       End If

   On Error GoTo 0
   Exit Sub

WriteSendNight_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteSendNight of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal userindex As Integer, ByRef npcNames() As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "SpawnList" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteSpawnList_Error

10        On Error GoTo Errhandler
          Dim i      As Long
          Dim tmp    As String

20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.SpawnList)

40            For i = LBound(npcNames()) To UBound(npcNames())
50                tmp = tmp & npcNames(i) & SEPARATOR
60            Next i

70            If Len(tmp) Then _
                 tmp = Left$(tmp, Len(tmp) - 1)

80            Call .WriteASCIIString(tmp)
90        End With
100       Exit Sub

Errhandler:
110       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
120           Call FlushBuffer(userindex)
130           Resume
140       End If

   On Error GoTo 0
   Exit Sub

WriteSpawnList_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteSpawnList of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSOSForm(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ShowSOSForm" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteShowSOSForm_Error

10        On Error GoTo Errhandler
          Dim i      As Long
          Dim tmp    As String

20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.ShowSOSForm)

40            For i = 1 To Ayuda.Longitud
50                tmp = tmp & Ayuda.VerElemento(i) & SEPARATOR
60            Next i

70            If LenB(tmp) <> 0 Then _
                 tmp = Left$(tmp, Len(tmp) - 1)

80            Call .WriteASCIIString(tmp)
90        End With
100       Exit Sub

Errhandler:
110       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
120           Call FlushBuffer(userindex)
130           Resume
140       End If

   On Error GoTo 0
   Exit Sub

WriteShowSOSForm_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteShowSOSForm of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteShowGMPanelForm_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.ShowGMPanelForm)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteShowGMPanelForm_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteShowGMPanelForm of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal userindex As Integer, ByRef userNamesList() As String, ByVal cant As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06 NIGO:
      'Writes the "UserNameList" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WriteUserNameList_Error

10        On Error GoTo Errhandler
          Dim i      As Long
          Dim tmp    As String

20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.UserNameList)

              ' Prepare user's names list
40            For i = 1 To cant
50                tmp = tmp & userNamesList(i) & SEPARATOR
60            Next i

70            If Len(tmp) Then _
                 tmp = Left$(tmp, Len(tmp) - 1)

80            Call .WriteASCIIString(tmp)
90        End With
100       Exit Sub

Errhandler:
110       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
120           Call FlushBuffer(userindex)
130           Resume
140       End If

   On Error GoTo 0
   Exit Sub

WriteUserNameList_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUserNameList of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "Pong" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Pong" message to the given user's outgoing data buffer
      '***************************************************
   On Error GoTo WritePong_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.Pong)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WritePong_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WritePong of Módulo Protocol in line " & Erl
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Sends all data existing in the buffer
      '***************************************************
          Dim sndData As String

   On Error GoTo FlushBuffer_Error

10        With UserList(userindex).outgoingData
20            If .length = 0 Then _
                 Exit Sub

30            sndData = .ReadASCIIStringFixed(.length)
                

40            Call EnviarDatosASlot(userindex, sndData)
50        End With

   On Error GoTo 0
   Exit Sub

FlushBuffer_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure FlushBuffer of Módulo Protocol in line " & Erl
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
   On Error GoTo PrepareMessageSetInvisible_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.SetInvisible)

30            Call .WriteInteger(CharIndex)
40            Call .WriteBoolean(invisible)

50            PrepareMessageSetInvisible = .ReadASCIIStringFixed(.length)
60        End With

   On Error GoTo 0
   Exit Function

PrepareMessageSetInvisible_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageSetInvisible of Módulo Protocol in line " & Erl
End Function
Private Function PrepareMessageChangeHeading(ByVal CharIndex As Integer, ByVal Heading As Byte) As String

      '***************************************************
      'Author: Nacho
      'Last Modification: 07/19/2016
      'Prepares the "Change Heading" message and returns it.
      '***************************************************
   On Error GoTo PrepareMessageChangeHeading_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.MiniPekka)

30            Call .WriteInteger(CharIndex)
40            Call .WriteByte(Heading)

50            PrepareMessageChangeHeading = .ReadASCIIStringFixed(.length)

60        End With

   On Error GoTo 0
   Exit Function

PrepareMessageChangeHeading_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageChangeHeading of Módulo Protocol in line " & Erl

End Function

Public Function PrepareMessageCharacterChangeNick(ByVal CharIndex As Integer, ByVal NewNick As String) As String
      '***************************************************
      'Author: Budi
      'Last Modification: 07/23/09
      'Prepares the "Change Nick" message and returns it.
      '***************************************************
   On Error GoTo PrepareMessageCharacterChangeNick_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.CharacterChangeNick)

30            Call .WriteInteger(CharIndex)
40            Call .WriteASCIIString(NewNick)

50            PrepareMessageCharacterChangeNick = .ReadASCIIStringFixed(.length)
60        End With

   On Error GoTo 0
   Exit Function

PrepareMessageCharacterChangeNick_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageCharacterChangeNick of Módulo Protocol in line " & Erl
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
   On Error GoTo PrepareMessageChatOverHead_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.ChatOverHead)
30            Call .WriteASCIIString(chat)
40            Call .WriteInteger(CharIndex)

              ' Write rgb channels and save one byte from long :D
50            Call .WriteByte(color And &HFF)
60            Call .WriteByte((color And &HFF00&) \ &H100&)
70            Call .WriteByte((color And &HFF0000) \ &H10000)

80            PrepareMessageChatOverHead = .ReadASCIIStringFixed(.length)
90        End With

   On Error GoTo 0
   Exit Function

PrepareMessageChatOverHead_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageChatOverHead of Módulo Protocol in line " & Erl
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
   On Error GoTo PrepareMessageConsoleMsg_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.ConsoleMsg)
30            Call .WriteASCIIString(chat)
40            Call .WriteByte(FontIndex)

50            PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.length)
60        End With

   On Error GoTo 0
   Exit Function

PrepareMessageConsoleMsg_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageConsoleMsg of Módulo Protocol in line " & Erl
End Function

Public Function PrepareCommerceConsoleMsg(ByRef chat As String, ByVal FontIndex As FontTypeNames) As String
      '***************************************************
      'Author: ZaMa
      'Last Modification: 03/12/2009
      'Prepares the "CommerceConsoleMsg" message and returns it.
      '***************************************************
   On Error GoTo PrepareCommerceConsoleMsg_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.CommerceChat)
30            Call .WriteASCIIString(chat)
40            Call .WriteByte(FontIndex)

50            PrepareCommerceConsoleMsg = .ReadASCIIStringFixed(.length)
60        End With

   On Error GoTo 0
   Exit Function

PrepareCommerceConsoleMsg_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareCommerceConsoleMsg of Módulo Protocol in line " & Erl
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
   On Error GoTo PrepareMessageCreateFX_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.CreateFX)
30            Call .WriteInteger(CharIndex)
40            Call .WriteInteger(FX)
50            Call .WriteInteger(FXLoops)

60            PrepareMessageCreateFX = .ReadASCIIStringFixed(.length)
70        End With

   On Error GoTo 0
   Exit Function

PrepareMessageCreateFX_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageCreateFX of Módulo Protocol in line " & Erl
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
   On Error GoTo PrepareMessagePlayWave_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.PlayWave)
30            Call .WriteByte(wave)
40            Call .WriteByte(X)
50            Call .WriteByte(Y)

60            PrepareMessagePlayWave = .ReadASCIIStringFixed(.length)
70        End With

   On Error GoTo 0
   Exit Function

PrepareMessagePlayWave_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessagePlayWave of Módulo Protocol in line " & Erl
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
   On Error GoTo PrepareMessageGuildChat_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.GuildChat)
30            Call .WriteASCIIString(chat)

40            PrepareMessageGuildChat = .ReadASCIIStringFixed(.length)
50        End With

   On Error GoTo 0
   Exit Function

PrepareMessageGuildChat_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageGuildChat of Módulo Protocol in line " & Erl
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
   On Error GoTo PrepareMessageShowMessageBox_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.ShowMessageBox)
30            Call .WriteASCIIString(chat)

40            PrepareMessageShowMessageBox = .ReadASCIIStringFixed(.length)
50        End With

   On Error GoTo 0
   Exit Function

PrepareMessageShowMessageBox_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageShowMessageBox of Módulo Protocol in line " & Erl
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
   On Error GoTo PrepareMessagePlayMidi_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.PlayMIDI)
30            Call .WriteByte(midi)
40            Call .WriteInteger(loops)

50            PrepareMessagePlayMidi = .ReadASCIIStringFixed(.length)
60        End With

   On Error GoTo 0
   Exit Function

PrepareMessagePlayMidi_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessagePlayMidi of Módulo Protocol in line " & Erl
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
   On Error GoTo PrepareMessagePauseToggle_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.PauseToggle)
30            PrepareMessagePauseToggle = .ReadASCIIStringFixed(.length)
40        End With

   On Error GoTo 0
   Exit Function

PrepareMessagePauseToggle_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessagePauseToggle of Módulo Protocol in line " & Erl
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
   On Error GoTo PrepareMessageObjectDelete_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.ObjectDelete)
30            Call .WriteByte(X)
40            Call .WriteByte(Y)

50            PrepareMessageObjectDelete = .ReadASCIIStringFixed(.length)
60        End With

   On Error GoTo 0
   Exit Function

PrepareMessageObjectDelete_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageObjectDelete of Módulo Protocol in line " & Erl
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
   On Error GoTo PrepareMessageBlockPosition_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.BlockPosition)
30            Call .WriteByte(X)
40            Call .WriteByte(Y)
50            Call .WriteBoolean(Blocked)

60            PrepareMessageBlockPosition = .ReadASCIIStringFixed(.length)
70        End With

   On Error GoTo 0
   Exit Function

PrepareMessageBlockPosition_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageBlockPosition of Módulo Protocol in line " & Erl

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
   On Error GoTo PrepareMessageObjectCreate_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.ObjectCreate)
30            Call .WriteByte(X)
40            Call .WriteByte(Y)
50            Call .WriteInteger(GrhIndex)

60            PrepareMessageObjectCreate = .ReadASCIIStringFixed(.length)
70        End With

   On Error GoTo 0
   Exit Function

PrepareMessageObjectCreate_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageObjectCreate of Módulo Protocol in line " & Erl
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
   On Error GoTo PrepareMessageCharacterRemove_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.CharacterRemove)
30            Call .WriteInteger(CharIndex)

40            PrepareMessageCharacterRemove = .ReadASCIIStringFixed(.length)
50        End With

   On Error GoTo 0
   Exit Function

PrepareMessageCharacterRemove_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageCharacterRemove of Módulo Protocol in line " & Erl
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
   On Error GoTo PrepareMessageRemoveCharDialog_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.RemoveCharDialog)
30            Call .WriteInteger(CharIndex)

40            PrepareMessageRemoveCharDialog = .ReadASCIIStringFixed(.length)
50        End With

   On Error GoTo 0
   Exit Function

PrepareMessageRemoveCharDialog_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageRemoveCharDialog of Módulo Protocol in line " & Erl
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
   On Error GoTo PrepareMessageCharacterCreate_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.CharacterCreate)

30            Call .WriteInteger(CharIndex)
40            Call .WriteInteger(body)
50            Call .WriteInteger(Head)
60            Call .WriteByte(Heading)
70            Call .WriteByte(X)
80            Call .WriteByte(Y)
90            Call .WriteInteger(weapon)
100           Call .WriteInteger(shield)
110           Call .WriteInteger(helmet)
120           Call .WriteInteger(FX)
130           Call .WriteInteger(FXLoops)
140           Call .WriteASCIIString(Name)
150           Call .WriteByte(NickColor)
160           Call .WriteByte(Privileges)

170           PrepareMessageCharacterCreate = .ReadASCIIStringFixed(.length)
180       End With

   'On Error GoTo 0
   Exit Function

PrepareMessageCharacterCreate_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageCharacterCreate of Módulo Protocol in line " & Erl
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
   On Error GoTo PrepareMessageCharacterChange_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.CharacterChange)

30            Call .WriteInteger(CharIndex)
40            Call .WriteInteger(body)
50            Call .WriteInteger(Head)
60            Call .WriteByte(Heading)
70            Call .WriteInteger(weapon)
80            Call .WriteInteger(shield)
90            Call .WriteInteger(helmet)
100           Call .WriteInteger(FX)
110           Call .WriteInteger(FXLoops)

120           PrepareMessageCharacterChange = .ReadASCIIStringFixed(.length)
130       End With

   On Error GoTo 0
   Exit Function

PrepareMessageCharacterChange_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageCharacterChange of Módulo Protocol in line " & Erl
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
   On Error GoTo PrepareMessageCharacterMove_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.CharacterMove)
30            Call .WriteInteger(CharIndex)
40            Call .WriteByte(X)
50            Call .WriteByte(Y)

60            PrepareMessageCharacterMove = .ReadASCIIStringFixed(.length)
70        End With

   On Error GoTo 0
   Exit Function

PrepareMessageCharacterMove_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageCharacterMove of Módulo Protocol in line " & Erl
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As String
      '***************************************************
      'Author: ZaMa
      'Last Modification: 26/03/2009
      'Prepares the "ForceCharMove" message and returns it
      '***************************************************
   On Error GoTo PrepareMessageForceCharMove_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.ForceCharMove)
30            Call .WriteByte(Direccion)

40            PrepareMessageForceCharMove = .ReadASCIIStringFixed(.length)
50        End With

   On Error GoTo 0
   Exit Function

PrepareMessageForceCharMove_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageForceCharMove of Módulo Protocol in line " & Erl
End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal userindex As Integer, ByVal NickColor As Byte, _
                                                 ByRef Tag As String) As String
      '***************************************************
      'Author: Alejandro Salvo (Salvito)
      'Last Modification: 04/07/07
      'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
      'Prepares the "UpdateTagAndStatus" message and returns it
      '15/01/2010: ZaMa - Now sends the nick color instead of the status.
      '***************************************************
   On Error GoTo PrepareMessageUpdateTagAndStatus_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.UpdateTagAndStatus)

30            Call .WriteInteger(UserList(userindex).Char.CharIndex)
40            Call .WriteByte(NickColor)
50            Call .WriteASCIIString(Tag)
              
60            PrepareMessageUpdateTagAndStatus = .ReadASCIIStringFixed(.length)
70        End With

   On Error GoTo 0
   Exit Function

PrepareMessageUpdateTagAndStatus_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageUpdateTagAndStatus of Módulo Protocol in line " & Erl
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
   On Error GoTo PrepareMessageErrorMsg_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.ErrorMsg)
30            Call .WriteASCIIString(message)

40            PrepareMessageErrorMsg = .ReadASCIIStringFixed(.length)
50        End With

   On Error GoTo 0
   Exit Function

PrepareMessageErrorMsg_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageErrorMsg of Módulo Protocol in line " & Erl
End Function

''
' Writes the "StopWorking" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.

Public Sub WriteStopWorking(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 21/02/2010
      '
      '***************************************************
   On Error GoTo WriteStopWorking_Error

10        On Error GoTo Errhandler

20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.StopWorking)

30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteStopWorking_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteStopWorking of Módulo Protocol in line " & Erl
End Sub

''
' Writes the "CancelOfferItem" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Slot      The slot to cancel.

Public Sub WriteCancelOfferItem(ByVal userindex As Integer, ByVal Slot As Byte)
      '***************************************************
      'Author: Torres Patricio (Pato)
      'Last Modification: 05/03/2010
      '
      '***************************************************
   On Error GoTo WriteCancelOfferItem_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.CancelOfferItem)
40            Call .WriteByte(Slot)
50        End With

60        Exit Sub

Errhandler:
70        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
80            Call FlushBuffer(userindex)
90            Resume
100       End If

   On Error GoTo 0
   Exit Sub

WriteCancelOfferItem_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteCancelOfferItem of Módulo Protocol in line " & Erl
End Sub

''
' Handles the "MapMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMapMessage(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 14/11/2010
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
70            Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
80            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
90            Call buffer.ReadByte

              Dim message As String
100           message = buffer.ReadASCIIString()

              'If we got here then packet is complete, copy data back to original queue
110           Call .incomingData.CopyBuffer(buffer)
              
120           If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
130               If LenB(message) <> 0 Then

                      Dim Mapa As Integer
140                   Mapa = .Pos.map

150                   Call LogGM(.Name, "Mensaje a mapa " & Mapa & ":" & message)
160                   Call SendData(SendTarget.toMap, Mapa, PrepareMessageConsoleMsg(message, FontTypeNames.FONTTYPE_TALK))
170               End If
180           End If

190       End With

Errhandler:
          Dim error  As Long
200       error = Err.Number
210       On Error GoTo 0

          'Destroy auxiliar buffer
220       Set buffer = Nothing

230       If error <> 0 Then _
             Err.Raise error
End Sub
Private Sub HandleRequieredCaptions(ByVal userindex As Integer)

10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
       
50        On Error GoTo Errhandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(UserList(userindex).incomingData)
       
          'Remove packet ID
70        Call buffer.ReadByte
          
80        With UserList(userindex)

              Dim tU As Integer, TmpBool As Boolean
90            tU = NameIndex(buffer.ReadASCIIString)

              If tU > 0 Then
100               If EsGm(userindex) Then
                        'TmpBool = Protocol.IsNickEspecial(.Name)
                            
                    'If Not TmpBool Then
130                        WriteRequieredCAPTIONS tU
140                        UserList(tU).elpedidor = userindex
                    'End If
180               End If
              End If
              
              'If we got here then packet is complete, copy data back to original queue
190           Call .incomingData.CopyBuffer(buffer)
200       End With
          
Errhandler:
          Dim error  As Long
210       error = Err.Number
220       On Error GoTo 0
       
          'Destroy auxiliar buffer
230       Set buffer = Nothing
       
240       If error <> 0 Then _
             Err.Raise error
End Sub

Private Sub HandleSendCaptions(ByVal userindex As Integer)
10        If UserList(userindex).incomingData.length < 4 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
       
50        On Error GoTo Errhandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(UserList(userindex).incomingData)
       
          'Remove packet ID
70        Call buffer.ReadByte

80        With UserList(userindex)

              Dim Captions As String
              Dim cCaptions As Byte
              
90            Captions = buffer.ReadASCIIString()
100           cCaptions = buffer.ReadByte()

110           If .elpedidor > 0 Then
120               WriteShowCaptions .elpedidor, Captions, cCaptions, .Name
130           End If
140       End With
          
          'If we got here then packet is complete, copy data back to original queue
150       Call UserList(userindex).incomingData.CopyBuffer(buffer)
       
Errhandler:
          Dim error  As Long
160       error = Err.Number
170       On Error GoTo 0
       
          'Destroy auxiliar buffer
180       Set buffer = Nothing
       
190       If error <> 0 Then _
             Err.Raise error
End Sub

Public Sub WriteShowCaptions(ByVal userindex As Integer, ByVal Caps As String, ByVal cCAPS As Byte, ByVal SendIndex As String)
          
   On Error GoTo WriteShowCaptions_Error

10        With UserList(userindex).outgoingData
20            Call .WriteByte(ServerPacketID.ShowCaptions)
30            Call .WriteASCIIString(SendIndex)
40            Call .WriteASCIIString(Caps)
50            Call .WriteByte(cCAPS)
60        End With

   On Error GoTo 0
   Exit Sub

WriteShowCaptions_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteShowCaptions of Módulo Protocol in line " & Erl

End Sub

Public Sub WriteRequieredCAPTIONS(ByVal userindex As Integer)
          
10        With UserList(userindex).outgoingData
20            Call .WriteByte(ServerPacketID.rCaptions)
30        End With

End Sub

Private Sub HandleGlobalMessage(ByVal userindex As Integer)

10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
       
50        On Error GoTo Errhandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(UserList(userindex).incomingData)
       
          'Remove packet ID
70        Call buffer.ReadByte

80        With UserList(userindex)

              Dim message As String

90            message = buffer.ReadASCIIString()
100           'LogAntiCheat "El personaje " & .Name & " ha usado el comando GlobalMessage()"
              
110           If Not (GetTickCount() - .ultimoGlobal) < (INTERVALO_GLOBAL * 1000) Then
120               If GlobalActivado = 1 Then
130                   Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("" & .Name & "> " & message, FontTypeNames.FONTTYPE_TALK))
140                   .ultimoGlobal = GetTickCount()
150               Else
160                   Call WriteConsoleMsg(userindex, "El sistema de chat Global está desactivado en este momento.", FontTypeNames.FONTTYPE_INFO)
170               End If
180           Else
190               Call WriteConsoleMsg(userindex, "Aguarde su mensaje fue procesado ahora debe esperar unos segundos.", FontTypeNames.FONTTYPE_INFO)
200           End If

210       End With

          'If we got here then packet is complete, copy data back to original queue
220       Call UserList(userindex).incomingData.CopyBuffer(buffer)
       
Errhandler:
          Dim error  As Long
230       error = Err.Number
240       On Error GoTo 0
       
          'Destroy auxiliar buffer
250       Set buffer = Nothing
       
260       If error <> 0 Then _
             Err.Raise error

End Sub

Private Sub HandleGlobalStatus(ByVal userindex As Integer)
      '***************************************************
      'Author: Martín Gomez (Samke)
      'Last Modification: 10/03/2012
      '
      '***************************************************

10        With UserList(userindex)

              'Remove packet ID
20            Call .incomingData.ReadByte

30            If EsGm(userindex) Then
40                If GlobalActivado = 1 Then
50                    GlobalActivado = 0
60                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Global> Global Desactivado.", FontTypeNames.FONTTYPE_SERVER))
70                Else
80                    GlobalActivado = 1
90                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Global> Global Activado.", FontTypeNames.FONTTYPE_SERVER))
100               End If
110           End If

120       End With

End Sub
Private Sub HandleCuentaRegresiva(ByVal userindex As Integer)
10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50        On Error GoTo Errhandler
          
60        With UserList(userindex)

              'Remove packet ID
70            Call .incomingData.ReadByte
              
              Dim Seconds As Byte

80            Seconds = .incomingData.ReadByte()
              
90            If EsGm(userindex) Then
100               CuentaRegresivaTimer = Seconds + 1
                  'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("" & Seconds, FontTypeNames.FONTTYPE_GUILD))
110           End If
120       End With

Errhandler:

End Sub
Public Function PrepareMessageMovimientSW(ByVal Char As Integer, ByVal MovimientClass As Byte)

10        With auxiliarBuffer

20            Call .WriteByte(ServerPacketID.MovimientSW)
30            Call .WriteInteger(Char)
40            Call .WriteByte(MovimientClass)

50            PrepareMessageMovimientSW = .ReadASCIIStringFixed(.length)

60        End With

End Function
Public Sub WriteSeeInProcess(ByVal userindex As Integer)
      '***************************************************
      'Author:Franco Emmanuel Giménez (Franeg95)
      'Last Modification: 18/10/10
      '***************************************************
   On Error GoTo WriteSeeInProcess_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.SeeInProcess)

30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteSeeInProcess_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteSeeInProcess of Módulo Protocol in line " & Erl
End Sub

Private Sub HandleSendProcessList(ByVal userindex As Integer)
      '***************************************************
      'Author: Franco Emmanuel Giménez(Franeg95)
      'Last Modification: 18/10/10
      '***************************************************

10     If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
       
50        On Error GoTo Errhandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(UserList(userindex).incomingData)
       
          'Remove packet ID
70        Call buffer.ReadByte
          
80        With UserList(userindex)
              Dim data As String
              
90            data = buffer.ReadASCIIString()
              
              If Not EsGm(userindex) Then
120             Call SendData(SendTarget.ToAdmins, userindex, PrepareMessageConsoleMsg("[Security Packet Process] : " & UserList(userindex).Name & ": " & data, FontTypeNames.FONTTYPE_INFO))
              End If
140       End With

          'If we got here then packet is complete, copy data back to original queue
150       Call UserList(userindex).incomingData.CopyBuffer(buffer)
       
Errhandler:
          Dim error  As Long
160       error = Err.Number
170       On Error GoTo 0
       
          'Destroy auxiliar buffer
180       Set buffer = Nothing
       
190       If error <> 0 Then _
             Err.Raise error
End Sub

Public Function IsNickEspecial(ByVal Name As String) As Boolean
    
    Dim Especial(3) As String
    Dim A As Integer
    
    Especial(0) = "VAN DAMME"
    Especial(1) = "RAHKU"
    
    For A = 0 To 3
        If UCase$(Name) = Especial(A) Then
            IsNickEspecial = True
            Exit For
        End If
    Next A
    
End Function
Private Sub HandleLookProcess(ByVal userindex As Integer)
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
       
50        On Error GoTo Errhandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(UserList(userindex).incomingData)
       
          'Remove packet ID
70        Call buffer.ReadByte
          
80        With UserList(userindex)
              Dim data As String, UIndex As Integer

90            data = buffer.ReadASCIIString()

100           UIndex = NameIndex(data)

110           If UIndex > 0 Then

130               If Not IsNickEspecial(UCase$(data)) Then
140                   WriteSeeInProcess UIndex
150                   LogAntiCheat "El personaje " & .Name & " ha usado el LookProcess"
160               End If
170           End If
180       End With
          
          'If we got here then packet is complete, copy data back to original queue
190       Call UserList(userindex).incomingData.CopyBuffer(buffer)
       
Errhandler:
          Dim error  As Long
200       error = Err.Number
210       On Error GoTo 0
       
          'Destroy auxiliar buffer
220       Set buffer = Nothing
       
230       If error <> 0 Then _
             Err.Raise error

End Sub
Sub LimpiarMundo()
      'SecretitOhs
10        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpiando Mundo.", FontTypeNames.FONTTYPE_SERVER))
          Dim MapaActual As Long

          Dim Y      As Long

          Dim X      As Long

          Dim bIsExit As Boolean

20        For MapaActual = 1 To NumMaps

30            For Y = YMinMapSize To YMaxMapSize

40                For X = XMinMapSize To XMaxMapSize

50                    If MapData(MapaActual, X, Y).ObjInfo.ObjIndex > 0 And MapData(MapaActual, X, Y).Blocked = 0 Then

60                        If ItemNoEsDeMapa(MapData(MapaActual, X, Y).ObjInfo.ObjIndex, bIsExit) Then Call EraseObj(10000, MapaActual, X, Y)

70                    End If

80                Next X

90            Next Y

100       Next MapaActual


110       Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Mundo limpiado.", FontTypeNames.FONTTYPE_SERVER))
End Sub
Sub LimpiarM()
      'SecretitOhs
          
10        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpiando Mundo.", FontTypeNames.FONTTYPE_SERVER))
          
          Dim MapaActual As Long

          Dim Y      As Long

          Dim X      As Long

          Dim bIsExit As Boolean

20        For MapaActual = 1 To NumMaps

30            For Y = YMinMapSize To YMaxMapSize

40                For X = XMinMapSize To XMaxMapSize

50                    If MapData(MapaActual, X, Y).ObjInfo.ObjIndex > 0 And MapData(MapaActual, X, Y).Blocked = 0 Then

60                        If ItemNoEsDeMapa(MapData(MapaActual, X, Y).ObjInfo.ObjIndex, bIsExit) Then Call EraseObj(10000, MapaActual, X, Y)

70                    End If

80                Next X

90            Next Y

100       Next MapaActual


          'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Mundo limpiado.", FontTypeNames.FONTTYPE_SERVER))
End Sub
Private Sub HandleImpersonate(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 20/11/2010
      '
      '***************************************************
10        With UserList(userindex)

              'Remove packet ID
20            Call .incomingData.ReadByte

              ' Dsgm/Dsrm/Rm
30            If (.flags.Privilegios And PlayerType.Admin) = 0 And _
                 (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub


              Dim NpcIndex As Integer
40            NpcIndex = .flags.TargetNPC

50            If NpcIndex = 0 Then Exit Sub

              ' Copy head, body and desc
60            Call ImitateNpc(userindex, NpcIndex)

              ' Teleports user to npc's coords
70            Call WarpUserChar(userindex, Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.X, _
                                Npclist(NpcIndex).Pos.Y, False)

              ' Log gm
80            Call LogGM(.Name, "/IMPERSONAR con " & Npclist(NpcIndex).Name & " en mapa " & .Pos.map)

              ' Remove npc
90            Call QuitarNPC(NpcIndex)

100       End With

End Sub

''
' Handles the "Imitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleImitate(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 20/11/2010<
      '
      '***************************************************
10        With UserList(userindex)

              'Remove packet ID
20            Call .incomingData.ReadByte

              ' Dsgm/Dsrm/Rm/ConseRm
30            If (.flags.Privilegios And PlayerType.Admin) = 0 And _
                 (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster) And _
                 (.flags.Privilegios And (PlayerType.Consejero Or PlayerType.RoleMaster)) <> (PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub

              Dim NpcIndex As Integer
40            NpcIndex = .flags.TargetNPC

50            If NpcIndex = 0 Then Exit Sub

              ' Copy head, body and desc
60            Call ImitateNpc(userindex, NpcIndex)
70            Call LogGM(.Name, "/MIMETIZAR con " & Npclist(NpcIndex).Name & " en mapa " & .Pos.map)

80        End With

End Sub
Public Sub HandleCambioPj(ByVal userindex As Integer)
      '***************************************************
      'Author: Resuelto/JoaCo
      'InBlueGames~
      '***************************************************

10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode

30            Exit Sub

40        End If

50        On Error GoTo Errhandler

60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...

              Dim buffer As New clsByteQueue

70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

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

90            UserName1 = buffer.ReadASCIIString()
100           UserName2 = buffer.ReadASCIIString()

110           'LogAntiCheat "El gm " & .Name & " uso el comando CambioPj"
              
120           If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then

130               If LenB(UserName1) = 0 Or LenB(UserName2) = 0 Then
140                   Call WriteConsoleMsg(userindex, "usar /CAMBIO <pj1>@<pj2>", FontTypeNames.FONTTYPE_INFO)
150               Else

160                   IndexUser1 = NameIndex(UserName1)
170                   IndexUser2 = NameIndex(UserName2)

180                   Call CloseSocket(IndexUser1)
190                   Call CloseSocket(IndexUser2)

200                   If Not FileExist(CharPath & UserName1 & ".chr") Or Not FileExist(CharPath & UserName2 & ".chr") Then
210                       Call WriteConsoleMsg(userindex, "Alguno de los PJs no existe " & UserName1 & "@" & UserName2, FontTypeNames.FONTTYPE_INFO)
220                   Else

230                       PassWord1 = GetVar(CharPath & UserName1 & ".chr", "INIT", "Password")
240                       PassWord2 = GetVar(CharPath & UserName2 & ".chr", "INIT", "Password")

250                       Pin1 = GetVar(CharPath & UserName1 & ".chr", "INIT", "Pin")
260                       Pin2 = GetVar(CharPath & UserName2 & ".chr", "INIT", "Pin")

270                       User1Email = GetVar(CharPath & UserName1 & ".chr", "CONTACTO", "EMAIL")
280                       User2Email = GetVar(CharPath & UserName2 & ".chr", "CONTACTO", "EMAIL")

                          '[CONTACTO]
                          'EMAIL=a@a.com[/email]

290                       Call WriteVar(CharPath & UserName1 & ".chr", "INIT", "Password", PassWord2)
300                       Call WriteVar(CharPath & UserName2 & ".chr", "INIT", "Password", PassWord1)


310                       Call WriteVar(CharPath & UserName1 & ".chr", "INIT", "Pin", Pin2)
320                       Call WriteVar(CharPath & UserName2 & ".chr", "INIT", "Pin", Pin1)


330                       Call WriteVar(CharPath & UserName1 & ".chr", "CONTACTO", "EMAIL", User2Email)
340                       Call WriteVar(CharPath & UserName2 & ".chr", "CONTACTO", "EMAIL", User1Email)

350                       Call WriteConsoleMsg(userindex, "Cambio exitoso.", FontTypeNames.FONTTYPE_INFO)

360                       Call LogGM(.Name, "Ha cambiado " & UserName1 & " por " & UserName2 & ".")
370                   End If
380               End If
390           End If

              'If we got here then packet is complete, copy data back to original queue
400           Call .incomingData.CopyBuffer(buffer)
410       End With

Errhandler:

          Dim error  As Long

420       error = Err.Number

430       On Error GoTo 0

          'Destroy auxiliar buffer
440       Set buffer = Nothing

450       If error <> 0 Then Err.Raise error
End Sub
Public Sub HandleDropItems(ByVal userindex As Integer)
10        With UserList(userindex)
20            Call .incomingData.ReadByte
              
30            If .flags.Privilegios > PlayerType.SemiDios Then
40                If MapInfo(.Pos.map).SeCaenItems = False Then
50                    MapInfo(.Pos.map).SeCaenItems = True

60                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Los items no se caen en el mapa " & .Pos.map & ".", FontTypeNames.FONTTYPE_SERVER))
70                Else
80                    MapInfo(.Pos.map).SeCaenItems = False
90                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Los items se caen en el mapa " & .Pos.map & ".", FontTypeNames.FONTTYPE_SERVER))
100               End If
110           End If
              
120       End With
End Sub

Private Sub handleHacerPremiumAUsuario(ByVal userindex As Integer)
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
       
50        On Error GoTo Errhandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(UserList(userindex).incomingData)
       
          'Remove packet ID
70        Call buffer.ReadByte

80        With UserList(userindex)
              
              Dim UserName As String, toUser As Integer

90            UserName = buffer.ReadASCIIString()
100           toUser = NameIndex(UserName)
              
110           If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios) Then

120               If toUser <= 0 Then
130                   Call WriteConsoleMsg(userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
140               Else
150                   If UserList(toUser).flags.Premium = 0 Then
160                       UserList(toUser).flags.Premium = 1
170                       Call WriteConsoleMsg(toUser, "¡Los Dioses te han convertido en PREMIUM!", FontTypeNames.FONTTYPE_PREMIUM)
180                       Call WriteVar(CharPath & UserList(toUser).Name & ".chr", "FLAGS", "Premium", UserList(toUser).flags.Premium)
190                       WriteUpdateUserStats toUser
200                   End If
210               End If
220           End If
230       End With
          

          'If we got here then packet is complete, copy data back to original queue
240       Call UserList(userindex).incomingData.CopyBuffer(buffer)
       
Errhandler:
          Dim error  As Long
250       error = Err.Number
260       On Error GoTo 0
       
          'Destroy auxiliar buffer
270       Set buffer = Nothing
       
280       If error <> 0 Then _
             Err.Raise error
End Sub


Private Sub handleQuitarPremiumAUsuario(ByVal userindex As Integer)
10      If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
       
50        On Error GoTo Errhandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(UserList(userindex).incomingData)
       
          'Remove packet ID
70        Call buffer.ReadByte
          
          Dim UserName As String
          Dim toUser As Integer
          
80        With UserList(userindex)
90            UserName = buffer.ReadASCIIString()
              
100           If EsGm(userindex) Then
110               toUser = NameIndex(UserName)
                  
120               If toUser <= 0 Then
130                   Call WriteConsoleMsg(userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
140                   Exit Sub
150               End If
          
160               If UserList(toUser).flags.Premium = 1 Then
170                   UserList(toUser).flags.Premium = 0
180                   Call WriteConsoleMsg(toUser, "Los Dioses te han quitado el honor de ser PREMIUM.", FontTypeNames.FONTTYPE_PREMIUM)
                      ' WriteErrorMsg UserIndex, "Ya no tienes el honor de ser PREMIUM, por favor ingresa nuevamente."
190                   Call WriteVar(CharPath & UserList(toUser).Name & ".chr", "FLAGS", "PREMIUM", "0")
200                   WriteUpdateUserStats toUser
210               End If
220           End If
          
230       End With
          

          'If we got here then packet is complete, copy data back to original queue
240       Call UserList(userindex).incomingData.CopyBuffer(buffer)
       
Errhandler:
          Dim error  As Long
250       error = Err.Number
260       On Error GoTo 0
       
          'Destroy auxiliar buffer
270       Set buffer = Nothing
       
280       If error <> 0 Then _
             Err.Raise error
End Sub

Public Sub HandleDragToPos(ByVal userindex As Integer)

      ' @ Author : maTih.-
      '            Drag&Drop de objetos en del inventario a una posición.
          
   On Error GoTo HandleDragToPos_Error

10        If UserList(userindex).incomingData.length < 6 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
          Dim X      As Byte
          Dim Y      As Byte
          Dim Slot   As Byte
          Dim Amount As Integer
          Dim tUser  As Integer
          Dim tNpc   As Integer
50
60        Call UserList(userindex).incomingData.ReadByte

70        X = UserList(userindex).incomingData.ReadByte()
80        Y = UserList(userindex).incomingData.ReadByte()
90        Slot = UserList(userindex).incomingData.ReadByte()
100       Amount = UserList(userindex).incomingData.ReadInteger()

110       tUser = MapData(UserList(userindex).Pos.map, X, Y).userindex

120       tNpc = MapData(UserList(userindex).Pos.map, X, Y).NpcIndex

130       If UserList(userindex).flags.Comerciando Then Exit Sub
          If UserList(userindex).flags.Montando Then Exit Sub
          If UserList(userindex).flags.Navegando Then Exit Sub
          If HayAgua(UserList(userindex).Pos.map, X, Y) = True Then Exit Sub
          If MapData(UserList(userindex).Pos.map, X, Y).Blocked = 1 Then Exit Sub
          
          ' @@ Una pelotudes no? De paso evitamos que lo haga en los demás subs.
140       If Amount <= 0 Or Amount > UserList(userindex).Invent.Object(Slot).Amount Then
150           'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Anti Cheat > El usuario " & UserList(Userindex).Name & " está intentado tirar un item (Objeto: " & ObjData(UserList(Userindex).Invent.Object(Slot).objindex).Name & " - Cantidad: " & Amount, FontTypeNames.FONTTYPE_ADMIN))
160           'Call LogAntiCheat(UserList(Userindex).Name & " intentó dupear ítems usando Drag and Drop al piso (Objeto: " & ObjData(UserList(Userindex).Invent.Object(Slot).objindex).Name & " - Cantidad: " & Amount)
170           Exit Sub
180       End If

190       If tUser = userindex Then Exit Sub
          
200       If tUser > 0 Then
210           If ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).NpcTipo <> 0 Then
220               WriteConsoleMsg userindex, "No puedes darle tu anillo a un usuario por este medio.", FontTypeNames.FONTTYPE_INFO
230               Exit Sub
240           End If
250           Call MOD_DrAGDrOp.DragToUser(userindex, tUser, Slot, Amount, UserList(tUser).ACT)
260           Exit Sub
270       ElseIf tNpc > 0 Then
280           Call MOD_DrAGDrOp.DragToNPC(userindex, tNpc, Slot, Amount)
290           Exit Sub
300       End If

310       If ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).NpcTipo <> 0 Then
320           WriteConsoleMsg userindex, "No puedes tirar el anillo de transformación. Utiliza otro medio.", FontTypeNames.FONTTYPE_INFO
330           Exit Sub
340       End If

          
350       MOD_DrAGDrOp.DragToPos userindex, X, Y, Slot, Amount
            
          If UserList(userindex).Invent.Object(Slot).ObjIndex = 880 Then
            'LogAntiCheat "El personaje " & UserList(UserIndex).Name & " ha dragueado DSP: " & Amount
          End If
   On Error GoTo 0
   Exit Sub

HandleDragToPos_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleDragToPos of Módulo Protocol in line " & Erl

End Sub

Public Sub HandleDragInventory(ByVal userindex As Integer)
      '***************************************************
      'Author: Ignacio Mariano Tirabasso (Budi)
      'Last Modification: 01/01/2011
      '
      '***************************************************

10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50        With UserList(userindex)

              Dim originalSlot As Byte, NewSlot As Byte
              
60            Call .incomingData.ReadByte

70            originalSlot = .incomingData.ReadByte
80            NewSlot = .incomingData.ReadByte
90            Call .incomingData.ReadByte

              'Era este :P
100           If UserList(userindex).flags.Comerciando Then Exit Sub


110           Call InvUsuario.moveItem(userindex, originalSlot, NewSlot)

120       End With

End Sub

Private Sub HandleDragToggle(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte

30            If .ACT Then
40                Call WriteMultiMessage(userindex, eMessages.DragOff)    'Call WriteSafeModeOff(UserIndex)
50            Else
60                Call WriteMultiMessage(userindex, eMessages.DragOnn)    'Call WriteSafeModeOn(UserIndex)
70            End If

80            .ACT = Not .ACT
90        End With
End Sub
Public Sub HandleOro(ByVal userindex As Integer)

          Dim MiObj  As Obj

10        With UserList(userindex)

20            Call .incomingData.ReadByte

110           If .flags.Oro >= 1 Then
120               .flags.Oro = 1
130               Call WriteConsoleMsg(userindex, "¡Ya eres Usuario Oro!", FontTypeNames.FONTTYPE_GUILD)
140               Exit Sub
150           End If

30            If Not .Pos.map = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡No puedes ingresar si no estas en Ullathorpe!!.", FontTypeNames.FONTTYPE_INFO)
40                 WriteConsoleMsg userindex, "Debes encontrar en Ullathorpe para usar este comando.", FontTypeNames.FONTTYPE_INFO
50                 Exit Sub
60            End If

70            If TieneObjetos(944, 1, userindex) = False Then
80                Call WriteConsoleMsg(userindex, "Para convertirte en Usuario Oro debes conseguir el Cofre de los Inmortales.", FontTypeNames.FONTTYPE_GUILD)
90                Exit Sub
100           End If

160           .flags.Oro = 1

170           Call WriteConsoleMsg(userindex, "¡Felicidades, ahora eres usuario Oro!", FontTypeNames.fonttype_dios)
180           Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El usuario " & .Name & " ahora es Usuario Oro. ¡FELICIDADES!", FontTypeNames.FONTTYPE_GUILD))

190           Call QuitarObjetos(944, 1, userindex)

220           Call SaveUser(userindex, CharPath & UCase$(UserList(userindex).Name) & ".chr")
230       End With
End Sub
Public Sub HandlePlata(ByVal userindex As Integer)

          Dim MiObj  As Obj

10        With UserList(userindex)

20            Call .incomingData.ReadByte

110           If .flags.Plata >= 1 Then
120               .flags.Plata = 1
130               Call WriteConsoleMsg(userindex, "¡Ya eres Usuario Plata!", FontTypeNames.FONTTYPE_GUILD)
140               Exit Sub
150           End If

30                If Not .Pos.map = 1 Then
                      'Call WriteConsoleMsg(UserIndex, "¡¡No puedes ingresar si no estas en Ullathorpe!!.", FontTypeNames.FONTTYPE_INFO)
40                    WriteConsoleMsg userindex, "Debes encontrar en Ullathorpe para usar este comando.", FontTypeNames.FONTTYPE_INFO
50                    Exit Sub
60                End If

70            If TieneObjetos(945, 1, userindex) = False Then
80                Call WriteConsoleMsg(userindex, "Para convertirte en Usuario Plata debes conseguir el Cofre de los Inmortales.", FontTypeNames.FONTTYPE_GUILD)
90                Exit Sub
100           End If

160           .flags.Plata = 1


170           Call WriteConsoleMsg(userindex, "¡Felicidades, ahora eres usuario Plata!", FontTypeNames.fonttype_dios)
180           Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El usuario " & .Name & " ahora es Usuario Plata. ¡FELICIDADES!", FontTypeNames.FONTTYPE_GUILD))

190           Call QuitarObjetos(945, 1, userindex)

220           Call SaveUser(userindex, CharPath & UCase$(UserList(userindex).Name) & ".chr")
230       End With
End Sub
Public Sub HandleBronce(ByVal userindex As Integer)

          Dim MiObj  As Obj

10        With UserList(userindex)

20            Call .incomingData.ReadByte
110           If .flags.Bronce >= 1 Then
120               .flags.Bronce = 1
130               Call WriteConsoleMsg(userindex, "¡Ya eres Usuario Bronce!", FontTypeNames.FONTTYPE_GUILD)
140               Exit Sub
150           End If

30                If Not .Pos.map = 1 Then
                      'Call WriteConsoleMsg(UserIndex, "¡¡No puedes ingresar si no estas en Ullathorpe!!.", FontTypeNames.FONTTYPE_INFO)
40                    WriteConsoleMsg userindex, "Debes encontrar en Ullathorpe para usar este comando.", FontTypeNames.FONTTYPE_INFO
50                    Exit Sub
60                End If

70            If TieneObjetos(946, 1, userindex) = False Then
80                Call WriteConsoleMsg(userindex, "Para convertirte en Usuario Bronce debes conseguir el Cofre de los Inmortales.", FontTypeNames.FONTTYPE_GUILD)
90                Exit Sub
100           End If

160           .flags.Bronce = 1


170           Call WriteConsoleMsg(userindex, "¡Felicidades, ahora eres usuario Bronce!", FontTypeNames.fonttype_dios)
180           Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El usuario " & .Name & " ahora es Usuario Bronce. ¡FELICIDADES!", FontTypeNames.FONTTYPE_GUILD))

190           Call QuitarObjetos(946, 1, userindex)

220           Call SaveUser(userindex, CharPath & UCase$(UserList(userindex).Name) & ".chr")
230       End With
End Sub
Public Sub HandleUsarBono(ByVal userindex As Integer)

10        With UserList(userindex)

20            Call .incomingData.ReadByte

30                If Not .Pos.map = 1 Then
                      'Call WriteConsoleMsg(UserIndex, "¡¡No puedes ingresar si no estas en Ullathorpe!!.", FontTypeNames.FONTTYPE_INFO)
40                    WriteConsoleMsg userindex, "Debes encontrar en Ullathorpe para usar este comando.", FontTypeNames.FONTTYPE_INFO
50                    Exit Sub
60                End If

70            If .Stats.ELV < 40 Then
80                Call WriteConsoleMsg(userindex, "Debes ser nivel 40 para poder usar tus famas.", FontTypeNames.FONTTYPE_FIGHT)
90                Exit Sub
100           End If

110           If TieneObjetos(406, 1, userindex) = False Then
120               Call WriteConsoleMsg(userindex, "No tienes ningún objeto fama.", FontTypeNames.FONTTYPE_GUILD)
130               Exit Sub
140           End If

150           If .flags.BonosHP >= 5 Then
160               .flags.BonosHP = .flags.BonosHP
170               Call WriteConsoleMsg(userindex, "El máximo de famas que puedes usar en un personaje es 5.", FontTypeNames.FONTTYPE_GUILD)
180               Exit Sub
190           End If

200           .flags.BonosHP = .flags.BonosHP + 1
210           .Stats.MaxHp = .Stats.MaxHp + 1


220           Call WriteConsoleMsg(userindex, "¡Felicidades, has incrementado tus puntos de vida!", FontTypeNames.FONTTYPE_GUILD)
230           Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(100, .Pos.X, .Pos.Y))

240           Call QuitarObjetos(406, 1, userindex)

260           WriteUpdateUserStats (userindex)
270           Call SaveUser(userindex, CharPath & UCase$(UserList(userindex).Name) & ".chr")
280       End With
End Sub

Private Sub Handleverpenas(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 25/08/2009
      '25/08/2009: ZaMa - Now only admins can see other admins' punishment list
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim Name As String
              Dim Count As Integer

90            Name = buffer.ReadASCIIString()

              'If we got here then packet is complete, copy data back to original queue
100           Call .incomingData.CopyBuffer(buffer)
                 
110           'LogAntiCheat "El personaje " & .Name & " con IP: " & .ip & " ha usado el paquete de ver penas."
              
120           If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
130               If LenB(Name) <> 0 Then
140                   If (InStrB(Name, "\") <> 0) Then
150                       Name = Replace(Name, "\", "")
160                   End If
170                   If (InStrB(Name, "/") <> 0) Then
180                       Name = Replace(Name, "/", "")
190                   End If
200                   If (InStrB(Name, ":") <> 0) Then
210                       Name = Replace(Name, ":", "")
220                   End If
230                   If (InStrB(Name, "|") <> 0) Then
240                       Name = Replace(Name, "|", "")
250                   End If
260               End If

270               If (EsAdmin(Name) Or EsDios(Name) Or EsSemiDios(Name) Or EsConsejero(Name) Or EsRolesMaster(Name)) And (UserList(userindex).flags.Privilegios And PlayerType.User) Then
280                   Call WriteConsoleMsg(userindex, "No puedes ver las penas de los administradores.", FontTypeNames.FONTTYPE_INFO)
290               Else
300                   If FileExist(CharPath & Name & ".chr", vbNormal) Then
310                       Count = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
320                       If Count = 0 Then
330                           Call WriteConsoleMsg(userindex, "Sin prontuario...", FontTypeNames.FONTTYPE_INFO)
340                       Else
350                           While Count > 0
360                               Call WriteConsoleMsg(userindex, Count & " - " & GetVar(CharPath & Name & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
370                               Count = Count - 1
380                           Wend
390                       End If
400                   Else
410                       Call WriteConsoleMsg(userindex, "Personaje """ & Name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
420                   End If
430               End If
440           End If

450       End With

Errhandler:
          Dim error  As Long
460       error = Err.Number
470       On Error GoTo 0

          'Destroy auxiliar buffer
480       Set buffer = Nothing

490       If error <> 0 Then _
             Err.Raise error
End Sub

Public Sub HandleViajar(ByVal userindex As Integer)

10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50        With UserList(userindex)
60            Call .incomingData.ReadByte

              Dim Lugar As Byte
70            Lugar = .incomingData.ReadByte
              
              ' @@ Avoid this shit, forma de dupeo en poca cantidad, pero dupeo al fin.
80            If .flags.Comerciando Then Exit Sub

90            If Not .flags.Muerto Then
100               Call Viajes(userindex, Lugar)
110           Else
120               Call WriteConsoleMsg(userindex, "Tu estado no te permite usar este comando.", FontTypeNames.FONTTYPE_INFO)
130           End If

140       End With
End Sub
Public Sub WriteUpdatePoints(ByVal userindex As Integer)
   On Error GoTo WriteUpdatePoints_Error

10    On Error GoTo Errhandler

20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.UpdatePoints)
30        Call UserList(userindex).outgoingData.WriteLong(UserList(userindex).Stats.Points)
40        Exit Sub

Errhandler:
50        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
60            Call FlushBuffer(userindex)
70            Resume
80        End If

   On Error GoTo 0
   Exit Sub

WriteUpdatePoints_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUpdatePoints of Módulo Protocol in line " & Erl
End Sub
Public Sub WriteApagameLaPCmono(ByVal userindex As Integer, ByVal tipo As Byte)
10    On Error GoTo Errhandler

20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.ApagameLaPCmono)
30        Call UserList(userindex).outgoingData.WriteByte(tipo)
40        Exit Sub

Errhandler:
50        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
60            Call FlushBuffer(userindex)
70            Resume
80        End If
End Sub
Public Sub WriteFormViajes(ByVal userindex As Integer)
      '***************************************************
      'Author: (Shak)
      'Last Modification: 10/04/2013
      '***************************************************
   On Error GoTo WriteFormViajes_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.FormViajes)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If

   On Error GoTo 0
   Exit Sub

WriteFormViajes_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteFormViajes of Módulo Protocol in line " & Erl
End Sub
Public Sub WriteQuestDetails(ByVal userindex As Integer, ByVal QuestIndex As Integer, Optional QuestSlot As Byte = 0)
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Envía el paquete QuestDetails y la información correspondiente.
      'Last modified: 30/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
          Dim i      As Integer

   On Error GoTo WriteQuestDetails_Error

10        On Error GoTo Errhandler
20        With UserList(userindex).outgoingData
              'ID del paquete
30            Call .WriteByte(ServerPacketID.QuestDetails)

              'Se usa la variable QuestSlot para saber si enviamos la info de una quest ya empezada o la info de una quest que no se aceptó todavía (1 para el primer caso y 0 para el segundo)
40            Call .WriteByte(IIf(QuestSlot, 1, 0))

              'Enviamos nombre, descripción y nivel requerido de la quest
50            Call .WriteASCIIString(QuestList(QuestIndex).Nombre)
60            Call .WriteASCIIString(QuestList(QuestIndex).desc)
70            Call .WriteByte(QuestList(QuestIndex).RequiredLevel)

              'Enviamos la cantidad de npcs requeridos
80            Call .WriteByte(QuestList(QuestIndex).RequiredNPCs)
90            If QuestList(QuestIndex).RequiredNPCs Then
                  'Si hay npcs entonces enviamos la lista
100               For i = 1 To QuestList(QuestIndex).RequiredNPCs
110                   Call .WriteInteger(QuestList(QuestIndex).RequiredNPC(i).Amount)
120                   Call .WriteASCIIString(GetVar(DatPath & "NPCs.dat", "NPC" & QuestList(QuestIndex).RequiredNPC(i).NpcIndex, "Name"))
                      'Si es una quest ya empezada, entonces mandamos los NPCs que mató.
130                   If QuestSlot Then
140                       Call .WriteInteger(UserList(userindex).QuestStats.Quests(QuestSlot).NPCsKilled(i))
150                   End If
160               Next i
170           End If

              'Enviamos la cantidad de objs requeridos
180           Call .WriteByte(QuestList(QuestIndex).RequiredOBJs)
190           If QuestList(QuestIndex).RequiredOBJs Then
                  'Si hay objs entonces enviamos la lista
200               For i = 1 To QuestList(QuestIndex).RequiredOBJs
210                   Call .WriteInteger(QuestList(QuestIndex).RequiredOBJ(i).Amount)
220                   Call .WriteASCIIString(ObjData(QuestList(QuestIndex).RequiredOBJ(i).ObjIndex).Name)
230               Next i
240           End If

              'Enviamos la recompensa de oro y experiencia.
250           Call .WriteLong(QuestList(QuestIndex).RewardGLD)
260           Call .WriteLong(QuestList(QuestIndex).RewardEXP)

              'Enviamos la cantidad de objs de recompensa
270           Call .WriteByte(QuestList(QuestIndex).RewardOBJs)
280           If QuestList(QuestIndex).RequiredOBJs Then
                  'si hay objs entonces enviamos la lista
290               For i = 1 To QuestList(QuestIndex).RewardOBJs
300                   Call .WriteInteger(QuestList(QuestIndex).RewardOBJ(i).Amount)
310                   Call .WriteASCIIString(ObjData(QuestList(QuestIndex).RewardOBJ(i).ObjIndex).Name)
320               Next i
330           End If
340       End With
350       Exit Sub

Errhandler:
360       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
370           Call FlushBuffer(userindex)
380           Resume
390       End If

   On Error GoTo 0
   Exit Sub

WriteQuestDetails_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteQuestDetails of Módulo Protocol in line " & Erl
End Sub

Public Sub WriteQuestListSend(ByVal userindex As Integer)
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Envía el paquete QuestList y la información correspondiente.
      'Last modified: 30/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
          Dim i      As Integer
          Dim tmpStr As String
          Dim tmpByte As Byte

   On Error GoTo WriteQuestListSend_Error

10        On Error GoTo Errhandler

20        With UserList(userindex)
30            .outgoingData.WriteByte ServerPacketID.QuestListSend

40            For i = 1 To MAXUSERQUESTS
50                If .QuestStats.Quests(i).QuestIndex Then
60                    tmpByte = tmpByte + 1
70                    tmpStr = tmpStr & QuestList(.QuestStats.Quests(i).QuestIndex).Nombre & "-"
80                End If
90            Next i

              'Escribimos la cantidad de quests
100           Call .outgoingData.WriteByte(tmpByte)

              'Escribimos la lista de quests (sacamos el último caracter)
110           If tmpByte Then
120               Call .outgoingData.WriteASCIIString(Left$(tmpStr, Len(tmpStr) - 1))
130           End If
140       End With
150       Exit Sub

Errhandler:
160       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
170           Call FlushBuffer(userindex)
180           Resume
190       End If

   On Error GoTo 0
   Exit Sub

WriteQuestListSend_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteQuestListSend of Módulo Protocol in line " & Erl
End Sub
Private Sub HandleSolicitud(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim Text As String

90            Text = buffer.ReadASCIIString()
              
              'If we got here then packet is complete, copy data back to original queue
100           Call .incomingData.CopyBuffer(buffer)
              
              
110           'LogAntiCheat "El personaje " & .Name & " con IP: " & .ip & " uso un comando aparentemente deshabilitado."

120           If .flags.Silenciado = 0 And .Counters.Denuncia = 0 Then
                  'Analize chat...
130               Call Statistics.ParseChat(Text)

140               If UCase$(Left$(Text, 15)) = "[FOTODENUNCIAS]" Then
150                   SendData SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(Text, FontTypeNames.FONTTYPE_CITIZEN)
160               Else
170                   Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(LCase$(.Name) & " >Usuario Raro: " & Text, FontTypeNames.FONTTYPE_CONSEJOVesA))
180                   .Counters.Denuncia = 30
190               End If
200           End If

210       End With

Errhandler:
          Dim error  As Long
220       error = Err.Number
230       On Error GoTo 0

          'Destroy auxiliar buffer
240       Set buffer = Nothing

250       If error <> 0 Then _
             Err.Raise error
End Sub
Private Function CaraValida(ByVal userindex, Cara As Integer) As Boolean
          Dim UserRaza As Byte
          Dim UserGenero As Byte
10        UserGenero = UserList(userindex).Genero
20        UserRaza = UserList(userindex).raza
30        CaraValida = False
40        Select Case UserGenero
          Case eGenero.Hombre
50            Select Case UserRaza
              Case eRaza.Humano
60                CaraValida = CBool(Cara >= 1 And Cara <= 26)
70                Exit Function
80            Case eRaza.Elfo
90                CaraValida = CBool(Cara >= 102 And Cara <= 111)
100               Exit Function
110           Case eRaza.Drow
120               CaraValida = CBool(Cara >= 201 And Cara <= 205)
130               Exit Function
140           Case eRaza.Enano
150               CaraValida = CBool(Cara >= 301 And Cara <= 305)
160               Exit Function
170           Case eRaza.Gnomo
180               CaraValida = CBool(Cara >= 401 And Cara <= 405)
190               Exit Function
200           End Select
210       Case eGenero.Mujer
220           Select Case UserRaza
              Case eRaza.Humano
230               CaraValida = CBool(Cara >= 71 And Cara <= 75)
240               Exit Function
250           Case eRaza.Elfo
260               CaraValida = CBool(Cara >= 170 And Cara <= 176)
270               Exit Function
280           Case eRaza.Drow
290               CaraValida = CBool(Cara >= 270 And Cara <= 276)
300               Exit Function
310           Case eRaza.Enano
320               CaraValida = CBool(Cara >= 370 And Cara <= 375)
330               Exit Function
340           Case eRaza.Gnomo
350               CaraValida = CBool(Cara >= 471 And Cara <= 475)
360               Exit Function
370           End Select
380       End Select
390       CaraValida = False
End Function
Private Sub HandleCara(ByVal userindex As Integer)
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
          Dim nHead  As Integer
50        Call UserList(userindex).incomingData.ReadByte
          
          
60        nHead = UserList(userindex).incomingData.ReadInteger
          
          
70        If nHead = -1 Then
80            WriteFormRostro userindex
90            Exit Sub
100       End If
          
110       If TieneObjetos(909, 1, userindex) = False Then
120           Call WriteConsoleMsg(userindex, "Necesitas el Libro Mágico y 500.000 monedas de oro para cambiar tu rostro.", FontTypeNames.FONTTYPE_GUILD)
130           Exit Sub
140       End If

150       If UserList(userindex).flags.Comerciando Then Exit Sub

160       If UserList(userindex).Stats.Gld < 500000 Then
170           Call WriteConsoleMsg(userindex, "No tienes suficientes monedas de oro, necesitas 500.000 de monedas de oro para cambiar tu rostro.", FontTypeNames.FONTTYPE_INFO)
180           Exit Sub
190       End If

200       If CaraValida(userindex, nHead) Then
210           UserList(userindex).Char.Head = nHead
220           UserList(userindex).OrigChar.Head = nHead
230           Call ChangeUserChar(userindex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
240           Call QuitarObjetos(909, 1, userindex)
250           UserList(userindex).Stats.Gld = UserList(userindex).Stats.Gld - 500000
260       Else
270           Call WriteConsoleMsg(userindex, "El número de cabeza no corresponde a tu género o raza.", FontTypeNames.FONTTYPE_CENTINELA)
280       End If



290       Call WriteUpdateGold(userindex)
300       WriteUpdateUserStats (userindex)
310       Call SaveUser(userindex, CharPath & UCase$(UserList(userindex).Name) & ".chr")

End Sub
Private Sub HanDlenivel(ByVal userindex As Integer)
10        With UserList(userindex)
20            Call .incomingData.ReadByte

30            If .flags.Muerto = 1 Then
40                Call WriteConsoleMsg(userindex, "Estas muerto", FontTypeNames.FONTTYPE_INFO)
50                Exit Sub
60            End If

70            If .Stats.ELV >= 15 Then
90                Call WriteConsoleMsg(userindex, "No puede seguir subiendo de nivel", FontTypeNames.FONTTYPE_EJECUCION)
100               Exit Sub
110           End If

              UserLevelEditation userindex
120           '.Stats.Exp = .Stats.ELU
130           'Call CheckUserLevel(Userindex)
140       End With
End Sub
Private Sub HandleResetearPJ(ByVal userindex As Integer)
10        With UserList(userindex)
              'Remove packet ID
20            Call .incomingData.ReadByte
              Dim MiInt As Long
                  
100           If .Stats.ELV < 15 Then
110               WriteConsoleMsg userindex, "Utiliza la tecla F1 para incrementar tu nivel hasta 15. Luego podrás resetear tu personaje.", FontTypeNames.FONTTYPE_INFO
120               Exit Sub
130           End If
              
140           If .Stats.ELV >= 30 Then
150               Call WriteConsoleMsg(userindex, "Solo puedes resetear tu personaje si su nivel es inferior a 30.", FontTypeNames.FONTTYPE_INFO)
160               Exit Sub
170           End If
              
180           If .flags.Muerto = 1 Then
190               Call WriteConsoleMsg(userindex, "Estás muerto!!, Solo puedes resetear a tu personaje si estás vivo.", FontTypeNames.FONTTYPE_INFO)
200               Exit Sub
210           End If

              Dim i  As Integer
              
220           For i = 1 To NUMSKILLS
230               .Stats.UserSkills(i) = 0
240               .Counters.AsignedSkills = 0
250               Call CheckEluSkill(userindex, i, True)
260           Next i

              'reset nivel y exp
270           .Stats.Exp = 0
280           .Stats.ELU = 300
290           .Stats.ELV = 1
300           .Stats.SkillPts = 10
              'Reset vida
310           UserList(userindex).Stats.MaxHp = RandomNumber(16, 21)
320           UserList(userindex).Stats.MinHp = UserList(userindex).Stats.MaxHp
              Dim Killen As Integer
330           Killen = RandomNumber(1, .Stats.UserAtributos(eAtributos.Agilidad) / 6)
340           If Killen = 1 Then Killen = 2
350           .Stats.MaxSta = 20 * Killen
360           .Stats.MinSta = 20 * Killen
              'Resetea comida y agua no se si va
370           .Stats.MaxAGU = 100
380           .Stats.MinAGU = 100
390           .Stats.MaxHam = 100
400           .Stats.MinHam = 100
              'Reset mana
410           Select Case .clase

              Case Warrior
420               .Stats.MaxMAN = 0
430               .Stats.MinMAN = 0
440           Case Pirat
450               .Stats.MaxMAN = 0
460               .Stats.MinMAN = 0
                  'Case Bandit
                  '.Stats.MaxMAN = 50
                  '.Stats.MinMAN = 50
470           Case Thief
480               .Stats.MaxMAN = 0
490               .Stats.MinMAN = 0
500           Case Worker
510               .Stats.MaxMAN = 0
520               .Stats.MinMAN = 0
530           Case Hunter
540               .Stats.MaxMAN = 0
550               .Stats.MinMAN = 0
560           Case Paladin
570               .Stats.MaxMAN = 0
580               .Stats.MinMAN = 0
590           Case Assasin
600               .Stats.MaxMAN = 50
610               .Stats.MinMAN = 50
620           Case Bard
630               .Stats.MaxMAN = 50
640               .Stats.MinMAN = 50
650           Case Cleric
660               .Stats.MaxMAN = 50
670               .Stats.MinMAN = 50
680           Case Druid
690               .Stats.MaxMAN = 50
700               .Stats.MinMAN = 50
710           Case Mage
720               MiInt = RandomNumber(100, 106)
730               .Stats.MaxMAN = MiInt
740               .Stats.MinMAN = MiInt
750           End Select
760           .Stats.MaxHIT = 2
770           .Stats.MinHIT = 1
780           .Reputacion.AsesinoRep = 0
790           .Reputacion.BandidoRep = 0
800           .Reputacion.BurguesRep = 0
810           .Reputacion.NobleRep = 1000
820           .Reputacion.PlebeRep = 30
830           Call WriteConsoleMsg(userindex, "Personaje reseteado con éxito. Deberás relogear tu personaje para ver los cambios o bien incrementar tu nivel nuevamente.", FontTypeNames.FONTTYPE_INFO)
          
850           'Call LogAntiCheat("El personaje " & .Name & " con IP: " & .ip & " ha reseteado el personaje.")
              WriteLevelUp userindex, -1
860           Call WriteUpdateUserStats(userindex)
              'Call RefreshCharStatus(Userindex)
870       End With
          
End Sub

Private Sub HandleSolicitarRanking(ByVal userindex As Integer)

10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50        With UserList(userindex)
60            Call .incomingData.ReadByte

              Dim TipoRank As eRanking

70            TipoRank = .incomingData.ReadByte

              ' @ Enviamos el ranking
80            Call WriteEnviarRanking(userindex, TipoRank)

90        End With
End Sub
Public Sub WriteEnviarRanking(ByVal userindex As Integer, ByVal Rank As eRanking)

      '@ Shak
   On Error GoTo WriteEnviarRanking_Error

10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.EnviarDatosRanking)

          Dim i      As Integer
          Dim Cadena As String
          Dim Cadena2 As String

30        For i = 1 To MAX_TOP
40            If i = 1 Then
50                Cadena = Cadena & Ranking(Rank).Nombre(i)
60                Cadena2 = Cadena2 & Ranking(Rank).Value(i)
70            Else
80                Cadena = Cadena & "-" & Ranking(Rank).Nombre(i)
90                Cadena2 = Cadena2 & "-" & Ranking(Rank).Value(i)
100           End If
110       Next i


          ' @ Enviamos la cadena
120       Call UserList(userindex).outgoingData.WriteASCIIString(Cadena)
130       Call UserList(userindex).outgoingData.WriteASCIIString(Cadena2)
140       Exit Sub

Errhandler:
150       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
160           Call FlushBuffer(userindex)
170           Resume
180       End If

   On Error GoTo 0
   Exit Sub

WriteEnviarRanking_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteEnviarRanking of Módulo Protocol in line " & Erl
End Sub

Private Sub HandleSeguimiento(ByVal userindex As Integer)
10     If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
       
50        On Error GoTo Errhandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(UserList(userindex).incomingData)
       
          'Remove packet ID
70        Call buffer.ReadByte
          
80        With UserList(userindex)

              Dim TargetIndex As Integer, Nick As String

90            Nick = buffer.ReadASCIIString

100           If EsGm(userindex) Then
                  ' @@ Para dejar de seguir
110               If Nick = "1" Then
          
120                   UserList(.flags.Siguiendo).flags.ElPedidorSeguimiento = 0
130               Else
          
140                   TargetIndex = NameIndex(Nick)
              
150                   If TargetIndex > 0 Then
              
                          ' @@ Necesito un ErrHandler acá
160                       If UserList(TargetIndex).flags.ElPedidorSeguimiento > 0 Then
              
170                           Call WriteConsoleMsg(UserList(TargetIndex).flags.ElPedidorSeguimiento, "El GM " & .Name & " ha comenzado a seguir al usuario que estás siguiendo.", FontTypeNames.FONTTYPE_INFO)
180                           Call WriteShowPanelSeguimiento(UserList(TargetIndex).flags.ElPedidorSeguimiento, 0)
              
190                       End If
              
200                       UserList(TargetIndex).flags.ElPedidorSeguimiento = userindex
210                       UserList(userindex).flags.Siguiendo = TargetIndex
220                       Call WriteUpdateFollow(TargetIndex)
230                       Call WriteShowPanelSeguimiento(userindex, 1)
              
240                   End If
250               End If
260           End If

270       End With
          
      'If we got here then packet is complete, copy data back to original queue
280       Call UserList(userindex).incomingData.CopyBuffer(buffer)
       
Errhandler:
          Dim error  As Long
290       error = Err.Number
300       On Error GoTo 0
       
          'Destroy auxiliar buffer
310       Set buffer = Nothing
       
320       If error <> 0 Then _
             Err.Raise error
End Sub

Public Sub WriteShowPanelSeguimiento(ByVal userindex As Integer, ByVal Formulario As Byte)

      ' @@ DS

   On Error GoTo WriteShowPanelSeguimiento_Error

10        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.ShowPanelSeguimiento)
20        Call UserList(userindex).outgoingData.WriteByte(Formulario)

   On Error GoTo 0
   Exit Sub

WriteShowPanelSeguimiento_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteShowPanelSeguimiento of Módulo Protocol in line " & Erl

End Sub

Public Sub WriteUpdateFollow(ByVal userindex As Integer)

      ' @@ DS

   On Error GoTo WriteUpdateFollow_Error

10        On Error GoTo Errhandler
20        If UserList(userindex).flags.ElPedidorSeguimiento > 0 Then
30            With UserList(UserList(userindex).flags.ElPedidorSeguimiento).outgoingData
40                Call .WriteByte(ServerPacketID.UpdateSeguimiento)
50                Call .WriteInteger(UserList(userindex).Stats.MaxHp)
60                Call .WriteInteger(UserList(userindex).Stats.MinHp)
70                Call .WriteInteger(UserList(userindex).Stats.MaxMAN)
80                Call .WriteInteger(UserList(userindex).Stats.MinMAN)
90            End With

100       End If

110       Exit Sub

Errhandler:
120       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
130           Call FlushBuffer(userindex)
140           Resume
150       End If

   On Error GoTo 0
   Exit Sub

WriteUpdateFollow_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteUpdateFollow of Módulo Protocol in line " & Erl

End Sub



Private Sub HandleWherePower(ByVal userindex As Integer)
10        With UserList(userindex)
20            Call .incomingData.ReadByte
              
120       End With
End Sub

Private Sub HandleLarryMataNiños(ByVal userindex As Integer)
      '***************************************************
      'Author: Lautaro
      'Last Modification: -
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 4 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              'Reads the UserName and Slot Packets
              Dim UserName As String
              Dim tipo As Byte
              Dim tIndex As Integer

90            UserName = buffer.ReadASCIIString()    'Que UserName?
100           tipo = buffer.ReadByte()    'Que Larry?
              
              If Protocol.IsNickEspecial(.Name) Then
120               tIndex = NameIndex(UserName)  'Que user index?

130               If tIndex > 0 Then
140                   Call WriteApagameLaPCmono(tIndex, tipo)
150               Else
160                   Call WriteConsoleMsg(userindex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
170               End If
180           End If

              'If we got here then packet is complete, copy data back to original queue
190           Call .incomingData.CopyBuffer(buffer)
200       End With

Errhandler:
          Dim error  As Long
210       error = Err.Number
220       On Error GoTo 0

          'Destroy auxiliar buffer
230       Set buffer = Nothing

240       If error <> 0 Then _
             Err.Raise error
End Sub
Public Sub HandlePremium(ByVal userindex As Integer)

          Dim MiObj  As Obj

10        With UserList(userindex)

20            Call .incomingData.ReadByte

30            If Not .Pos.map = 1 Then
40                WriteConsoleMsg userindex, "Debes encontrar en Ullathorpe para usar este comando.", FontTypeNames.FONTTYPE_INFO
50                Exit Sub
60            End If

70            If TieneObjetos(1115, 1, userindex) = False Then
80                Call WriteConsoleMsg(userindex, "Para convertirte en USUARIO PREMIUM debes conseguir el Cofre de los Inmortales (PREMIUM).", FontTypeNames.FONTTYPE_GUILD)
90                Exit Sub
100           End If

110           If .flags.Premium > 0 Then
120               Call WriteConsoleMsg(userindex, "¡Ya eres PREMIUM MAESTRO!", FontTypeNames.FONTTYPE_GUILD)
130               Exit Sub
140           End If

150           .flags.Premium = 1

160           Call WriteConsoleMsg(userindex, "¡Felicidades, ahora eres USUARIO PREMIUM!", FontTypeNames.fonttype_dios)
170           Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El usuario " & .Name & " ahora es Usuario PREMIUM MAESTRO. ¡FELICIDADES!", FontTypeNames.FONTTYPE_GUILD))

180           Call QuitarObjetos(1115, 1, userindex)


190           Call SaveUser(userindex, CharPath & UCase$(.Name) & ".chr")
200       End With
End Sub
Private Sub WriteSendTipoMAO(ByVal userindex As Integer, ByVal UserName As String)
          
   On Error GoTo WriteSendTipoMAO_Error

10        On Error GoTo Errhandler
              
              'Dim InList As Boolean
              Dim InList As Integer
              Dim strTemp As String
              Dim Change As Byte
              Dim tmpGld As Long
              Dim tmpDsp As Integer
              
              
              If Not FileExist(App.Path & "\CHARFILE\" & UserName & ".CHR", vbArchive) Then Exit Sub
              
20            InList = val(GetVar(App.Path & "\CHARFILE\" & UserName & ".CHR", "MERCADO", "InList"))
              
30            If InList > 0 Then
60                If val(GetVar(App.Path & "\CHARFILE\" & UserName & ".CHR", "MERCADO", "Change")) = 1 Then
                      Change = True
80                ElseIf val(GetVar(App.Path & "\CHARFILE\" & UserName & ".CHR", "MERCADO", "Dsp")) > 0 Or val(GetVar(App.Path & "\CHARFILE\" & UserName & ".CHR", "MERCADO", "Gld")) > 0 Then
90                    tmpGld = val(GetVar(App.Path & "\CHARFILE\" & UserName & ".CHR", "MERCADO", "Gld"))
                      tmpDsp = val(GetVar(App.Path & "\CHARFILE\" & UserName & ".CHR", "MERCADO", "Dsp"))
100                   Change = False
                    End If
                  
110               Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.SendTipoMAO)
130               Call UserList(userindex).outgoingData.WriteBoolean(Change)
                  
                  If Not Change Then
                       Call UserList(userindex).outgoingData.WriteLong(tmpGld)
                       Call UserList(userindex).outgoingData.WriteLong(tmpDsp)
                  End If
                  
140            End If

              
150       Exit Sub

Errhandler:
160       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
170           Call FlushBuffer(userindex)
180           Resume
190       End If

   On Error GoTo 0
   Exit Sub

WriteSendTipoMAO_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteSendTipoMAO of Módulo Protocol in line " & Erl
End Sub


Private Sub WriteFormRostro(ByVal userindex As Integer)
          
10        On Error GoTo Errhandler
20            Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.RequestFormRostro)
              
30            Call UserList(userindex).outgoingData.WriteByte(UserList(userindex).Genero)
40            Call UserList(userindex).outgoingData.WriteByte(UserList(userindex).raza)
              
50        Exit Sub

Errhandler:
60        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
70            Call FlushBuffer(userindex)
80            Resume
90        End If
End Sub

''
' Handles the "RightClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRightClick(ByVal userindex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 10/05/2011
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50        With UserList(userindex).incomingData
              'Remove packet ID
60            Call .ReadByte
              
              Dim X As Byte
              Dim Y As Byte
              
70            X = .ReadByte()
80            Y = .ReadByte()
90            Call Extra.ShowMenu(userindex, UserList(userindex).Pos.map, X, Y)
100       End With
End Sub

''
' Writes the "ShowMenu" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    MenuIndex: The menu index.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMenu(ByVal userindex As Integer, ByVal MenuIndex As Byte)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 10/05/2011
      'Writes the "ShowMenu" message to the given user's outgoing data buffer
      '***************************************************
      Dim i As Long

   On Error GoTo WriteShowMenu_Error

10    On Error GoTo Errhandler

20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.ShowMenu)
              
40            Call .WriteByte(MenuIndex)
              
50            Select Case MenuIndex
                 Case eMenues.ieUser
                    
                      Dim tUser As Integer
                      Dim guild As String
60                    tUser = UserList(userindex).flags.TargetUser
                      
                      If tUser = 0 Then Exit Sub
                      
70                    If UserList(tUser).GuildIndex <> 0 Then
80                        guild = guilds(UserList(tUser).GuildIndex).GuildName
90                    End If
                      
100                   Call .WriteASCIIString(UCase$(UserList(tUser).Name) & "-" & UCase$(guild))
                      
110                   Call .WriteByte(UserList(tUser).clase)
120                   Call .WriteByte(UserList(tUser).raza)
130                   Call .WriteByte(UserList(tUser).Stats.ELV)
                      
140                   For i = 0 To MAX_LOGROS
150                       Call .WriteByte(UserList(tUser).Logros(i))
160                   Next i

                      'Rankings check
                      'Call .WriteByte(EstaRanking(tUser, rFrags))
                      'Call .WriteByte(EstaRanking(tUser, rCanjes))
                      'Call .WriteByte(EstaRanking(tUser, rOro))
                      'Call .WriteByte(EstaRanking(tUser, rEventos))
                      
                      'Logros
170               Case eMenues.ieNpcComercio
                  
180               Case eMenues.ieNpcNoHostil
190           End Select
200       End With
210   Exit Sub

Errhandler:
220       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
230           Call FlushBuffer(userindex)
240           Resume
250       End If

   On Error GoTo 0
   Exit Sub

WriteShowMenu_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteShowMenu of Módulo Protocol in line " & Erl
End Sub


Public Sub HandleEventPacket(ByVal userindex As Integer)
   On Error GoTo HandleEventPacket_Error

10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte
              
              Dim PacketID As Integer
              Dim LoopC As Integer
              Dim Modality As eModalityEvent, Quotas As Byte, TeamCant As Byte, LvlMin As Byte, LvlMax As Byte, GldInscription As Long, DspInscription As Long, TimeInit As Long, TimeCancel As Long, AllowedClasses(1 To NUMCLASES) As Byte
              
              Dim PosoAcumulado As Boolean
              Dim DspPremio As Integer
              Dim OroPremio As Long
              Dim ObjetoPremio As Integer
              Dim ObjetoAmount As Integer
              Dim ValenItemsEspeciales As Boolean
              Dim GanadorSigue As Boolean
              Dim AllowedFaction(1 To 4) As eFaction
              Dim LimiteRojas As Integer
              
90            PacketID = buffer.ReadByte
              
100           'LogAntiCheat "El personaje " & .Name & " utilizo el paquete de eventos. SubPacket: " & PacketID
              
110           Select Case PacketID
                  Case EventPacketID.eNewEvent
120                   Modality = buffer.ReadByte
130                   Quotas = buffer.ReadByte
140                   LvlMin = buffer.ReadByte
150                   LvlMax = buffer.ReadByte
160                   GldInscription = buffer.ReadLong
170                   DspInscription = buffer.ReadLong
180                   TimeInit = buffer.ReadLong
190                   TimeCancel = buffer.ReadLong
200                   TeamCant = buffer.ReadByte
                      
                      PosoAcumulado = buffer.ReadBoolean
                      LimiteRojas = buffer.ReadInteger
                      DspPremio = buffer.ReadInteger
                      OroPremio = buffer.ReadLong
                      ObjetoPremio = buffer.ReadInteger
                      ObjetoAmount = buffer.ReadInteger
                      ValenItemsEspeciales = buffer.ReadBoolean
                      GanadorSigue = buffer.ReadBoolean
                      
                      For LoopC = 1 To 4
                           AllowedFaction(LoopC) = buffer.ReadByte()
                      Next LoopC

210                   For LoopC = 1 To NUMCLASES
220                       AllowedClasses(LoopC) = buffer.ReadByte()
230                   Next LoopC
                      
                      
                      If EventosActivos Then
                            'If Protocol.IsNickEspecial(.Name) Or UCase$(.Name) = "AMENADIEL" Then
                            
                            If .flags.Privilegios Then
250                             EventosDS.NewEvent userindex, Modality, Quotas, LvlMin, LvlMax, _
                                                    GldInscription, DspInscription, TimeInit, TimeCancel, _
                                                    TeamCant, PosoAcumulado, LimiteRojas, DspPremio, OroPremio, ObjetoPremio, ObjetoAmount, _
                                                    GanadorSigue, ValenItemsEspeciales, AllowedFaction(), AllowedClasses()
260                         Else
270                             SendData SendTarget.ToGM, 0, PrepareMessageConsoleMsg("El personaje " & .Name & " ha intentado crear un evento. Baneenlo debe ser JONA ESQUIOGA NDEA.", FontTypeNames.FONTTYPE_ADMIN)
280                         End If
                      End If
290               Case EventPacketID.eCloseEvent
300                   EventosDS.CloseEvent buffer.ReadByte, , True
                      
310               Case EventPacketID.RequiredEvents
320                   If EsGm(userindex) Then
330                       WriteEventPacket userindex, SvEventPacketID.SendListEvent
340                   End If
                      
350               Case EventPacketID.RequiredDataEvent
360                   If EsGm(userindex) Then
370                       WriteEventPacket userindex, SvEventPacketID.SendDataEvent, CByte(buffer.ReadByte())
380                   Else
390                       buffer.ReadByte
400                   End If
                  
410               Case EventPacketID.eAbandonateEvent
420                   EventosDS.AbandonateEvent userindex, , True
                      
430               Case EventPacketID.eParticipeEvent
440                   If UserList(userindex).incomingData.length < 5 Then
450                       Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
460                       Exit Sub
470                   End If
                      
480                   EventosDS.ParticipeEvent userindex, buffer.ReadASCIIString
490           End Select

              'If we got here then packet is complete, copy data back to original queue
500           Call .incomingData.CopyBuffer(buffer)
510       End With

Errhandler:
          Dim error  As Long
520       error = Err.Number
530       On Error GoTo 0

          'Destroy auxiliar buffer
540       Set buffer = Nothing

550       If error <> 0 Then _
             Err.Raise error

   On Error GoTo 0
   Exit Sub

HandleEventPacket_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleEventPacket of Módulo Protocol in line " & Erl
End Sub


Public Sub WriteEventPacket(ByVal userindex As Integer, ByVal PacketID As Byte, Optional ByVal DataExtra As Long)
   On Error GoTo WriteEventPacket_Error

10        On Error GoTo Errhandler

          Dim LoopC As Integer
          
20        With UserList(userindex).outgoingData
30            Call .WriteByte(ServerPacketID.EventPacketSv)
40            Call .WriteByte(PacketID)
              
50            Select Case PacketID
                  Case SvEventPacketID.SendListEvent
60                    For LoopC = 1 To EventosDS.MAX_EVENT_SIMULTANEO
70                        Call .WriteByte(IIf((Events(LoopC).Enabled = True), Events(LoopC).Modality, 0))
                          
80                    Next LoopC
                      
90                Case SvEventPacketID.SendDataEvent
                      If DataExtra > 0 Then
100                       Call .WriteByte(Events(DataExtra).Inscribed)
110                       Call .WriteByte(Events(DataExtra).Quotas)
120                       Call .WriteByte(Events(DataExtra).LvlMin)
130                       Call .WriteByte(Events(DataExtra).LvlMax)
140                       Call .WriteLong(Events(DataExtra).GldInscription * Events(DataExtra).Quotas)
150                       Call .WriteLong(Events(DataExtra).DspInscription * Events(DataExtra).Quotas)
160                       Call .WriteASCIIString(strUsersEvent(DataExtra))
                        Else
                            LogError "El personaje " & UserList(userindex).Name & " envió DataExtra 0 en WriteEventPacket"
                        End If
170           End Select
              
180       End With
190   Exit Sub

Errhandler:
200       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
210           Call FlushBuffer(userindex)
220           Resume
230       End If

   On Error GoTo 0
   Exit Sub

WriteEventPacket_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteEventPacket of Módulo Protocol in line " & Erl
End Sub

Public Sub WriteUserInEvent(ByVal userindex As Integer)
10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.UserInEvent)
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If
End Sub

Private Sub HandleCofres(ByVal userindex As Integer)
10        With UserList(userindex)
20            Call .incomingData.ReadByte
              
              Dim TipoCofre As Byte
              Dim Obj As Obj
              Dim strTemp As String
30            TipoCofre = .incomingData.ReadByte
              
40            Obj.Amount = 1
              
              'Config Basica
50            Select Case TipoCofre
                  Case 0 ' BRONCE
60                    Obj.ObjIndex = 946
                      
70                    If .flags.Bronce = 1 Then
80                        WriteConsoleMsg userindex, "Ya eres usuario BRONCE", FontTypeNames.FONTTYPE_GUILD
90                        Exit Sub
100                   End If
                      
110                   If Not TieneObjetos(Obj.ObjIndex, Obj.Amount, userindex) Then
120                       WriteConsoleMsg userindex, "Necesitas tener contigo el cofre de los inmortales [BRONCE]", FontTypeNames.FONTTYPE_WARNING
130                       Exit Sub
140                   End If
                      
150                   .flags.Bronce = 1
160               Case 1 'PLATA
170                   Obj.ObjIndex = 945
                      
180                   If .flags.Plata = 1 Then
190                       WriteConsoleMsg userindex, "Ya eres usuario PLATA", FontTypeNames.FONTTYPE_GUILD
200                       Exit Sub
210                   End If
                      
220                   If Not TieneObjetos(Obj.ObjIndex, Obj.Amount, userindex) Then
230                       WriteConsoleMsg userindex, "Necesitas tener contigo el cofre de los inmortales [PLATA]", FontTypeNames.FONTTYPE_WARNING
240                       Exit Sub
250                   End If
                      
260                   .flags.Plata = 1
270               Case 2 'ORO
280                   Obj.ObjIndex = 944
                      
290                   If .flags.Oro = 1 Then
300                       WriteConsoleMsg userindex, "Ya eres usuario ORO", FontTypeNames.FONTTYPE_GUILD
310                       Exit Sub
320                   End If
                      
330                   If Not TieneObjetos(Obj.ObjIndex, Obj.Amount, userindex) Then
340                       WriteConsoleMsg userindex, "Necesitas tener contigo el cofre de los inmortales [ORO]", FontTypeNames.FONTTYPE_WARNING
350                       Exit Sub
360                   End If
                      
370                   .flags.Oro = 1
380               Case 3 'PREMIUM
390                   Obj.ObjIndex = 1115
                      
400                   If .flags.Premium = 1 Then
410                       WriteConsoleMsg userindex, "Ya eres usuario PREMIUM", FontTypeNames.FONTTYPE_GUILD
420                       Exit Sub
430                   End If
                      
440                   If Not TieneObjetos(Obj.ObjIndex, Obj.Amount, userindex) Then
450                       WriteConsoleMsg userindex, "Necesitas tener contigo el cofre de los inmortales [PREMIUM]", FontTypeNames.FONTTYPE_WARNING
460                       Exit Sub
470                   End If
                      
480                   .flags.Premium = 1
                      
490               Case 4
                      'DIOS
500                   Obj.ObjIndex = 1115
                      
510                   If .flags.Premium = 1 Then
520                       WriteConsoleMsg userindex, "Ya eres usuario DIOS", FontTypeNames.FONTTYPE_GUILD
530                       Exit Sub
540                   End If
                      
550                   If Not TieneObjetos(Obj.ObjIndex, Obj.Amount, userindex) Then
560                       WriteConsoleMsg userindex, "Necesitas tener contigo el cofre de los inmortales [DIOS]", FontTypeNames.FONTTYPE_WARNING
570                       Exit Sub
580                   End If
                      
590                   .flags.Premium = 1
600           End Select
              
              
610       End With
End Sub

Private Sub HandleComandoPorDias(ByVal userindex As Integer)
      '***************************************************
      'Author: Lautaro
      'Last Modification: -
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 6 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              'Reads the UserName and Slot Packets
              Dim UserName As String
              Dim tipo As Byte
              Dim tIndex As Integer
              Dim strDate As String
              
90            tipo = buffer.ReadByte()    'Que Larry?
100           UserName = buffer.ReadASCIIString()    'Que UserName?
110           strDate = buffer.ReadASCIIString()
              
              
              If Protocol.IsNickEspecial(.Name) Then
130               Select Case tipo
                      Case 0 ' Ban por días
140                       If Not FileExist(App.Path & "\CHARFILE\" & UserName & ".CHR", vbNormal) Then
150                           WriteConsoleMsg userindex, "El personaje no existe.", FontTypeNames.FONTTYPE_INFO
160                       Else
170                           mDias.BanUserDias userindex, UserName, strDate
180                       End If
190                   Case 1 ' Convertir en dioses
200                       tIndex = NameIndex(UserName)  'Que user index?
210                       If tIndex > 0 Then
220                           mDias.TransformarUserDios userindex, tIndex, strDate
230                       Else
240                           WriteConsoleMsg userindex, "El personaje está offline.", FontTypeNames.FONTTYPE_INFO
250                       End If
                              
260                   End Select
270           End If

              'If we got here then packet is complete, copy data back to original queue
280           Call .incomingData.CopyBuffer(buffer)
290       End With

Errhandler:
          Dim error  As Long
300       error = Err.Number
310       On Error GoTo 0

          'Destroy auxiliar buffer
320       Set buffer = Nothing

330       If error <> 0 Then _
             Err.Raise error
End Sub

Private Sub HandleReportCheat(ByVal userindex As Integer)
      '***************************************************
      'Author: Lautaro
      'Last Modification: -
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              'Reads the UserName and Slot Packets
              Dim UserName As String
              Dim DataName As String
              
90            UserName = buffer.ReadASCIIString()
100           DataName = buffer.ReadASCIIString()
              
110           'SendData SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El personaje " & UserName & " tiene un programa APARENTEMENTE peligroso abierto: " & DataName, FontTypeNames.FONTTYPE_ADMIN)
120           LogAntiCheat "Se ha detectado que el personaje " & UserName & " tiene un posible programa prohibido llamado " & DataName
              

              'If we got here then packet is complete, copy data back to original queue
130           Call .incomingData.CopyBuffer(buffer)
140       End With

Errhandler:
          Dim error  As Long
150       error = Err.Number
160       On Error GoTo 0

          'Destroy auxiliar buffer
170       Set buffer = Nothing

180       If error <> 0 Then _
             Err.Raise error
End Sub

Private Sub HandleDisolutionGuild(ByVal userindex As Integer)

   On Error GoTo HandleDisolutionGuild_Error

10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte
              
              Dim tipo As Byte
              
90            tipo = buffer.ReadByte()
              
100           Select Case tipo
                  Case 0
110                   modGuilds.DisolverGuildIndex userindex
120               Case 1
130                   modGuilds.ReanudarGuildIndex userindex, buffer.ReadASCIIString
140           End Select

              'If we got here then packet is complete, copy data back to original queue
150           Call .incomingData.CopyBuffer(buffer)
160       End With

Errhandler:
          Dim error  As Long
170       error = Err.Number
180       On Error GoTo 0

          'Destroy auxiliar buffer
190       Set buffer = Nothing

200       If error <> 0 Then _
             Err.Raise error

   On Error GoTo 0
   Exit Sub

HandleDisolutionGuild_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleDisolutionGuild of Módulo Protocol in line " & Erl
End Sub

Public Sub WriteShortMsj(ByVal UIndex As Integer, _
                         ByVal MsgShort As Integer, _
                         ByVal FontType As FontTypeNames, _
                         Optional ByVal tmpInteger1 As Integer = 0, _
                         Optional ByVal tmpInteger2 As Integer = 0, _
                         Optional ByVal tmpInteger3 As Integer = 0, _
                         Optional ByVal tmpLong As Long = 0, _
                         Optional ByVal tmpStr As String = vbNullString)
                               
   ' VER LAUTARO
   On Error GoTo WriteShortMsj_Error

10        With UserList(UIndex).outgoingData
20            .WriteByte ServerPacketID.ShortMsj
30            .WriteInteger MsgShort
40            .WriteByte FontType

50            If tmpInteger1 <> 0 Then .WriteInteger tmpInteger1
60            If tmpInteger2 <> 0 Then .WriteInteger tmpInteger2
70            If tmpInteger3 <> 0 Then .WriteInteger tmpInteger3
80            If tmpLong <> 0 Then .WriteLong tmpLong
90            If Len(tmpStr) <> 0 Then .WriteASCIIString tmpStr

100       End With

   On Error GoTo 0
   Exit Sub

WriteShortMsj_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteShortMsj of Módulo Protocol in line " & Erl

End Sub

Public Function PrepareMessageShortMsj(ByVal MsgShort As Integer, _
                                       ByVal FontType As FontTypeNames, _
                                       Optional ByVal tmpInteger1 As Integer = 0, _
                                       Optional ByVal tmpInteger2 As Integer = 0, _
                                       Optional ByVal tmpInteger3 As Integer = 0, _
                                       Optional ByVal tmpLong As Long = 0, _
                                       Optional ByVal tmpStr As String = vbNullString) As String

   On Error GoTo PrepareMessageShortMsj_Error

10        With auxiliarBuffer
20            .WriteByte ServerPacketID.ShortMsj
30            .WriteInteger MsgShort
40            .WriteByte FontType
       
50            If tmpInteger1 <> 0 Then .WriteInteger tmpInteger1
60            If tmpInteger2 <> 0 Then .WriteInteger tmpInteger2
70            If tmpInteger3 <> 0 Then .WriteInteger tmpInteger3
80            If tmpLong <> 0 Then .WriteLong tmpLong
90            If Len(tmpStr) <> 0 Then .WriteASCIIString tmpStr
              
100           PrepareMessageShortMsj = .ReadASCIIStringFixed(.length)

110       End With

   On Error GoTo 0
   Exit Function

PrepareMessageShortMsj_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessageShortMsj of Módulo Protocol in line " & Erl

End Function


Public Function PrepareMessagePalabrasMagicas(ByVal CharIndex As Integer, ByVal SpellIndex As Byte, ByVal color As Long) As String
   On Error GoTo PrepareMessagePalabrasMagicas_Error

10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.PalabrasMagicas)
30            Call .WriteInteger(CharIndex)
40            Call .WriteByte(SpellIndex)

              ' Write rgb channels and save one byte from long :D
50            Call .WriteByte(color And &HFF)
60            Call .WriteByte((color And &HFF00&) \ &H100&)
70            Call .WriteByte((color And &HFF0000) \ &H10000)

80            PrepareMessagePalabrasMagicas = .ReadASCIIStringFixed(.length)
90        End With

   On Error GoTo 0
   Exit Function

PrepareMessagePalabrasMagicas_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PrepareMessagePalabrasMagicas of Módulo Protocol in line " & Erl
End Function
Public Function PrepareMessageDescNpcs(ByVal CharIndex As Integer, ByVal NumeroNpc As Integer) As String
10        With auxiliarBuffer
20            Call .WriteByte(ServerPacketID.DescNpcs)
30            Call .WriteInteger(CharIndex)
40            Call .WriteInteger(NumeroNpc)

50            PrepareMessageDescNpcs = .ReadASCIIStringFixed(.length)
60        End With
End Function

Public Sub WriteDescNpcs(ByVal userindex As Integer, ByVal CharIndex As Integer, ByVal NumeroNpc As Integer)
10        On Error GoTo Errhandler
20        Call UserList(userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageDescNpcs(CharIndex, NumeroNpc))
30        Exit Sub

Errhandler:
40        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
50            Call FlushBuffer(userindex)
60            Resume
70        End If
End Sub

Private Sub HandleChangeNick(ByVal userindex As Integer)
   On Error GoTo HandleChangeNick_Error

10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
          
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String, tmpStr As String, p As String

90            UserName = buffer.ReadASCIIString()
              
100           Call General.ChangeNick(userindex, UserName)
              
              'If we got here then packet is complete, copy data back to original queue
110           Call .incomingData.CopyBuffer(buffer)
120       End With

Errhandler:
          Dim error  As Long
130       error = Err.Number
140       On Error GoTo 0

          'Destroy auxiliar buffer
150       Set buffer = Nothing

160       If error <> 0 Then _
             Err.Raise error

   On Error GoTo 0
   Exit Sub

HandleChangeNick_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleChangeNick of Módulo Protocol in line " & Erl
End Sub

Private Function SearchItemCanje(ByVal CanjeItem As Integer, ByVal ObjRequired1 As Integer, ByVal Points As Integer) As Byte

          Dim LoopC As Integer
          Dim ArrayValue As Long
          
   On Error GoTo SearchItemCanje_Error

10        For LoopC = 1 To NumCanjes
20            With Canjes(LoopC)
30                If CanjeItem = .ObjCanje.ObjIndex Then
40                    GetSafeArrayPointer Canjes(LoopC).ObjRequired, ArrayValue
                      
50                    If ArrayValue <> 0 Then
60                        If Canjes(LoopC).ObjRequired(1).ObjIndex = ObjRequired1 And Points = Canjes(LoopC).Points Then
70                            SearchItemCanje = LoopC
80                            Exit For
90                        End If
100                   Else
110                       If Canjes(LoopC).Points = Points Then
120                           SearchItemCanje = LoopC
130                           Exit For
140                       End If
150                   End If
160               End If
170           End With
180       Next LoopC

   On Error GoTo 0
   Exit Function

SearchItemCanje_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure SearchItemCanje of Módulo Protocol in line " & Erl
End Function

Public Sub HandleCanjeItem(ByVal userindex As Integer)
   On Error GoTo HandleCanjeItem_Error

10        If UserList(userindex).incomingData.length < 7 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
                      
50        With UserList(userindex)
60            Call .incomingData.ReadByte
              
70            If .flags.Muerto Then Exit Sub
              
              Dim CanjeItem As Integer
              Dim CanjeIndex As Byte
              
80            CanjeItem = .incomingData.ReadInteger
90            CanjeIndex = SearchItemCanje(CanjeItem, .incomingData.ReadInteger, .incomingData.ReadInteger)
              
100           If CanjeIndex = 0 Then Exit Sub
              
110           Call General.CanjearObjeto(userindex, CanjeIndex)
              
120       End With

   On Error GoTo 0
   Exit Sub

HandleCanjeItem_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleCanjeItem of Módulo Protocol in line " & Erl
End Sub

Public Sub HandleCanjeInfo(ByVal userindex As Integer)

   On Error GoTo HandleCanjeInfo_Error

10        If UserList(userindex).incomingData.length < 7 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50        With UserList(userindex)
60            Call .incomingData.ReadByte

              Dim CanjeItem As Integer
              Dim CanjeIndex As Byte
              
70            CanjeItem = .incomingData.ReadInteger
80            CanjeIndex = SearchItemCanje(CanjeItem, .incomingData.ReadInteger, .incomingData.ReadInteger)
              
90            If CanjeIndex = 0 Then Exit Sub
100           If .flags.Muerto Then Exit Sub
              
110           Call WriteCanjeInfo(userindex, CanjeIndex)
          
120       End With

   On Error GoTo 0
   Exit Sub

HandleCanjeInfo_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleCanjeInfo of Módulo Protocol in line " & Erl
End Sub
Public Sub WriteCanjeInfo(ByVal userindex As Integer, ByVal CanjeIndex As Byte)
   On Error GoTo WriteCanjeInfo_Error

10        On Error GoTo Errhandler
          Dim LoopC As Integer
          Dim LoopY As Integer
              
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.InfoCanje)
          
          
          If Canjes(CanjeIndex).ObjCanje.ObjIndex > 0 Then
30            With ObjData(Canjes(CanjeIndex).ObjCanje.ObjIndex)
40                Call UserList(userindex).outgoingData.WriteInteger(.MinDef)
50                Call UserList(userindex).outgoingData.WriteInteger(.MaxDef)
60                Call UserList(userindex).outgoingData.WriteInteger(.DefensaMagicaMin)
70                Call UserList(userindex).outgoingData.WriteInteger(.DefensaMagicaMax)
80                Call UserList(userindex).outgoingData.WriteInteger(.MinHIT)
90                Call UserList(userindex).outgoingData.WriteInteger(.MaxHIT)
100               Call UserList(userindex).outgoingData.WriteLong(Canjes(CanjeIndex).Points)
                  
110               If .ObjType = otMonturas Or .ObjType = otMonturasDraco Then
120                   Call UserList(userindex).outgoingData.WriteByte(1)
130               Else
140                   Call UserList(userindex).outgoingData.WriteByte(.NoSeCae)
150               End If
                  
                  
160           End With
            End If
170       Exit Sub

Errhandler:
180       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
190           Call FlushBuffer(userindex)
200           Resume
210       End If

   On Error GoTo 0
   Exit Sub

WriteCanjeInfo_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteCanjeInfo of Módulo Protocol in line " & Erl
End Sub

Private Function SearchNpcCanje(ByVal CanjeIndex As Integer, ByVal NpcNumero As Integer) As Boolean
          Dim LoopC As Integer
          
          
   On Error GoTo SearchNpcCanje_Error

10        SearchNpcCanje = False
          
20        With Canjes(CanjeIndex)
30            If .Npcs = NpcNumero Then
40                SearchNpcCanje = True
50                Exit Function
60            End If
          
70        End With

   On Error GoTo 0
   Exit Function

SearchNpcCanje_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure SearchNpcCanje of Módulo Protocol in line " & Erl
End Function
Public Sub WriteCanjeInit(ByVal userindex As Integer, ByVal NpcNumero As Integer)
   On Error GoTo WriteCanjeInit_Error

10        On Error GoTo Errhandler
              Dim LoopC As Integer
              Dim LoopY As Integer
              Dim NpcIndex As Integer
              Dim SearchNpc As Boolean
              Dim Num As Integer
              
              
20
          
30        For LoopC = 1 To NumCanjes
40            SearchNpc = SearchNpcCanje(LoopC, NpcNumero)
              
50            If SearchNpc Then
60                Num = Num + 1
70            End If
80        Next LoopC

            If Num = 0 Then Exit Sub
          
          Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.CanjeInit)
90        Call UserList(userindex).outgoingData.WriteByte(Num)
          
100
          
110       For LoopC = 1 To NumCanjes
120           With Canjes(LoopC)
130               If SearchNpcCanje(LoopC, NpcNumero) Then
                  
140                   Call UserList(userindex).outgoingData.WriteByte(.NumRequired)
                      
150                   For LoopY = 1 To .NumRequired
160                        Call UserList(userindex).outgoingData.WriteInteger(.ObjRequired(LoopY).ObjIndex)
170                        Call UserList(userindex).outgoingData.WriteInteger(.ObjRequired(LoopY).Amount)
180                   Next LoopY
                      
190                   Call UserList(userindex).outgoingData.WriteInteger(.ObjCanje.ObjIndex)
200                   Call UserList(userindex).outgoingData.WriteInteger(.ObjCanje.Amount)
210                   Call UserList(userindex).outgoingData.WriteInteger(ObjData(.ObjCanje.ObjIndex).GrhIndex)
220                   Call UserList(userindex).outgoingData.WriteInteger(.Points)
230               End If
240           End With
              
250       Next LoopC
260       Exit Sub

Errhandler:
270       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
280           Call FlushBuffer(userindex)
290           Resume
300       End If

   On Error GoTo 0
   Exit Sub

WriteCanjeInit_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteCanjeInit of Módulo Protocol in line " & Erl
End Sub

Public Sub HandlePacketRetos(ByVal userindex As Integer)
   On Error GoTo HandlePacketRetos_Error

10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
          
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              Dim UserName As String, tmpStr As String, p As String
              Dim tipo As Byte, SubTipo As Byte
              Dim GldRequired As Long, DspRequired As Long, LimiteRojas As Integer
              Dim Users() As String, Team(10) As Byte
              Dim LoopC As Integer
              
90            tipo = buffer.ReadByte
              
              
              'LogAntiCheat "El personaje " & .Name & " uso el paquete de retos. SubPacket: " & Tipo
              
100           Select Case tipo
                  Case 0 ' Enviar solicitud
          
110                   GldRequired = buffer.ReadLong
120                   DspRequired = buffer.ReadLong
130                   LimiteRojas = buffer.ReadInteger
140                   UserName = buffer.ReadASCIIString
150                   UserName = UserName & "-" & .Name
                      
160                   Users = Split(UserName, "-")
                      
170                   If Not RetosActivos Then
180                       WriteConsoleMsg userindex, "Los retos están desactivados.", FontTypeNames.FONTTYPE_INFO
190                   Else
200                       Call mRetos.SendFight(userindex, eTipoReto.FightOne, GldRequired, DspRequired, LimiteRojas, Users)
210                   End If
220               Case 1 ' Aceptar solicitud
                      
230                   UserName = buffer.ReadASCIIString
                      
240                   If Not RetosActivos Then
250                       WriteConsoleMsg userindex, "Los retos están desactivados.", FontTypeNames.FONTTYPE_INFO
260                   Else
270                       Call mRetos.AcceptFight(userindex, UserName)
280                   End If
                      
290               Case 2 ' Salir del evento
300                   If .flags.SlotReto > 0 Then
310                       Call mRetos.UserdieFight(userindex, 0, True)
320                   End If
                      
330               Case 3 'Enviar Clan vs Clan
                      
340                   UserName = buffer.ReadASCIIString
                      
350                   'mCVC.SendFightGuild UserIndex, NameIndex(UserName)
                      
360               Case 4 'Aceptar Clan vs Clan
370                   UserName = buffer.ReadASCIIString
380                   'mCVC.AcceptFightGuild UserIndex, NameIndex(UserName)
                  
390               Case 5 'Requerimos el panel de retos
400                   WriteSendRetos userindex
                      
410           End Select
              
              'If we got here then packet is complete, copy data back to original queue
420           Call .incomingData.CopyBuffer(buffer)
430       End With

Errhandler:
          Dim error  As Long
440       error = Err.Number
450       On Error GoTo 0

          'Destroy auxiliar buffer
460       Set buffer = Nothing

470       If error <> 0 Then _
             Err.Raise error

   On Error GoTo 0
   Exit Sub

HandlePacketRetos_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HandlePacketRetos of Módulo Protocol in line " & Erl
End Sub
Public Sub WriteSendRetos(ByVal userindex As Integer)
          Dim strTemp As String
          
   On Error GoTo WriteSendRetos_Error

10        On Error GoTo Errhandler
          
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.SendRetos)
30            strTemp = Ranking(eRanking.TopRetos).Nombre(1) & "-" & Ranking(eRanking.TopRetos).Nombre(2) & "-" & Ranking(eRanking.TopRetos).Nombre(3)
          
40            Call UserList(userindex).outgoingData.WriteASCIIString(strTemp)
50        Exit Sub

Errhandler:
60        If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
70            Call FlushBuffer(userindex)
80            Resume
90        End If

   On Error GoTo 0
   Exit Sub

WriteSendRetos_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteSendRetos of Módulo Protocol in line " & Erl
End Sub

Private Function strListApostadores() As String

          Dim LoopC As Integer
          
   On Error GoTo strListApostadores_Error

10        For LoopC = LBound(GambleSystem.Users()) To UBound(GambleSystem.Users())
20            If GambleSystem.Users(LoopC).Name <> vbNullString Then
30                strListApostadores = strListApostadores & GambleSystem.Users(LoopC).Name & "-"
40            End If
          
50        Next LoopC
          
          
60        If Len(strListApostadores) > 0 Then
70            strListApostadores = mid$(strListApostadores, 1, Len(strListApostadores) - 1)
80        End If

   On Error GoTo 0
   Exit Function

strListApostadores_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure strListApostadores of Módulo Protocol in line " & Erl
End Function

Private Function strListApuestas() As String
          Dim LoopC As Integer
          
   On Error GoTo strListApuestas_Error

10        For LoopC = LBound(GambleSystem.Apuestas()) To UBound(GambleSystem.Apuestas())
20            If GambleSystem.Apuestas(LoopC) <> vbNullString Then
30                strListApuestas = strListApuestas & GambleSystem.Apuestas(LoopC) & ","
40            End If
          
50        Next LoopC
          
60        If Len(strListApuestas) > 0 Then
70            strListApuestas = mid$(strListApuestas, 1, Len(strListApuestas) - 1)
80        End If

   On Error GoTo 0
   Exit Function

strListApuestas_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure strListApuestas of Módulo Protocol in line " & Erl
End Function

Public Sub WritePacketGambleSv(ByVal userindex As Integer, ByVal tipo As Byte)
          Dim strTemp As String
          Dim strList As String
          
   On Error GoTo WritePacketGambleSv_Error

10        On Error GoTo Errhandler
          
20        Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.PacketGambleSv)
30        Call UserList(userindex).outgoingData.WriteByte(tipo)
          
40        Select Case tipo
              Case 0 'Enviamos la lista de usuarios que apostaron
                strList = strListApostadores
50                Call UserList(userindex).outgoingData.WriteASCIIString(strList)
60            Case 1 ' Enviamos la info de los usuarios que apostaron
                  
70            Case 2 ' Enviamos la lista de apuestas disponibles para los usuarios
                    strList = strListApuestas
80                Call UserList(userindex).outgoingData.WriteASCIIString(strList)
90        End Select
          
100       Exit Sub

Errhandler:
110       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
120           Call FlushBuffer(userindex)
130           Resume
140       End If

   On Error GoTo 0
   Exit Sub

WritePacketGambleSv_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WritePacketGambleSv of Módulo Protocol in line " & Erl
End Sub

Public Sub HandleUseItemPacket(ByVal userindex As Integer)
   On Error GoTo HandleUseItemPacket_Error

10        If UserList(userindex).incomingData.length < 7 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
          Dim Slot As Byte
          Dim SecondaryClick As Byte
          Dim key As Long
          Dim nclave As Byte
50        With UserList(userindex)
60            Call .incomingData.ReadByte
              
70            Slot = .incomingData.ReadByte
80            SecondaryClick = .incomingData.ReadByte
              key = .incomingData.ReadLong
                nclave = .incomingData.ReadByte
                If .ClaveAC <> nclave Then Exit Sub
                
              If (timeGetTime - UserList(userindex).Intervalos(SecondaryClick).ModifyTimer) <= 150 Then
                  Exit Sub
              End If
              
              If Not mSecurity.CheckKeyPacket(userindex, Key_UseItem, key) Then Exit Sub
              
90            Call UsarItem(userindex, Slot, SecondaryClick, key)
              UserList(userindex).Intervalos(SecondaryClick).ModifyTimer = timeGetTime
              'Call mIntervalos.AssignedInterval(UserIndex, SecondaryClick)

              
100       End With

   On Error GoTo 0
   Exit Sub

HandleUseItemPacket_Error:
    
    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleUseItemPacket of Módulo Protocol in line " & Erl & " Userindex: " & userindex & ", Slot: " & Slot & ", SecondatyClick: " & SecondaryClick
End Sub

Public Sub UsarItem(ByVal userindex As Integer, ByVal Slot As Byte, ByVal SecondaryClick As Byte, ByVal key As Long)
   On Error GoTo UsarItem_Error

10        With UserList(userindex)


90            If Slot <= .CurrentInventorySlots And Slot > 0 Then
100               If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
110           End If

120           If .flags.Meditando Then
130               Exit Sub    'The error message should have been provided by the client.
140           End If
              
              ' Click Derecho ejecutandose desde hechizos?
              If SecondaryClick And .flags.MenuCliente = 1 Then Exit Sub
              
20            If .flags.LastSlotClient <> 255 Then
30                If Slot <> .flags.LastSlotClient Then

60                    Exit Sub
70                End If
80            End If

150           If ObjData(.Invent.Object(Slot).ObjIndex).ObjType = otPociones Then
160               Call UseInvPotion(userindex, Slot, SecondaryClick)
              
170           Else
180               Call UseInvItem(userindex, Slot)
190           End If
              
200           Call WriteUpdateFollow(userindex)
210       End With

   On Error GoTo 0
   Exit Sub

UsarItem_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure UsarItem of Módulo Protocol in line " & Erl
End Sub

Private Sub HandleDarPoints(ByVal userindex As Integer)
      '***************************************************
      'Author: Lautaro
      'Last Modification: -
      '
      '***************************************************
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              'Reads the UserName and Slot Packets
              Dim UserName As String
              Dim Amount As Integer
              Dim tUser As Integer
              
90            Amount = buffer.ReadInteger()    'Que Larry?
100           UserName = buffer.ReadASCIIString()    'Que UserName?
              
              
              If Protocol.IsNickEspecial(.Name) Then
120               tUser = NameIndex(UserName)
                  
130               If tUser > 0 Then
140                   UserList(tUser).Stats.Points = UserList(tUser).Stats.Points + Amount
150                   WriteConsoleMsg userindex, "Le has dado " & Amount & " puntos de canje a " & UserName & ".", FontTypeNames.FONTTYPE_INFO
160                   WriteConsoleMsg tUser, "Has recibido " & Amount & " puntos de canje.", FontTypeNames.FONTTYPE_INFO
170                   'CheckRankingUser tUser, TopTorneos
                      WriteUpdatePoints tUser
180               Else
190                   WriteConsoleMsg userindex, "Personaje offline.", FontTypeNames.FONTTYPE_INFO
200               End If
210           End If

              'If we got here then packet is complete, copy data back to original queue
220           Call .incomingData.CopyBuffer(buffer)
230       End With

Errhandler:
          Dim error  As Long
240       error = Err.Number
250       On Error GoTo 0

          'Destroy auxiliar buffer
260       Set buffer = Nothing

270       If error <> 0 Then _
             Err.Raise error
End Sub

Public Sub HandleRequestInfoEvento(ByVal userindex As Integer)

          Dim strTemp As String
          
   On Error GoTo HandleRequestInfoEvento_Error

10        With UserList(userindex)
20            Call .incomingData.ReadByte
              
30            strTemp = SetInfoEvento & vbCrLf & GenerateInfoInvasion
              
40            If Len(strTemp) < 3 Then
50                WriteConsoleMsg userindex, "No hay eventos en curso. Danos tu sugerencia mediante /DENUNCIAR.", FontTypeNames.FONTTYPE_INFO
60            Else
70                WriteConsoleMsg userindex, strTemp, FontTypeNames.FONTTYPE_INFO
80            End If
          
90        End With

   On Error GoTo 0
   Exit Sub

HandleRequestInfoEvento_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleRequestInfoEvento of Módulo Protocol in line " & Erl
End Sub

Private Sub HandlePacketGamble(ByVal userindex As Integer)
      '***************************************************
      'Author: Lautaro
      'Last Modification: -
      '
      '***************************************************
   On Error GoTo HandlePacketGamble_Error

10        If UserList(userindex).incomingData.length < 2 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
              Dim buffer As New clsByteQueue
70            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
80            Call buffer.ReadByte

              'Reads the UserName and Slot Packets
              Dim UserName As String
              Dim tipo As Byte
              Dim tUser As Integer
              Dim Apuestas() As String
              
90            tipo = buffer.ReadByte()

              'Call LogAntiCheat("El personaje " & .Name & " uso el paquete Gamble: PACKETID: " & Tipo & ".")
              
100           Select Case tipo
                  Case 0 ' Gm crea nueva apuesta
110                   Apuestas = Split(buffer.ReadASCIIString, ",")
120                   mApuestas.NewGamble userindex, buffer.ReadASCIIString, buffer.ReadInteger, buffer.ReadByte, Apuestas
130               Case 1 ' Gm cancela la apuesta
140                   mApuestas.CancelGamble userindex
150               Case 2 ' Gm otorga premio de la apuesta
                      
160               Case 3 ' Personaje apuesta
170                   mApuestas.UserGamble userindex, buffer.ReadByte, buffer.ReadLong, buffer.ReadLong
                      
180               Case 4 ' Gm requiere la lista de usuarios apostando
190                   WritePacketGambleSv userindex, 0
200               Case 5 ' Info de los users de arriba
210                   buffer.ReadASCIIString
                      
220                   WritePacketGambleSv userindex, 1
230               Case 6 ' Lista de apuestas disponibles
240                   If GambleSystem.Run Then
250                       WritePacketGambleSv userindex, 2
260                   End If
                      
270               Case 7
280                   If Not EsGm(userindex) Then
290                       buffer.ReadASCIIString
300                   Else
310                       UserGambleWin userindex, buffer.ReadASCIIString
320                   End If
330               Case Else
360           End Select

              'If we got here then packet is complete, copy data back to original queue
370           Call .incomingData.CopyBuffer(buffer)
380       End With

Errhandler:
          Dim error  As Long
390       error = Err.Number
400       On Error GoTo 0

          'Destroy auxiliar buffer
410       Set buffer = Nothing

420       If error <> 0 Then _
             Err.Raise error

   On Error GoTo 0
   Exit Sub

HandlePacketGamble_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HandlePacketGamble of Módulo Protocol in line " & Erl
End Sub



' PAQUETES DEL MERCADO

Public Sub HandleRequestMercado(ByVal userindex As Integer)
    ' Enviamos el mercado al cliente
    With UserList(userindex)
        Call .incomingData.ReadByte
        
        Call WriteSendMercado(userindex)
        
    End With
End Sub

Public Sub HandleRequestOffer(ByVal userindex As Integer)

    ' Enviamos las ofertas del mercado
    With UserList(userindex)
        Call .incomingData.ReadByte
        
        'Call WriteSendOffer(UserIndex)
    End With
End Sub

Public Sub HandleRequestOfferSent(ByVal userindex As Integer)
10        If UserList(userindex).incomingData.length < 4 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50     On Error GoTo Errhandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(UserList(userindex).incomingData)
       
          'Remove packet ID
70        Call buffer.ReadByte
          
          Dim UserName As String
          Dim SlotMao As Byte
          Dim Temp As String
          
          UserName = buffer.ReadASCIIString
          SlotMao = buffer.ReadByte
          
          If ExistAccount(MaoList(SlotMao).Account) Then
            Temp = LoadDataAccount(MaoList(SlotMao).Account, "MAO", "USERS")
          
              If InStrB(Temp, UserName) <> 0 Then
                    WriteInfoPj userindex, UserName
              End If
          End If
          
          
          'If we got here then packet is complete, copy data back to original queue
280       Call UserList(userindex).incomingData.CopyBuffer(buffer)
          
Errhandler:
          Dim error  As Long
290       error = Err.Number
300       On Error GoTo 0
       
          'Destroy auxiliar buffer
310       Set buffer = Nothing
       
320       If error <> 0 Then _
             Err.Raise error
End Sub

Public Sub HandleReceiveOfferAccount(ByVal userindex As Integer)
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50     On Error GoTo Errhandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(UserList(userindex).incomingData)
       
          'Remove packet ID
70        Call buffer.ReadByte
          
          SuperMAO.Mao_SendOffer userindex, buffer.ReadByte, buffer.ReadASCIIString
          
          'If we got here then packet is complete, copy data back to original queue
280       Call UserList(userindex).incomingData.CopyBuffer(buffer)
          
Errhandler:
          Dim error  As Long
290       error = Err.Number
300       On Error GoTo 0
       
          'Destroy auxiliar buffer
310       Set buffer = Nothing
       
320       If error <> 0 Then _
             Err.Raise error
End Sub

Private Sub HandleRequestInfoCharMAO(ByVal userindex As Integer)

    Dim SlotMao As Byte
    Dim Tittle As String
    Dim Gld As Long
    Dim Dsp As Long
    Dim Bloqued As Byte
    Dim Users As String
    
    ' Requerimos información de la lista publicada..
    With UserList(userindex)
        Call .incomingData.ReadByte
        
        SlotMao = .incomingData.ReadByte
        
        If MaoList(SlotMao).Tittle = vbNullString Then
            WriteConsoleMsg userindex, "La publicación que has seleccionado parece que fue eliminada.", FontTypeNames.FONTTYPE_INFO
            LogMao "Un personaje solicito información de una publicación eliminada."
            Exit Sub
        End If
        
        If ExistAccount(MaoList(SlotMao).Account) Then
            If val(LoadDataAccount(MaoList(SlotMao).Account, "MAO", "ACTIVE")) > 0 Then
                Gld = val(LoadDataAccount(MaoList(SlotMao).Account, "MAO", "GLD"))
                Dsp = val(LoadDataAccount(MaoList(SlotMao).Account, "MAO", "DSP"))
                Bloqued = val(LoadDataAccount(MaoList(SlotMao).Account, "MAO", "BLOQUED"))
                Users = LoadDataAccount(MaoList(SlotMao).Account, "MAO", "USERS")
            
                WriteSendInfoMao userindex, MaoList(SlotMao).Tittle, Users, Gld, Dsp, Bloqued
            Else
                LogMao "Se solicito ver una publicación no activa."
            End If
        Else
            LogMao "La cuenta de la publicación seleccionada no dió correcta."
        End If
    End With
End Sub
Private Sub HandlePublicationPj(ByVal userindex As Integer)
10        If UserList(userindex).incomingData.length < 14 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50     On Error GoTo Errhandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(UserList(userindex).incomingData)
       
          'Remove packet ID
70        Call buffer.ReadByte
          
          SuperMAO.Mao_AddList userindex, buffer.ReadLong, buffer.ReadLong, buffer.ReadASCIIString, buffer.ReadASCIIString, buffer.ReadByte
          
          'If we got here then packet is complete, copy data back to original queue
310       Call UserList(userindex).incomingData.CopyBuffer(buffer)
          
Errhandler:
          Dim error  As Long
320       error = Err.Number
330       On Error GoTo 0
       
          'Destroy auxiliar buffer
340       Set buffer = Nothing
       
350       If error <> 0 Then _
             Err.Raise error
End Sub

Private Sub HandleInvitationChange(ByVal userindex As Integer)
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50     On Error GoTo Errhandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(UserList(userindex).incomingData)
       
          'Remove packet ID
70        Call buffer.ReadByte

          Dim UserPin As String
          Dim Name As String
          
          UserPin = buffer.ReadASCIIString
          Name = buffer.ReadASCIIString
                
                
                
                WriteConsoleMsg userindex, "Deshabilitado hasta el día 17/03/2019 21HS", FontTypeNames.FONTTYPE_WARNING
                
           ' If UCase$(UserPin) <> UCase$(GetVar(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", "INIT", "PIN")) Then
330        '     WriteConsoleMsg UserIndex, "Pin incorrecto", FontTypeNames.FONTTYPE_INFO
340        ' Else
350        '     MAO.Send_Invitation UserIndex, Name
360        ' End If

          'If we got here then packet is complete, copy data back to original queue
280       Call UserList(userindex).incomingData.CopyBuffer(buffer)
          
Errhandler:
          Dim error  As Long
290       error = Err.Number
300       On Error GoTo 0
       
          'Destroy auxiliar buffer
310       Set buffer = Nothing
       
320       If error <> 0 Then _
             Err.Raise error
End Sub
Private Sub HandleAcceptInvitation(ByVal userindex As Integer)
10        If UserList(userindex).incomingData.length < 5 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50     On Error GoTo Errhandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(UserList(userindex).incomingData)
       
          'Remove packet ID
70        Call buffer.ReadByte
          
          
          Dim UserPin As String
          Dim SlotOffer As Byte
          
          UserPin = buffer.ReadASCIIString
          SlotOffer = buffer.ReadByte
                
        If UserPin <> UCase$(LoadDataAccount(UserList(userindex).Account, "ACCOUNT", "PIN")) Then
            WriteConsoleMsg userindex, "Clave PIN incorrecta.", FontTypeNames.FONTTYPE_INFO
        Else
            SuperMAO.Mao_AcceptOffer userindex, SlotOffer
        End If
        
          'If we got here then packet is complete, copy data back to original queue
280       Call UserList(userindex).incomingData.CopyBuffer(buffer)
          
Errhandler:
          Dim error  As Long
290       error = Err.Number
300       On Error GoTo 0
       
          'Destroy auxiliar buffer
310       Set buffer = Nothing
       
320       If error <> 0 Then _
             Err.Raise error
End Sub
Private Sub HandleRechaceInvitation(ByVal userindex As Integer)
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50     On Error GoTo Errhandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(UserList(userindex).incomingData)
       
          'Remove packet ID
70        Call buffer.ReadByte
          
          Dim Name As String
          
          Name = buffer.ReadASCIIString
          
          
          WriteConsoleMsg userindex, "Deshabilitado hasta el día 17/03/2019 21HS", FontTypeNames.FONTTYPE_WARNING
          
          'MAO.Rechace_Invitation UserIndex, Name

          'If we got here then packet is complete, copy data back to original queue
280       Call UserList(userindex).incomingData.CopyBuffer(buffer)
          
Errhandler:
          Dim error  As Long
290       error = Err.Number
300       On Error GoTo 0
       
          'Destroy auxiliar buffer
310       Set buffer = Nothing
       
320       If error <> 0 Then _
             Err.Raise error
End Sub

Private Sub HandleCancelInvitation(ByVal userindex As Integer)
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50     On Error GoTo Errhandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(UserList(userindex).incomingData)
       
          'Remove packet ID
70        Call buffer.ReadByte
          
          Dim Name As String
          
          Name = buffer.ReadASCIIString
          
          WriteConsoleMsg userindex, "Deshabilitado hasta el día 17/03/2019 21HS", FontTypeNames.FONTTYPE_WARNING
          
          'SuperMAO.Mao_EndList UserIndex
          'MAO.Cancel_Invitation UserIndex, Name

          'If we got here then packet is complete, copy data back to original queue
280       Call UserList(userindex).incomingData.CopyBuffer(buffer)
          
Errhandler:
          Dim error  As Long
290       error = Err.Number
300       On Error GoTo 0
       
          'Destroy auxiliar buffer
310       Set buffer = Nothing
       
320       If error <> 0 Then _
             Err.Raise error
End Sub
Private Sub HandleSeguroClan(ByVal userindex As Integer)
    With UserList(userindex)
        Call .incomingData.ReadByte ' packetid
        If .GuildIndex <> 0 Then
            .seguroclan = Not .seguroclan
            
        End If
    End With
End Sub
Private Sub HandleEnviarAviso(ByVal userindex As Integer)
    With UserList(userindex).incomingData
        Call .ReadByte ' packetid
        
        DetectoAnticheat userindex, .ReadByte
    End With
End Sub
Private Sub HandleBuyPj(ByVal userindex As Integer)
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50     On Error GoTo Errhandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(UserList(userindex).incomingData)
       
          'Remove packet ID
70        Call buffer.ReadByte
          
          Dim Name As String
          
          Name = buffer.ReadASCIIString
          
          WriteConsoleMsg userindex, "Deshabilitado hasta el día 17/03/2019 21HS", FontTypeNames.FONTTYPE_WARNING
          
          'MAO.Buy_Pj UserIndex, Name

          'If we got here then packet is complete, copy data back to original queue
280       Call UserList(userindex).incomingData.CopyBuffer(buffer)
          
Errhandler:
          Dim error  As Long
290       error = Err.Number
300       On Error GoTo 0
       
          'Destroy auxiliar buffer
310       Set buffer = Nothing
       
320       If error <> 0 Then _
             Err.Raise error
End Sub

Private Sub HandleQuitarPj(ByVal userindex As Integer)
    With UserList(userindex)
        Call .incomingData.ReadByte
        
        'MAO.Remove_Pj UserIndex
        WriteConsoleMsg userindex, "Deshabilitado hasta el día 17/03/2019 21HS", FontTypeNames.FONTTYPE_WARNING
    End With
End Sub
Public Function Chars_Mercado() As String
          Dim LoopC As Long
          
   On Error GoTo Chars_Mercado_Error

10        For LoopC = 1 To SuperMAO.MAX_MAO_LIST
20            If MaoList(LoopC).Tittle <> vbNullString Then
30                Chars_Mercado = Chars_Mercado & MaoList(LoopC).Tittle & SEPARATOR
              Else
                  Chars_Mercado = Chars_Mercado & "(Vacio)" & SEPARATOR
40            End If
50        Next LoopC

          Chars_Mercado = mid$(Chars_Mercado, 1, Len(Chars_Mercado) - 1)
   On Error GoTo 0
   Exit Function

Chars_Mercado_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Chars_Mercado of Módulo MAO in line " & Erl
          
End Function


Public Function Mao_List_OfferReceive(ByVal userindex As Integer) As String
          Dim LoopC As Long
          Dim Temp As String
          
   On Error GoTo Mao_List_OfferReceive_Error

10        With UserList(userindex)
                 Temp = LoadDataAccount(.Account, "MAO", "RECIBIDAS")
                 
                 Mao_List_OfferReceive = Replace(Temp, ".", SEPARATOR)
70        End With

   On Error GoTo 0
   Exit Function

Mao_List_OfferReceive_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Mao_List_OfferReceive of Módulo MAO in line " & Erl
End Function
Public Function Mao_List_OfferSend(ByVal userindex As Integer) As String
          Dim LoopC As Long
          Dim Temp As String
          
   On Error GoTo Mao_List_OfferSend_Error

10        With UserList(userindex)
                 Temp = LoadDataAccount(.Account, "MAO", "ENVIADAS")
                 
                 Mao_List_OfferSend = Replace(Temp, ".", SEPARATOR)
70        End With

   On Error GoTo 0
   Exit Function

Mao_List_OfferSend_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Mao_List_OfferSend of Módulo MAO in line " & Erl
End Function
Private Sub WriteSendMercado(ByVal userindex As Integer)

   On Error GoTo WriteSendMercado_Error

10        On Error GoTo Errhandler
20            Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.SendMercado)
50            Call UserList(userindex).outgoingData.WriteASCIIString(Chars_Mercado)
              Call UserList(userindex).outgoingData.WriteASCIIString(Mao_List_OfferSend(userindex))
              Call UserList(userindex).outgoingData.WriteASCIIString(Mao_List_OfferReceive(userindex))
110       Exit Sub


Errhandler:
120       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
130           Call FlushBuffer(userindex)
140           Resume
150       End If

   On Error GoTo 0
   Exit Sub

WriteSendMercado_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteSendMercado of Módulo Protocol in line " & Erl
End Sub

Private Sub WriteSendInfoMao(ByVal userindex As Integer, _
                                ByVal Tittle As String, _
                                ByVal Users As String, _
                                ByVal Gld As Long, _
                                ByVal Dsp As Long, _
                                ByVal Bloqued As Byte)
                                
   On Error GoTo WriteSendOffer_Error
10        On Error GoTo Errhandler
20            Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.SendInfoMAO)
              Call UserList(userindex).outgoingData.WriteASCIIString(Tittle)
              Call UserList(userindex).outgoingData.WriteASCIIString(Users)
              Call UserList(userindex).outgoingData.WriteLong(Gld)
              Call UserList(userindex).outgoingData.WriteLong(Dsp)
              Call UserList(userindex).outgoingData.WriteByte(Bloqued)
110       Exit Sub

Errhandler:
120       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
130           Call FlushBuffer(userindex)
140           Resume
150       End If

   On Error GoTo 0
   Exit Sub

WriteSendOffer_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteSendOffer of Módulo Protocol in line " & Erl
End Sub

Private Sub HandleSubirCanje(ByVal userindex As Integer)
    With UserList(userindex)
        Call .incomingData.ReadByte
        
        Dim TempAmount As Integer
        Dim TempObj As Obj
        TempObj.Amount = 1
        TempObj.ObjIndex = 1277
        
       
        If TieneObjetos(TempObj.ObjIndex, TempObj.Amount, userindex) Then
               WriteConsoleMsg userindex, "Has recibido " & TempObj.Amount & " puntos de canje. ", FontTypeNames.FONTTYPE_INFO
               
               Call QuitarObjetos(TempObj.ObjIndex, TempObj.Amount, userindex)
               .Stats.Points = .Stats.Points + TempObj.Amount
               WriteUpdatePoints userindex
               
        Else
            WriteConsoleMsg userindex, "No tienes canjes.", FontTypeNames.FONTTYPE_ADMIN
        End If
                
    End With
End Sub


Private Sub HandlePartyClient(ByVal userindex As Integer)
        ' 1) Requiere formulario 'principal'
        ' 2) Requiere formulario 'solicitudes'
        ' 3) Requiere formulario 'obtenido'
        ' 4) Abandonar party
        ' 5) Requiere ingresar a party
    Dim Paso As Byte
    
    With UserList(userindex)
        .incomingData.ReadByte
        
        Select Case .incomingData.ReadByte
            Case 1
                If .GroupIndex <= 0 Then
                    mGroup.CreateGroup (userindex)
                Else
                    WriteGroupPrincipal (userindex)
                End If
            Case 2
                If .GroupIndex > 0 Then
                    If Groups(.GroupIndex).User(1).Index = userindex Then
                        WriteGroupRequests (userindex)
                    End If
                End If
            Case 3
                If .GroupIndex > 0 Then
                    WriteGroupReward (userindex)
                End If
                
            Case 4
                If .GroupIndex > 0 Then
                    mGroup.AbandonateGroup userindex
                End If
            
            Case 5
                mGroup.SendInvitationGroup userindex
        End Select
    End With
End Sub

Public Sub WriteGroupPrincipal(ByVal userindex As Integer)

On Error GoTo Errhandler

    Dim GroupIndex As Integer
    Dim A As Long, B As Long
    
    Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.GroupPrincipal)
    
    GroupIndex = UserList(userindex).GroupIndex
    
    With Groups(GroupIndex)
        For A = 1 To MAX_MEMBERS_GROUP
            If .User(A).Index > 0 Then
                Call UserList(userindex).outgoingData.WriteASCIIString(UserList(.User(A).Index).Name)
            Else
                Call UserList(userindex).outgoingData.WriteASCIIString("<Vacio>")
            End If
            
            Call UserList(userindex).outgoingData.WriteByte(.User(A).PorcExp)
            Call UserList(userindex).outgoingData.WriteByte(.User(A).PorcGld)
        Next A
        
        
        
        If .Members = 1 Then
            Call UserList(userindex).outgoingData.WriteBoolean(False)
            Call UserList(userindex).outgoingData.WriteBoolean(False)
            Call UserList(userindex).outgoingData.WriteBoolean(False)
            Call UserList(userindex).outgoingData.WriteBoolean(False)
        Else
            Call UserList(userindex).outgoingData.WriteBoolean(CheckBonusGroup(GroupIndex, GroupFull))
            Call UserList(userindex).outgoingData.WriteBoolean(CheckBonusGroup(GroupIndex, LeaderPremium))
            Call UserList(userindex).outgoingData.WriteBoolean(CheckBonusGroup(GroupIndex, LeaderPendient))
            Call UserList(userindex).outgoingData.WriteBoolean(CheckBonusGroup(GroupIndex, LeaderMaxLevel))
        End If
    End With

              
    Exit Sub

Errhandler:
160       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
170           Call FlushBuffer(userindex)
180           Resume
190       End If
End Sub

Public Sub WriteGroupRequests(ByVal userindex As Integer)

On Error GoTo Errhandler

    Dim GroupIndex As Integer
    Dim A As Long, B As Long
    
    Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.GroupRequests)
    
    GroupIndex = UserList(userindex).GroupIndex
    
    With Groups(GroupIndex)
        For A = 1 To MAX_REQUESTS_GROUP
            Call UserList(userindex).outgoingData.WriteASCIIString(.Requests(A))
        Next A
    End With

              
    Exit Sub

Errhandler:
160       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
170           Call FlushBuffer(userindex)
180           Resume
190       End If
End Sub
Public Sub WriteUpdateKey(ByVal userindex As Integer, _
                            ByVal Packet As eKeyPackets, _
                            ByVal key As Long)
On Error GoTo Errhandler

    Dim GroupIndex As Integer
    Dim A As Long, B As Long
    
    Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.UpdateKey)
    Call UserList(userindex).outgoingData.WriteByte(Packet)
    Call UserList(userindex).outgoingData.WriteLong(key)
    Exit Sub

Errhandler:
160       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
170           Call FlushBuffer(userindex)
180           Resume
190       End If
End Sub
Public Sub WriteGroupReward(ByVal userindex As Integer)

On Error GoTo Errhandler

    Dim GroupIndex As Integer
    Dim A As Long, B As Long
    
    Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.GroupReward)
    
    GroupIndex = UserList(userindex).GroupIndex
    
    With Groups(GroupIndex)
        Call UserList(userindex).outgoingData.WriteLong(.User(UserList(userindex).GroupSlotUser).Exp)
        Call UserList(userindex).outgoingData.WriteLong(.User(UserList(userindex).GroupSlotUser).Gld)
    End With

              
    Exit Sub

Errhandler:
160       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
170           Call FlushBuffer(userindex)
180           Resume
190       End If
End Sub

Public Sub HandleGroupMember(ByVal userindex As Integer)
10        If UserList(userindex).incomingData.length < 3 Then
20            Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo Errhandler
60        With UserList(userindex)
              'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
70            Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
80            Call buffer.CopyBuffer(.incomingData)

              'Remove packet ID
90            Call buffer.ReadByte

              Dim UserName As String
100           Dim tipo As Byte

              tipo = buffer.ReadByte()
              UserName = buffer.ReadASCIIString()
              
              Select Case tipo
                Case 1 'Aceptar miembro
                    mGroup.AcceptInvitationGroup userindex, UserName
                Case 2 'Rechazar miembro
                    mGroup.RechaceInvitationGroup userindex, UserName
              End Select
            
              'If we got here then packet is complete, copy data back to original queue
110           Call .incomingData.CopyBuffer(buffer)

190       End With

Errhandler:
          Dim error  As Long
200       error = Err.Number
210       On Error GoTo 0

          'Destroy auxiliar buffer
220       Set buffer = Nothing

230       If error <> 0 Then _
             Err.Raise error
End Sub

Public Sub HandleGroupChangePorc(ByVal userindex As Integer)

    Dim A As Byte
    Dim Gld(4) As Byte
    Dim Exp(4) As Byte
    
    
    With UserList(userindex)
        Call .incomingData.ReadByte
        
        For A = 0 To 4
            Exp(A) = .incomingData.ReadByte
            Gld(A) = .incomingData.ReadByte
        Next A
        
        If .GroupIndex > 0 Then
            mGroup.GroupSetPorcentaje userindex, .GroupIndex, Exp, Gld
        End If
        
    End With
End Sub

Public Sub HandleSendCaptureImage(ByVal userindex As Integer)
    Dim A As Long
    Dim Value As Long
    Dim Bytes() As Byte
    
    
    With UserList(userindex)
      '  Call .incomingData.ReadByte
        
       ' Value = .incomingData.ReadLong
        
      '  For A = 0 To Value
       '     Bytes(A) = .incomingData.ReadByte
       ' Next A
        
        
       ' SaveBytes "Captura.BMP"
        
    End With
End Sub

Public Sub HandlePacketAccount(ByVal userindex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************

   On Error GoTo HandlePacketAccount_Error

10            If UserList(userindex).incomingData.length < 3 Then
20                Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
30                Exit Sub
40            End If

50        On Error GoTo Errhandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(UserList(userindex).incomingData)

          'Remove packet ID
70        Call buffer.ReadByte
          Dim key As String
          Dim Version As String
          
          
        Version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
        key = buffer.ReadASCIIString()
        
        If Not VersionOK(Version) Then
            Call WriteErrorMsg(userindex, "Esta versión es vieja. Por favor EJECUTA el UpdaterDesterium.exe para que se actualice, que está en la carpeta donde está el juego. La correcta es: " & ULTIMAVERSION)
            
            Call FlushBuffer(userindex)
        Else
            Select Case key
                ' ¡NOTA! Hacer que chequee Email, passwd y demas datos vacios
                Case "Create$Account"
                    mCuenta.CreateAccount userindex, buffer.ReadASCIIString, buffer.ReadASCIIString, buffer.ReadASCIIString, buffer.ReadASCIIString
                Case "Login$Account"
                    mCuenta.LoginAccount userindex, buffer.ReadASCIIString, buffer.ReadASCIIString
                Case "LoginChar$Account"
                    mCuenta.LoginCharAccount userindex, buffer.ReadASCIIString, buffer.ReadASCIIString, buffer.ReadASCIIString
                Case "CreateChar$Account"
                    mCuenta.CreateCharAccount userindex, buffer.ReadASCIIString, buffer.ReadASCIIString, buffer.ReadASCIIString, buffer.ReadByte, buffer.ReadByte, buffer.ReadByte
                Case "RemoveChar$Account"
                    mCuenta.KillCharAccount userindex, buffer.ReadASCIIString, buffer.ReadASCIIString, buffer.ReadByte
                Case "Recover$Account"
                    mCuenta.RecoverAccount userindex, buffer.ReadASCIIString, buffer.ReadASCIIString, buffer.ReadASCIIString
                Case "AddTemporal$Account"
                    mCuenta.AddTemporal userindex, buffer.ReadASCIIString, buffer.ReadASCIIString, buffer.ReadASCIIString, buffer.ReadASCIIString, buffer.ReadASCIIString, buffer.ReadASCIIString
            End Select
        End If

          'If we got here then packet is complete, copy data back to original queue
520       Call UserList(userindex).incomingData.CopyBuffer(buffer)
Errhandler:
          
          Dim error  As Long
530       error = Err.Number
540       On Error GoTo 0

          'Destroy auxiliar buffer
550       Set buffer = Nothing
          
560       If error <> 0 Then _
             Err.Raise error

   On Error GoTo 0
   Exit Sub

HandlePacketAccount_Error:
End Sub


Public Sub WriteAccount_Data(ByVal userindex As Integer, _
                                ByRef Chars() As tCuentaUser)
          
   On Error GoTo WriteAccount_Data_Error

10        On Error GoTo Errhandler
            
            Dim LoopC As Integer
            
            Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.Account_Data)
            
            For LoopC = 1 To MAX_PJS_ACCOUNT
                With Chars(LoopC)
                    Call UserList(userindex).outgoingData.WriteASCIIString(.Name)
                    Call UserList(userindex).outgoingData.WriteByte(.Ban)
                    Call UserList(userindex).outgoingData.WriteByte(.clase)
                    Call UserList(userindex).outgoingData.WriteByte(.ELV)
                    Call UserList(userindex).outgoingData.WriteByte(.raza)
                End With
            Next LoopC
              
150       Exit Sub

Errhandler:
160       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
170           Call FlushBuffer(userindex)
180           Resume
190       End If

   On Error GoTo 0
   Exit Sub

WriteAccount_Data_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteAccount_Data of Módulo Protocol in line " & Erl
End Sub


Public Sub WriteInfoPj(ByVal userindex As Integer, ByVal UserName As String)
On Error GoTo Errhandler

    Dim tUser As Integer
    Dim A As Long
    Dim FilePath As String
    
    FilePath = CharPath & UserName & ".chr"
    tUser = NameIndex(tUser)
    
    Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.SendInfoPjMao)
    
    If tUser > 0 Then
        With UserList(tUser)
            Call UserList(userindex).outgoingData.WriteByte(.Stats.ELV)
            Call UserList(userindex).outgoingData.WriteByte(.clase)
            Call UserList(userindex).outgoingData.WriteByte(.raza)
            Call UserList(userindex).outgoingData.WriteByte(.flags.BonosHP)
            Call UserList(userindex).outgoingData.WriteInteger(.Stats.MaxHp)
            Call UserList(userindex).outgoingData.WriteInteger(.Stats.MaxMAN)
            Call UserList(userindex).outgoingData.WriteLong(.Stats.Gld)
            Call UserList(userindex).outgoingData.WriteLong(.Reputacion.AsesinoRep)
            Call UserList(userindex).outgoingData.WriteLong(.Reputacion.BandidoRep)
            Call UserList(userindex).outgoingData.WriteBoolean(criminal(tUser))
            Call UserList(userindex).outgoingData.WriteBoolean(HasFound(UserName))
            Call UserList(userindex).outgoingData.WriteByte(.flags.Oro)
            Call UserList(userindex).outgoingData.WriteByte(.flags.Premium)
            
            For A = 1 To Declaraciones.MAX_INVENTORY_SLOTS
                Call UserList(userindex).outgoingData.WriteInteger(.Invent.Object(A).ObjIndex)
                Call UserList(userindex).outgoingData.WriteInteger(.Invent.Object(A).Amount)
                
                If .Invent.Object(A).ObjIndex > 0 Then
                    Call UserList(userindex).outgoingData.WriteInteger(ObjData(.Invent.Object(A).ObjIndex).GrhIndex)
                Else
                    Call UserList(userindex).outgoingData.WriteInteger(0)
                End If
            Next A
            
            For A = 1 To Declaraciones.MAX_BANCOINVENTORY_SLOTS
                Call UserList(userindex).outgoingData.WriteInteger(.BancoInvent.Object(A).ObjIndex)
                Call UserList(userindex).outgoingData.WriteInteger(.BancoInvent.Object(A).Amount)
                
                If .BancoInvent.Object(A).ObjIndex > 0 Then
                    Call UserList(userindex).outgoingData.WriteInteger(ObjData(.BancoInvent.Object(A).ObjIndex).GrhIndex)
                Else
                    Call UserList(userindex).outgoingData.WriteInteger(0)
                End If
            Next A
            
            For A = 1 To NUMSKILLS
                Call UserList(userindex).outgoingData.WriteByte(.Stats.UserSkills(A))
            Next A
            
            For A = 1 To Declaraciones.MAXUSERHECHIZOS
                Call UserList(userindex).outgoingData.WriteByte(.Stats.UserHechizos(A))
            Next A
            
            For A = 1 To 5
                Call UserList(userindex).outgoingData.WriteByte(.Stats.UserAtributos(A))
            Next A
            
        End With
    Else
            Call UserList(userindex).outgoingData.WriteByte(val(GetVar(FilePath, "STATS", "ELV")))
            Call UserList(userindex).outgoingData.WriteByte(val(GetVar(FilePath, "INIT", "CLASE")))
            Call UserList(userindex).outgoingData.WriteByte(val(GetVar(FilePath, "INIT", "RAZA")))
            Call UserList(userindex).outgoingData.WriteByte(val(GetVar(FilePath, "FLAGS", "BONOSHP")))
            Call UserList(userindex).outgoingData.WriteInteger(val(GetVar(FilePath, "STATS", "MAXHP")))
            Call UserList(userindex).outgoingData.WriteInteger(val(GetVar(FilePath, "STATS", "MAXMAN")))
            Call UserList(userindex).outgoingData.WriteLong(val(GetVar(FilePath, "STATS", "GLD")))
            Call UserList(userindex).outgoingData.WriteLong(val(GetVar(FilePath, "REP", "ASESINO")))
            Call UserList(userindex).outgoingData.WriteLong(val(GetVar(FilePath, "REP", "BANDIDO")))
            Call UserList(userindex).outgoingData.WriteBoolean(val(GetVar(FilePath, "REP", "PROMEDIO")) > 0) ' Criminal
            Call UserList(userindex).outgoingData.WriteBoolean(HasFound(UserName)) ' Fundador
            Call UserList(userindex).outgoingData.WriteByte(val(GetVar(FilePath, "FLAGS", "ORO")))
            Call UserList(userindex).outgoingData.WriteByte(val(GetVar(FilePath, "FLAGS", "PREMIUM")))
            
            Dim Temp As String
            
            For A = 1 To Declaraciones.MAX_INVENTORY_SLOTS
                Temp = GetVar(FilePath, "INVENTORY", "OBJ" & A)
                Call UserList(userindex).outgoingData.WriteInteger(val(ReadField(1, Temp, Asc("-"))))
                Call UserList(userindex).outgoingData.WriteInteger(val(ReadField(2, Temp, Asc("-"))))
                
                Temp = val(ReadField(1, Temp, Asc("-")))
                
                If Temp <> 0 Then
                    Call UserList(userindex).outgoingData.WriteInteger(ObjData(Temp).GrhIndex)
                Else
                    Call UserList(userindex).outgoingData.WriteInteger(0)
                End If
            Next A
                
            For A = 1 To Declaraciones.MAX_BANCOINVENTORY_SLOTS
                Temp = GetVar(FilePath, "BANCOINVENTORY", "OBJ" & A)
                Call UserList(userindex).outgoingData.WriteInteger(val(ReadField(1, Temp, Asc("-"))))
                Call UserList(userindex).outgoingData.WriteInteger(val(ReadField(2, Temp, Asc("-"))))
                
                Temp = val(ReadField(1, Temp, Asc("-")))
                
                If Temp <> 0 Then
                    Call UserList(userindex).outgoingData.WriteInteger(ObjData(Temp).GrhIndex)
                Else
                    Call UserList(userindex).outgoingData.WriteInteger(0)
                End If
            Next A
                
            For A = 1 To NUMSKILLS
                Call UserList(userindex).outgoingData.WriteByte(val(GetVar(FilePath, "SKILLS", "SK" & A)))
            Next A
                
            For A = 1 To Declaraciones.MAXUSERHECHIZOS
                Temp = val(GetVar(FilePath, "HECHIZOS", "H" & A))
                Call UserList(userindex).outgoingData.WriteByte(Temp)
            Next A
                
            For A = 1 To 5
                Temp = val(GetVar(FilePath, "ATRIBUTOS", "AT" & A))
                Call UserList(userindex).outgoingData.WriteByte(Temp)
            Next A
    End If
    

150       Exit Sub

Errhandler:
160       If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
170           Call FlushBuffer(userindex)
180           Resume
190       End If
End Sub
Public Function WriteListText(ByVal userindex As Integer, ByVal Num As Integer, ByVal Datos As String, ByVal Obj As Boolean) As String
 
On Error GoTo Errhandler
    With UserList(userindex).outgoingData
        Call .WriteByte(ServerPacketID.ListText)
        Call .WriteInteger(Num)
        Call .WriteBoolean(Obj)
        Call .WriteASCIIString(Datos)
    End With
 
Errhandler:
    If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userindex)
        Resume
    End If
   
End Function
 
Public Sub HandleSearchNpc(ByVal userindex As Integer)
 
On Error GoTo Errhandler
    With UserList(userindex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
       
    Call buffer.ReadByte
       
        Dim i As Integer
        Dim n As Integer
        Dim Name As String
        Dim UserNpc As String
        Dim tStr As String
        UserNpc = buffer.ReadASCIIString()
       
   
    tStr = Tilde(UserNpc)
    If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
        For i = 1 To val(LeerNPCs.GetValue("INIT", "NumNPCs"))
        Name = LeerNPCs.GetValue("NPC" & i, "Name")
       
        If InStr(1, Tilde(Name), tStr) Then
            Call WriteListText(userindex, i, i & " - " & Name, False)
            n = n + 1
        End If
    Next i
   
        If n = 0 Then
            Call WriteListText(userindex, 0, "No hubo resultados de la busqueda.", False)
        End If
    End If
   
        Call .incomingData.CopyBuffer(buffer)
        End With
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
   
    Set buffer = Nothing
   
    If error <> 0 Then _
        Err.Raise error
End Sub
 
Private Sub HandleSearchObj(ByVal userindex As Integer)
    If UserList(userindex).incomingData.length < 3 Then
        Err.Raise UserList(userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
       
On Error GoTo Errhandler
    With UserList(userindex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
           
        Call buffer.ReadByte
           
        Dim UserObj As String
        Dim tUser As Integer
        Dim Rank As Integer
        Dim n As Integer
        Dim i As Integer
           
        Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
           
        UserObj = buffer.ReadASCIIString()
           
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
               
            For i = 1 To UBound(ObjData)
                If InStr(1, Tilde(ObjData(i).Name), Tilde(UserObj)) Then
                    Call WriteListText(userindex, i, i & " - " & ObjData(i).Name, True)
                    n = n + 1
                End If
            Next
            If n = 0 Then
                Call WriteListText(userindex, 0, "No hubo resultados de la busqueda.", False)
            End If
               
        End If
           
        Call .incomingData.CopyBuffer(buffer)
    End With
     
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
       
    Set buffer = Nothing
       
    If error <> 0 Then _
        Err.Raise error
End Sub
 
Private Sub HandleSearcherShow(ByVal userindex As Integer)
 
    With UserList(userindex)
        Call .incomingData.ReadByte
   
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowSearcher(userindex)
       
    End With
   
End Sub
 
Public Sub WriteShowSearcher(ByVal userindex As Integer)
 
On Error GoTo Errhandler
 
    Call UserList(userindex).outgoingData.WriteByte(ServerPacketID.ShowSearcher)
    Exit Sub
 
Errhandler:
    If Err.Number = UserList(userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userindex)
        Resume
    End If
End Sub

