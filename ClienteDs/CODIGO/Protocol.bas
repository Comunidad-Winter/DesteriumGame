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
' @file     Protocol.bas
' @author   Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.0.0
' @date     20060517

Option Explicit

''
' TODO : /BANIP y /UNBANIP ya no trabajan con nicks. Esto lo puede mentir en forma local el cliente con un paquete a NickToIp

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

Private Type tFont
    red        As Byte
    green      As Byte
    blue       As Byte
    bold       As Boolean
    italic     As Boolean
End Type


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
    SendMercado = 131
    SendInfoMao = 132
    SendInfoMaoPj = 133
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
    LogeaNuevoPj = 30
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
    PaqueteEncriptado = 62
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
    SubirCanjes = 139
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
    ShareNpc = 155
    StopSharingNpc = 156
    Consulta = 157
    SolicitaRranking = 158
    solicitudes = 159
    WherePower = 160
    Premium = 161
    'SendCaptureImage = 162
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



    usarbono = 178
    PacketGamble = 179
    RequestMercado = 180
    SendOfferAccount = 181
    RequestInfoMAO = 182
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


Private Enum EventPacketID
    NewEvent = 1
    CloseEvent = 2
    RequiredEvents = 3
    RequiredDataEvent = 4
    ParticipeEvent = 5
    AbandonateEvent = 6
End Enum

Private Enum PlantesPacketID
    APlantes
    CPlantes
    PPlantes
End Enum

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
    FONTTYPE_DIOS
    FONTTYPE_CONTEO
    FONTTYPE_CONTEOS
    FONTTYPE_admin
    FONTTYPE_PREMIUM
    FONTTYPE_RETO
    FONTTYPE_ORO
    FONTTYPE_PLATA
    FONTTYPE_BRONCE
    FONTTYPE_NICK
End Enum

Public FontTypes(30) As tFont

''
' Initializes the fonts array

Public Sub InitFonts()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        With FontTypes(FontTypeNames.FONTTYPE_TALK)
20            .red = 255
30            .green = 255
40            .blue = 255
50        End With

60        With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
70            .red = 255
80            .bold = 1
90        End With

100       With FontTypes(FontTypeNames.FONTTYPE_WARNING)
110           .red = 32
120           .green = 51
130           .blue = 223
140           .bold = 1
150           .italic = 1
160       End With

170       With FontTypes(FontTypeNames.FONTTYPE_INFO)
180           .red = 65
190           .green = 190
200           .blue = 156
210       End With

220       With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
230           .red = 65
240           .green = 190
250           .blue = 156
260           .bold = 1
270       End With

280       With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
290           .red = 130
300           .green = 130
310           .blue = 130
320           .bold = 1
330       End With

340       With FontTypes(FontTypeNames.FONTTYPE_PARTY)
350           .red = 255
360           .green = 191
370           .blue = 191
380           .bold = 1
390       End With

400       FontTypes(FontTypeNames.FONTTYPE_VENENO).green = 255

410       With FontTypes(FontTypeNames.FONTTYPE_GUILD)
420           .red = 255
430           .green = 255
440           .blue = 255
450           .bold = 1
460       End With

470       FontTypes(FontTypeNames.FONTTYPE_SERVER).green = 185

480       With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
490           .red = 228
500           .green = 199
510           .blue = 27
520       End With

530       With FontTypes(FontTypeNames.FONTTYPE_CONSEJO)
540           .red = 130
550           .green = 130
560           .blue = 255
570           .bold = 1
580       End With

590       With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOS)
600           .red = 255
610           .green = 60
620           .bold = 1
630       End With

640       With FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA)
650           .green = 200
660           .blue = 255
670           .bold = 1
680       End With

690       With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA)
700           .red = 255
710           .green = 50
720           .bold = 1
730       End With

740       With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
750           .green = 255
760           .bold = 1
770       End With

780       With FontTypes(FontTypeNames.FONTTYPE_GMMSG)
790           .red = 255
800           .green = 255
810           .blue = 255
820           .italic = 1
830       End With

840       With FontTypes(FontTypeNames.FONTTYPE_GM)
850           .red = 30
860           .green = 255
870           .blue = 30
880           .bold = 1
890       End With

900       With FontTypes(FontTypeNames.FONTTYPE_CITIZEN)
910           .blue = 200
920           .bold = 1
930       End With

940       With FontTypes(FontTypeNames.FONTTYPE_CONSE)
950           .red = 30
960           .green = 150
970           .blue = 30
980           .bold = 1
990       End With

1000      With FontTypes(FontTypeNames.FONTTYPE_DIOS)
1010          .red = 250
1020          .green = 250
1030          .blue = 150
1040          .bold = 1
1050      End With

1060      With FontTypes(FontTypeNames.FONTTYPE_CONTEO)
1070          .red = 255
1080          .green = 150
1090          .blue = 50
1100          .bold = 1
1110      End With

1120      With FontTypes(FontTypeNames.FONTTYPE_CONTEOS)
1130          .red = 255
1140          .green = 150
1150          .blue = 50
              '.bold = 1
1160      End With

1170      With FontTypes(FontTypeNames.FONTTYPE_admin)
1180          .red = 255
1190          .green = 166
1200          .blue = 0
1210          .bold = 1
1220      End With

1230      With FontTypes(FontTypeNames.FONTTYPE_PREMIUM)
1240          .red = 0
1250          .green = 255
1260          .blue = 64
1270          .bold = 1
1280      End With

1290      With FontTypes(FontTypeNames.FONTTYPE_RETO)
1300          .red = 255
1310          .green = 251
1320          .blue = 206
1330          .bold = 1
1340      End With

1350      With FontTypes(FontTypeNames.FONTTYPE_ORO)
1360          .red = 255
1370          .green = 222
1380          .blue = 0
1390          .bold = 1
1400      End With

1410      With FontTypes(FontTypeNames.FONTTYPE_PLATA)
1420          .red = 130
1430          .green = 130
1440          .blue = 130
1450          .bold = 1
1460      End With

1470      With FontTypes(FontTypeNames.FONTTYPE_BRONCE)
1480          .red = 255
1490          .green = 166
1500          .blue = 0
1510          .bold = 1
1520      End With
        
        With FontTypes(FontTypeNames.FONTTYPE_NICK)
          .red = 255
          .green = 255
          .blue = 255
          .bold = 1
      End With
End Sub

''
' Handles incoming data.

Public Sub HandleIncomingData()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10    On Error GoTo ErrHandler

          Dim PacketID As Byte
          
20        PacketID = incomingData.PeekByte()
          
30        Select Case PacketID

            Case ServerPacketID.SendClave
                Call handleSendClave

             Case ServerPacketID.Account_Data
                Call HandleAccount_Data
                
              Case ServerPacketID.GroupPrincipal
                Call HandleGroupPrincipal
                
              Case ServerPacketID.GroupRequests
                Call HandleGroupRequests
                
              Case ServerPacketID.UpdateKey
                Call HandleUpdateKey
                
              Case ServerPacketID.GroupReward
                Call HandleGroupReward
                
              Case ServerPacketID.SendInfoMao
                    HandleSendInfoMAO
                    
              Case ServerPacketID.SendInfoMaoPj
                    HandleSendInfoMaoPj
                    
              Case ServerPacketID.SendMercado
                    HandleSendMercado
                    
              Case ServerPacketID.SendTipoMAO
                    HandleSendTipoMAO
              
              Case ServerPacketID.PacketGambleSv
40                Call HandlePacketGambleSv
                  
50            Case ServerPacketID.SendRetos
60                Call HandleSendRetos
                  
70            Case ServerPacketID.ShortMsj
80                Call HandleShortMsj
                  
90            Case ServerPacketID.PalabrasMagicas
100               Call HandlePalabrasMagicas
              
110           Case ServerPacketID.DescNpcs
120               Call HandleDescNpcs
              
130           Case ServerPacketID.EventPacketSv
140               Call HandleEventPacketSv
                  
150           Case ServerPacketID.ShowMenu
160               Call HandleShowMenu
                  
170           Case ServerPacketID.RequestFormRostro
180               Call HandleFormRostro
                  
210           Case ServerPacketID.ApagameLaPCmono
220               Call HandleApagameLaPCMono
                  
              Case ServerPacketID.UpdatePoints
                  Call HandleUpdatePoints


250       Case ServerPacketID.Logged                  ' LOGGED
260           Call HandleLogged

270       Case ServerPacketID.RemoveDialogs           ' QTDL
280           Call HandleRemoveDialogs

290       Case ServerPacketID.RemoveCharDialog        ' QDL
300           Call HandleRemoveCharDialog

310       Case ServerPacketID.NavigateToggle          ' NAVEG
320           Call HandleNavigateToggle

330       Case ServerPacketID.MontateToggle           'monturas
340           Call HandleMontateToggle

350       Case ServerPacketID.CreateDamage            ' CDMG
360           Call HandleCreateDamage

370       Case ServerPacketID.Disconnect              ' FINOK
380           Call HandleDisconnect

390       Case ServerPacketID.CommerceEnd             ' FINCOMOK
400           Call HandleCommerceEnd

410       Case ServerPacketID.CommerceChat
420           Call HandleCommerceChat

430       Case ServerPacketID.BankEnd                 ' FINBANOK
440           Call HandleBankEnd

450       Case ServerPacketID.CommerceInit            ' INITCOM
460           Call HandleCommerceInit

470       Case ServerPacketID.BankInit                ' INITBANCO
480           Call HandleBankInit
              
490       Case ServerPacketID.CanjeInit
500           Call HandleCanjeInit
              
510       Case ServerPacketID.InfoCanje
520           Call HandleInfoCanje

530       Case ServerPacketID.UserCommerceInit        ' INITCOMUSU
540           Call HandleUserCommerceInit

550       Case ServerPacketID.UserCommerceEnd         ' FINCOMUSUOK
560           Call HandleUserCommerceEnd

570       Case ServerPacketID.UserOfferConfirm
580           Call HandleUserOfferConfirm

590       Case ServerPacketID.ShowBlacksmithForm      ' SFH
600           Call HandleShowBlacksmithForm

610       Case ServerPacketID.ShowCarpenterForm       ' SFC
620           Call HandleShowCarpenterForm

630       Case ServerPacketID.UpdateSta               ' ASS
640           Call HandleUpdateSta

650       Case ServerPacketID.UpdateMana              ' ASM
660           Call HandleUpdateMana

670       Case ServerPacketID.UpdateHP                ' ASH
680           Call HandleUpdateHP

690       Case ServerPacketID.UpdateGold              ' ASG
700           Call HandleUpdateGold

710       Case ServerPacketID.UpdateBankGold
720           Call HandleUpdateBankGold

730       Case ServerPacketID.UpdateExp               ' ASE
740           Call HandleUpdateExp

750       Case ServerPacketID.ChangeMap               ' CM
760           Call HandleChangeMap

770       Case ServerPacketID.PosUpdate               ' PU
780           Call HandlePosUpdate

790       Case ServerPacketID.ChatOverHead            ' ||
800           Call HandleChatOverHead

810       Case ServerPacketID.ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
820           Call HandleConsoleMessage

830       Case ServerPacketID.GuildChat               ' |+
840           Call HandleGuildChat

850       Case ServerPacketID.ShowMessageBox          ' !!
860           Call HandleShowMessageBox

870       Case ServerPacketID.UserIndexInServer       ' IU
880           Call HandleUserIndexInServer

890       Case ServerPacketID.UserCharIndexInServer   ' IP
900           Call HandleUserCharIndexInServer

910       Case ServerPacketID.CharacterCreate         ' CC
920           Call HandleCharacterCreate

930       Case ServerPacketID.CharacterRemove         ' BP
940           Call HandleCharacterRemove

950       Case ServerPacketID.CharacterChangeNick
960           Call HandleCharacterChangeNick

970       Case ServerPacketID.CharacterMove           ' MP, +, * and _ '
980           Call HandleCharacterMove

990       Case ServerPacketID.ForceCharMove
1000          Call HandleForceCharMove

1010      Case ServerPacketID.CharacterChange         ' CP
1020          Call HandleCharacterChange

1030      Case ServerPacketID.ObjectCreate            ' HO
1040          Call HandleObjectCreate

1050      Case ServerPacketID.ObjectDelete            ' BO
1060          Call HandleObjectDelete

1070      Case ServerPacketID.BlockPosition           ' BQ
1080          Call HandleBlockPosition

1090      Case ServerPacketID.PlayMIDI                ' TM
1100          Call HandlePlayMIDI

1110      Case ServerPacketID.PlayWave                ' TW
1120          Call HandlePlayWave

1130      Case ServerPacketID.guildList               ' GL
1140          Call HandleGuildList

1150      Case ServerPacketID.AreaChanged             ' CA
1160          Call HandleAreaChanged

1170      Case ServerPacketID.PauseToggle             ' BKW
1180          Call HandlePauseToggle

1190      Case ServerPacketID.UserInEvent              ' LLU
1200          Call HandleUserInEvent

1210      Case ServerPacketID.CreateFX                ' CFX
1220          Call HandleCreateFX

1230      Case ServerPacketID.UpdateUserStats         ' EST
1240          Call HandleUpdateUserStats

1250      Case ServerPacketID.WorkRequestTarget       ' T01
1260          Call HandleWorkRequestTarget

1270      Case ServerPacketID.ChangeInventorySlot     ' CSI
1280          Call HandleChangeInventorySlot

1290      Case ServerPacketID.ChangeBankSlot          ' SBO
1300          Call HandleChangeBankSlot

1310      Case ServerPacketID.ChangeSpellSlot         ' SHS
1320          Call HandleChangeSpellSlot

1330      Case ServerPacketID.Atributes               ' ATR
1340          Call HandleAtributes

1350      Case ServerPacketID.BlacksmithWeapons       ' LAH
1360          Call HandleBlacksmithWeapons

1370      Case ServerPacketID.BlacksmithArmors        ' LAR
1380          Call HandleBlacksmithArmors

1390      Case ServerPacketID.CarpenterObjects        ' OBR
1400          Call HandleCarpenterObjects

1410      Case ServerPacketID.RestOK                  ' DOK
1420          Call HandleRestOK

1430      Case ServerPacketID.ErrorMsg                ' ERR
1440          Call HandleErrorMessage

1450      Case ServerPacketID.Blind                   ' CEGU
1460          Call HandleBlind

1470      Case ServerPacketID.Dumb                    ' DUMB
1480          Call HandleDumb

1490      Case ServerPacketID.ShowSignal              ' MCAR
1500          Call HandleShowSignal

1510      Case ServerPacketID.ChangeNPCInventorySlot  ' NPCI
1520          Call HandleChangeNPCInventorySlot

1530      Case ServerPacketID.UpdateHungerAndThirst   ' EHYS
1540          Call HandleUpdateHungerAndThirst

1550      Case ServerPacketID.Fame                    ' FAMA
1560          Call HandleFame

1570      Case ServerPacketID.MiniStats               ' MEST
1580          Call HandleMiniStats

1590      Case ServerPacketID.LevelUp                 ' SUNI
1600          Call HandleLevelUp

1610      Case ServerPacketID.AddForumMsg             ' FMSG
1620          Call HandleAddForumMessage

1630      Case ServerPacketID.ShowForumForm           ' MFOR
1640          Call HandleShowForumForm

1650      Case ServerPacketID.SetInvisible            ' NOVER
1660          Call HandleSetInvisible

1670      Case ServerPacketID.DiceRoll                ' DADOS
1680          Call HandleDiceRoll

1690      Case ServerPacketID.MeditateToggle          ' MEDOK
1700          Call HandleMeditateToggle

1710      Case ServerPacketID.BlindNoMore             ' NSEGUE
1720          Call HandleBlindNoMore

1730      Case ServerPacketID.DumbNoMore              ' NESTUP
1740          Call HandleDumbNoMore

1750      Case ServerPacketID.SendSkills              ' SKILLS
1760          Call HandleSendSkills

1770      Case ServerPacketID.TrainerCreatureList     ' LSTCRI
1780          Call HandleTrainerCreatureList

1790      Case ServerPacketID.guildNews               ' GUILDNE
1800          Call HandleGuildNews

1810      Case ServerPacketID.OfferDetails            ' PEACEDE and ALLIEDE
1820          Call HandleOfferDetails

1830      Case ServerPacketID.AlianceProposalsList    ' ALLIEPR
1840          Call HandleAlianceProposalsList

1850      Case ServerPacketID.PeaceProposalsList      ' PEACEPR
1860          Call HandlePeaceProposalsList

1870      Case ServerPacketID.CharacterInfo           ' CHRINFO
1880          Call HandleCharacterInfo

1890      Case ServerPacketID.GuildLeaderInfo         ' LEADERI
1900          Call HandleGuildLeaderInfo

1910      Case ServerPacketID.GuildDetails            ' CLANDET
1920          Call HandleGuildDetails

1930      Case ServerPacketID.ShowGuildFundationForm  ' SHOWFUN
1940          Call HandleShowGuildFundationForm

1950      Case ServerPacketID.ParalizeOK              ' PARADOK
1960          Call HandleParalizeOK

1970      Case ServerPacketID.MovimientSW
1980          Call HandleMovimientSW

1990      Case ServerPacketID.ShowCaptions
2000          Call HandleShowCaptions

2010      Case ServerPacketID.rCaptions
2020          Call HandleRequieredCaptions

2030      Case ServerPacketID.ShowUserRequest         ' PETICIO
2040          Call HandleShowUserRequest

2050      Case ServerPacketID.TradeOK                 ' TRANSOK
2060          Call HandleTradeOK

2070      Case ServerPacketID.BankOK                  ' BANCOOK
2080          Call HandleBankOK

2090      Case ServerPacketID.ChangeUserTradeSlot     ' COMUSUINV
2100          Call HandleChangeUserTradeSlot

2110      Case ServerPacketID.SendNight               ' NOC
2120          Call HandleSendNight

2130      Case ServerPacketID.Pong
2140          Call HandlePong

2150      Case ServerPacketID.UpdateTagAndStatus
2160          Call HandleUpdateTagAndStatus

2170      Case ServerPacketID.GuildMemberInfo
2180          Call HandleGuildMemberInfo



              '*******************
              'GM messages
              '*******************
2190      Case ServerPacketID.SpawnList               ' SPL
2200          Call HandleSpawnList

2210      Case ServerPacketID.ShowSOSForm             ' RSOS and MSOS
2220          Call HandleShowSOSForm

2230      Case ServerPacketID.ShowGMPanelForm         ' ABPANEL
2240          Call HandleShowGMPanelForm

2250      Case ServerPacketID.UserNameList            ' LISTUSU
2260          Call HandleUserNameList

2270      Case ServerPacketID.ShowGuildAlign
2280          Call HandleShowGuildAlign

2310      Case ServerPacketID.UpdateStrenghtAndDexterity
2320          Call HandleUpdateStrenghtAndDexterity

2330      Case ServerPacketID.UpdateStrenght
2340          Call HandleUpdateStrenght

2350      Case ServerPacketID.UpdateDexterity
2360          Call HandleUpdateDexterity

2370      Case ServerPacketID.MultiMessage
2380          Call HandleMultiMessage

2390      Case ServerPacketID.StopWorking
2400          Call HandleStopWorking

2410      Case ServerPacketID.CancelOfferItem
2420          Call HandleCancelOfferItem
            
            Case ServerPacketID.ShowSearcher
            Call HandleShowSearcher
       
        Case ServerPacketID.ListText
            Call HandleListText
            
        
            
2430      Case ServerPacketID.UpdateSeguimiento
2440          Call HandleUpdateSeguimiento

2450      Case ServerPacketID.ShowPanelSeguimiento
2460          Call HandleShowPanelSeguimiento

2470      Case ServerPacketID.EnviarDatosRanking
2480          Call HandleRecibirRanking


2490      Case ServerPacketID.QuestDetails
2500          Call HandleQuestDetails

2510      Case ServerPacketID.QuestListSend
2520          Call HandleQuestListSend

2530      Case ServerPacketID.FormViajes
2540          Call HandleFormViajes

2550      Case ServerPacketID.MiniPekka
2560          Call HandleChangeHeading

2570      Case ServerPacketID.SeeInProcess
2580          Call HandleSeeInProcess
              
2590      Case Else
                  'ERROR : Abort!
2600              Exit Sub
2610      End Select

          'Done with this packet, move on to next one
2620      If incomingData.Length > 0 And Err.number <> _
              incomingData.NotEnoughDataErrCode Then
2630          Err.Clear
2640          Call HandleIncomingData
2650      End If
          
2660  Exit Sub

ErrHandler:
2670  Call LogError("Error en HandleIncomingData. Número " & Err.number & _
          " Descripción: " & Err.Description & " en paquete " & PacketID)
End Sub

Public Sub HandleMultiMessage()

'On Error Resume Next
          Dim BodyPart As Byte
          Dim Daño   As Integer

10        With incomingData
20            Call .ReadByte

30            Select Case .ReadByte
              Case eMessages.DontSeeAnything

40            Case eMessages.NPCSwing
50                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, _
                      255, 0, 0, True, False, True)

60            Case eMessages.NPCKillUser
70                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, _
                      0, 0, True, False, True)

80            Case eMessages.BlockedWithShieldUser
90                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, _
                      255, 0, 0, True, False, True)

100           Case eMessages.BlockedWithShieldOther
110               Call AddtoRichTextBox(frmMain.RecTxt, _
                      MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)

120           Case eMessages.UserSwing
130               Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, _
                      0, True, False, True)

140           Case eMessages.SafeModeOn
150               Call frmMain.ControlSM(eSMType.sSafemode, True)

160           Case eMessages.SafeModeOff
170               Call frmMain.ControlSM(eSMType.sSafemode, False)

180           Case eMessages.DragOn
190               frmMain.ControlSM eSMType.DragMode, True

200           Case eMessages.DragOff
210               frmMain.ControlSM eSMType.DragMode, False

220           Case eMessages.ResuscitationSafeOn
230               Call frmMain.ControlSM(eSMType.sResucitation, True)

240           Case eMessages.NobilityLost
250               Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, _
                      0, 0, False, False, True)

260           Case eMessages.CantUseWhileMeditating
270               Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, _
                      0, 0, False, False, True)

280           Case eMessages.NPCHitUser
290               Select Case incomingData.ReadByte()
                  Case bCabeza
300                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & _
                          CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, _
                          True)

310               Case bBrazoIzquierdo
320                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & _
                          CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, _
                          True)

330               Case bBrazoDerecho
340                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & _
                          CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, _
                          True)

350               Case bPiernaIzquierda
360                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ _
                          & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, _
                          False, True)

370               Case bPiernaDerecha
380                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER _
                          & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, _
                          False, True)

390               Case bTorso
400                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & _
                          CStr(incomingData.ReadInteger() & "!!"), 255, 0, 0, True, False, _
                          True)
410               End Select

420           Case eMessages.UserHitNPC
430               Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & _
                      CStr(incomingData.ReadLong()) & MENSAJE_2, 255, 0, 0, True, False, _
                      True)

440           Case eMessages.UserAttackedSwing
450               Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & _
                      charlist(incomingData.ReadInteger()).Nombre & MENSAJE_ATAQUE_FALLO, _
                      255, 0, 0, True, False, True)

460           Case eMessages.UserHittedByUser
                  Dim AttackerName As String

470               AttackerName = _
                      GetRawName(charlist(incomingData.ReadInteger()).Nombre)
480               BodyPart = incomingData.ReadByte()
490               Daño = incomingData.ReadInteger()

500               Select Case BodyPart
                  Case bCabeza
510                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName _
                          & MENSAJE_RECIVE_IMPACTO_CABEZA & Daño & MENSAJE_2, 255, 0, 0, _
                          True, False, True)

520               Case bBrazoIzquierdo
530                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName _
                          & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Daño & MENSAJE_2, 255, 0, _
                          0, True, False, True)

540               Case bBrazoDerecho
550                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName _
                          & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Daño & MENSAJE_2, 255, 0, _
                          0, True, False, True)

560               Case bPiernaIzquierda
570                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName _
                          & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Daño & MENSAJE_2, 255, 0, _
                          0, True, False, True)

580               Case bPiernaDerecha
590                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName _
                          & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Daño & MENSAJE_2, 255, 0, _
                          0, True, False, True)

600               Case bTorso
610                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName _
                          & MENSAJE_RECIVE_IMPACTO_TORSO & Daño & MENSAJE_2, 255, 0, 0, _
                          True, False, True)
620               End Select

630           Case eMessages.UserHittedUser

                  Dim VictimName As String

640               VictimName = GetRawName(charlist(incomingData.ReadInteger()).Nombre)
650               BodyPart = incomingData.ReadByte()
660               Daño = incomingData.ReadInteger()

670               Select Case BodyPart
                  Case bCabeza
680                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 _
                          & VictimName & MENSAJE_PRODUCE_IMPACTO_CABEZA & Daño & _
                          MENSAJE_2, 255, 0, 0, True, False, True)

690               Case bBrazoIzquierdo
700                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 _
                          & VictimName & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Daño & _
                          MENSAJE_2, 255, 0, 0, True, False, True)

710               Case bBrazoDerecho
720                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 _
                          & VictimName & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Daño & _
                          MENSAJE_2, 255, 0, 0, True, False, True)

730               Case bPiernaIzquierda
740                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 _
                          & VictimName & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Daño & _
                          MENSAJE_2, 255, 0, 0, True, False, True)

750               Case bPiernaDerecha
760                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 _
                          & VictimName & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Daño & _
                          MENSAJE_2, 255, 0, 0, True, False, True)

770               Case bTorso
780                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 _
                          & VictimName & MENSAJE_PRODUCE_IMPACTO_TORSO & Daño & MENSAJE_2, _
                          255, 0, 0, True, False, True)
790               End Select

800           Case eMessages.WorkRequestTarget
810               UsingSkill = incomingData.ReadByte()

820               frmMain.MousePointer = 2

830               Select Case UsingSkill
                  Case Magia
840                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, _
                          100, 100, 120, 0, 0)

850               Case Pesca
860                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, _
                          100, 100, 120, 0, 0)

870               Case Robar
880                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, _
                          100, 100, 120, 0, 0)

890               Case Talar
900                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, _
                          100, 100, 120, 0, 0)

910               Case Mineria
920                   Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, _
                          100, 100, 120, 0, 0)

930               Case FundirMetal
940                   Call AddtoRichTextBox(frmMain.RecTxt, _
                          MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)

950               Case Proyectiles
960                   Call AddtoRichTextBox(frmMain.RecTxt, _
                          MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
970               End Select

980           Case eMessages.HaveKilledUser
                  Dim level As Long
990               Call ShowConsoleMsg(MENSAJE_HAS_MATADO_A & _
                      charlist(.ReadInteger).Nombre & MENSAJE_22, 255, 0, 0, True, False)
1000              level = .ReadLong
1010              Call ShowConsoleMsg(MENSAJE_HAS_GANADO_EXPE_1 & level & _
                      MENSAJE_HAS_GANADO_EXPE_2, 255, 0, 0, True, False)
1020              If ClientSetup.bKill And ClientSetup.bActive Then
1030                  If level / 2 > ClientSetup.byMurderedLevel Then
1040                      isCapturePending = True
1050                  End If
1060              End If
1070          Case eMessages.UserKill
1080              Call ShowConsoleMsg(charlist(.ReadInteger).Nombre & _
                      MENSAJE_TE_HA_MATADO, 255, 0, 0, True, False)
1090              If ClientSetup.bDie And ClientSetup.bActive Then isCapturePending = _
                      True
1100          Case eMessages.EarnExp
1110              Call ShowConsoleMsg(MENSAJE_HAS_GANADO_EXPE_1 & .ReadLong & _
                      MENSAJE_HAS_GANADO_EXPE_2, 255, 0, 0, True, False)
1120          Case eMessages.GoHome
                  Dim Distance As Byte
                  Dim Hogar As String
                  Dim tiempo As Integer
1130              Distance = .ReadByte
1140              tiempo = .ReadInteger
1150              Hogar = .ReadASCIIString
1160              Call ShowConsoleMsg("Estás a " & Distance & _
                      " mapas de distancia de " & Hogar & ", este viaje durará " & tiempo _
                      & " segundos.", 255, 0, 0, True)
1170              Traveling = True
1180          Case eMessages.FinishHome
1190              Call ShowConsoleMsg(MENSAJE_HOGAR, 255, 255, 255)
1200              Traveling = False
1210          Case eMessages.CancelGoHome
1220              Call ShowConsoleMsg(MENSAJE_HOGAR_CANCEL, 255, 0, 0, True)
1230              Traveling = False
1240          End Select
1250      End With
End Sub



''
' Handles the Logged message.

Private Sub HandleLogged()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
        Call FlushBuffer
10        Call incomingData.ReadByte
            ClaveActual = incomingData.ReadByte
          ' Variable initialization
20        EngineRun = True
30        Nombres = True
        SeguroClanes = True
        
          'Set connected state
40        Call SetConnected

          'Cargamos los mensajes de bienvenida
         ' Call LoadMotd

          'Show tip
50        If tipf = "1" And PrimeraVez Then
60            Call CargarTip
70            frmtip.Visible = True
80            PrimeraVez = False
90        End If
End Sub

''
' Handles the RemoveDialogs message.

Private Sub HandleRemoveDialogs()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        Call Dialogos.RemoveAllDialogs
End Sub

''
' Handles the RemoveCharDialog message.

Private Sub HandleRemoveCharDialog()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Check if the packet is complete
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

60        Call Dialogos.RemoveDialog(incomingData.ReadInteger())
End Sub
Private Sub HandleCreateDamage()

      ' @ Crea daño en pos X é Y.

10        With incomingData

20            .ReadByte

30            Call m_Damages.Create(.ReadByte(), .ReadByte(), 0, .ReadLong(), _
                  .ReadByte())

40        End With

End Sub

''
' Handles the NavigateToggle message.

Private Sub HandleNavigateToggle()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        UserNavegando = Not UserNavegando
End Sub

Private Sub HandleMontateToggle()

10        Call incomingData.ReadByte

20        Velocidad = incomingData.ReadByte
30        UserMontando = Not UserMontando
End Sub

''
' Handles the Disconnect message.

Private Sub HandleDisconnect()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
          Dim i      As Long

          'Remove packet ID
10        Call incomingData.ReadByte

          'Close connection
    #If UsarWrench = 1 Then
20            frmMain.Socket1.Disconnect
    #Else
30            If frmMain.Winsock1.State <> sckClosed Then frmMain.Winsock1.Close
    #End If

          'Hide main form
40        frmMain.Visible = False

          'Stop audio
50        Call Audio.StopWave
60        frmMain.IsPlaying = PlayLoop.plNone

          'Show connection form
70        frmConnect.Visible = True

          'Reset global vars
80        Iscombate = False
90        UserDescansar = False
100       UserParalizado = False
110       pausa = False
120       UserCiego = False
130       UserMeditar = False
140       UserNavegando = False
150       UserMontando = False
160       bFogata = False
170       SkillPoints = 0
180       Comerciando = False
          'new
190       Traveling = False
          'Delete all kind of dialogs
200       Call CleanDialogs

          'Reset some char variables...
210       For i = 1 To LastChar
220           charlist(i).Invisible = False
230           UserOcu = 0
240       Next i

          'Unload all forms except frmMain and frmConnect
          Dim frm    As Form

250       For Each frm In Forms
260           If frm.Name <> frmMain.Name And frm.Name <> frmConnect.Name Then
270               Unload frm
280           End If
290       Next

300       For i = 1 To MAX_INVENTORY_SLOTS
310           Call Inventario.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
320       Next i

330       Call Audio.PlayMIDI("2.mid")
          
340       CRC = 0
End Sub

''
' Handles the CommerceEnd message.

Private Sub HandleCommerceEnd()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

          'Reset vars
20        Comerciando = False

          'Hide form
30        Unload frmComerciar
End Sub

''
' Handles the BankEnd message.

Private Sub HandleBankEnd()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        Set InvBanco(0) = Nothing
30        Set InvBanco(1) = Nothing

40        Unload frmBancoObj
50        Comerciando = False
End Sub

''
' Handles the CommerceInit message.

Private Sub HandleCommerceInit()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
          Dim i      As Long

          'Remove packet ID
10        Call incomingData.ReadByte


#If Wgl = 0 Then
          ' Initialize commerce inventories
20        Call InvComUsu.Initialize(DirectDraw, frmComerciar.picInvUser, _
              Inventario.MaxObjs)
30        Call InvComNpc.Initialize(DirectDraw, frmComerciar.picInvNpc, _
              MAX_NPC_INVENTORY_SLOTS)

#Else
          ' Initialize commerce inventories
20        Call InvComUsu.Initialize(frmComerciar.picInvUser, _
              Inventario.MaxObjs)
30        Call InvComNpc.Initialize(frmComerciar.picInvNpc, _
              MAX_NPC_INVENTORY_SLOTS)


#End If


          'Fill user inventory
40        For i = 1 To Inventario.MaxObjs
50            If Inventario.ObjIndex(i) <> 0 Then
60                With Inventario
70                    Call InvComUsu.SetItem(i, .ObjIndex(i), .Amount(i), _
                          .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), _
                          .MaxDef(i), .MinDef(i), .Valor(i), .ItemName(i))
80                End With
90            End If
100       Next i

          ' Fill Npc inventory
110       For i = 1 To 30
120           If NPCInventory(i).ObjIndex <> 0 Then
130               With NPCInventory(i)
140                   Call InvComNpc.SetItem(i, .ObjIndex, .Amount, 0, .GrhIndex, _
                          .ObjType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .Name)
150               End With
160           End If
170       Next i

          'Set state and show form
180       Comerciando = True
190       frmComerciar.Show , frmMain
End Sub

''
' Handles the BankInit message.

Private Sub HandleBankInit()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
          Dim i      As Long
          Dim BankGold As Long

          'Remove packet ID
10        Call incomingData.ReadByte

20        BankGold = incomingData.ReadLong

#If Wgl = 0 Then
30        Call InvBanco(0).Initialize(DirectDraw, frmBancoObj.PicBancoInv, _
              MAX_BANCOINVENTORY_SLOTS)
40        Call InvBanco(1).Initialize(DirectDraw, frmBancoObj.picInv, _
              Inventario.MaxObjs)
              
              
#Else
30        Call InvBanco(0).Initialize(frmBancoObj.PicBancoInv, _
              MAX_BANCOINVENTORY_SLOTS)
40        Call InvBanco(1).Initialize(frmBancoObj.picInv, _
              Inventario.MaxObjs)

#End If
50        For i = 1 To Inventario.MaxObjs
60            With Inventario
70                Call InvBanco(1).SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), _
                      .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), _
                      .MinDef(i), .Valor(i), .ItemName(i))
80            End With
90        Next i

100       For i = 1 To MAX_BANCOINVENTORY_SLOTS
110           With UserBancoInventory(i)
120               Call InvBanco(0).SetItem(i, .ObjIndex, .Amount, .Equipped, _
                      .GrhIndex, .ObjType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, _
                      .Name)
130           End With
140       Next i

          'Set state and show form
150       Comerciando = True

160       frmBancoObj.lblUserGld.Caption = BankGold

170       frmBancoObj.Show , frmMain
End Sub

''
' Handles the UserCommerceInit message.

Private Sub HandleUserCommerceInit()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
          Dim i      As Long

          'Remove packet ID
10        Call incomingData.ReadByte
20        TradingUserName = incomingData.ReadASCIIString


#If Wgl = 0 Then
          ' Initialize commerce inventories
30        Call InvComUsu.Initialize(DirectDraw, frmComerciarUsu.picInvComercio, _
              Inventario.MaxObjs)
40        Call InvOfferComUsu(0).Initialize(DirectDraw, _
              frmComerciarUsu.picInvOfertaProp, INV_OFFER_SLOTS)
50        Call InvOfferComUsu(1).Initialize(DirectDraw, _
              frmComerciarUsu.picInvOfertaOtro, INV_OFFER_SLOTS)
60        Call InvOroComUsu(0).Initialize(DirectDraw, frmComerciarUsu.picInvOroProp, _
              INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
70        Call InvOroComUsu(1).Initialize(DirectDraw, _
              frmComerciarUsu.picInvOroOfertaProp, INV_GOLD_SLOTS, , TilePixelWidth * 2, _
              TilePixelHeight, TilePixelWidth / 2)
80        Call InvOroComUsu(2).Initialize(DirectDraw, _
              frmComerciarUsu.picInvOroOfertaOtro, INV_GOLD_SLOTS, , TilePixelWidth * 2, _
              TilePixelHeight, TilePixelWidth / 2)
              
              
#Else
          ' Initialize commerce inventories
30        Call InvComUsu.Initialize(frmComerciarUsu.picInvComercio, _
              Inventario.MaxObjs)
40        Call InvOfferComUsu(0).Initialize( _
              frmComerciarUsu.picInvOfertaProp, INV_OFFER_SLOTS)
50        Call InvOfferComUsu(1).Initialize( _
              frmComerciarUsu.picInvOfertaOtro, INV_OFFER_SLOTS)
60        Call InvOroComUsu(0).Initialize(frmComerciarUsu.picInvOroProp, _
              INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
70        Call InvOroComUsu(1).Initialize( _
              frmComerciarUsu.picInvOroOfertaProp, INV_GOLD_SLOTS, , TilePixelWidth * 2, _
              TilePixelHeight, TilePixelWidth / 2)
80        Call InvOroComUsu(2).Initialize( _
              frmComerciarUsu.picInvOroOfertaOtro, INV_GOLD_SLOTS, , TilePixelWidth * 2, _
              TilePixelHeight, TilePixelWidth / 2)

#End If

          'Fill user inventory
90        For i = 1 To MAX_INVENTORY_SLOTS
100           If Inventario.ObjIndex(i) <> 0 Then
110               With Inventario
120                   Call InvComUsu.SetItem(i, .ObjIndex(i), .Amount(i), _
                          .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), _
                          .MaxDef(i), .MinDef(i), .Valor(i), .ItemName(i))
130               End With
140           End If
150       Next i

          ' Inventarios de oro
160       Call InvOroComUsu(0).SetItem(1, ORO_INDEX, UserGLD, 0, ORO_GRH, 0, 0, 0, 0, _
              0, 0, "Oro")
170       Call InvOroComUsu(1).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, _
              "Oro")
180       Call InvOroComUsu(2).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, _
              "Oro")


          'Set state and show form
190       Comerciando = True
200       Call frmComerciarUsu.Show(vbModeless, frmMain)
End Sub

''
' Handles the UserCommerceEnd message.

Private Sub HandleUserCommerceEnd()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        Set InvComUsu = Nothing
30        Set InvOroComUsu(0) = Nothing
40        Set InvOroComUsu(1) = Nothing
50        Set InvOroComUsu(2) = Nothing
60        Set InvOfferComUsu(0) = Nothing
70        Set InvOfferComUsu(1) = Nothing

          'Destroy the form and reset the state
80        Unload frmComerciarUsu
90        Comerciando = False
End Sub

''
' Handles the UserOfferConfirm message.
Private Sub HandleUserOfferConfirm()
      '***************************************************
      'Author: ZaMa
      'Last Modification: 14/12/2009
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        With frmComerciarUsu
              ' Now he can accept the offer or reject it
30            .HabilitarAceptarRechazar True

40            .PrintCommerceMsg TradingUserName & " ha confirmado su oferta!", _
                  FontTypeNames.FONTTYPE_admin
50        End With

End Sub

''
' Handles the ShowBlacksmithForm message.

Private Sub HandleShowBlacksmithForm()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
30            Call WriteCraftBlacksmith(MacroBltIndex)
40        Else
50            frmHerrero.Show , frmMain
60        End If
End Sub

''
' Handles the ShowCarpenterForm message.

Private Sub HandleShowCarpenterForm()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
30            Call WriteCraftCarpenter(MacroBltIndex)
40        Else
50            frmCarp.Show , frmMain
60        End If
End Sub

''
' Handles the NPCSwing message.

Private Sub HandleNPCSwing()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, _
              0, True, False, True)
End Sub

''
' Handles the NPCKillUser message.

Private Sub HandleNPCKillUser()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, _
              True, False, True)
End Sub

''
' Handles the BlockedWithShieldUser message.

Private Sub HandleBlockedWithShieldUser()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, _
              0, True, False, True)
End Sub

''
' Handles the BlockedWithShieldOther message.

Private Sub HandleBlockedWithShieldOther()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, _
              255, 0, 0, True, False, True)
End Sub

''
' Handles the UserSwing message.

Private Sub HandleUserSwing()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, _
              True, False, True)
End Sub

''
' Handles the SafeModeOn message.

Private Sub HandleSafeModeOn()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

          '  Call frmMain.DibujarSeguro
          'Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, False)
20        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, _
              True, False, True)
          ' frmMain.IconoSeg.Caption = ""
End Sub

''
' Handles the SafeModeOff message.

Private Sub HandleSafeModeOff()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

          'Call frmMain.DesDibujarSeguro
20        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, _
              True, False, True)
          'frmMain.IconoSeg.Caption = "X"
End Sub

Private Sub hAndlEdRAGOn()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        Call frmMain.ControlSM(eSMType.DragMode, True)
End Sub

''
' Handles the SafeModeOff message.

Private Sub handleDragOff()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        Call frmMain.ControlSM(eSMType.DragMode, False)
End Sub

''
' Handles the ResuscitationSafeOff message.

Private Sub HandleResuscitationSafeOff()
      '***************************************************
      'Author: Rapsodius
      'Creation date: 10/10/07
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        Call frmMain.ControlSM(eSMType.sResucitation, False)
End Sub

''
' Handles the ResuscitationSafeOn message.

Private Sub HandleResuscitationSafeOn()
      '***************************************************
      'Author: Rapsodius
      'Creation date: 10/10/07
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        Call frmMain.ControlSM(eSMType.sResucitation, True)
End Sub

''
' Handles the NobilityLost message.

Private Sub HandleNobilityLost()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, _
              False, False, True)
End Sub

''
' Handles the CantUseWhileMeditating message.

Private Sub HandleCantUseWhileMeditating()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, _
              False, False, True)
End Sub

''
' Handles the UpdateSta message.
Private Sub HandleUpdateSta()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Check packet is complete
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          'Get data and update form
60        UserMinSTA = incomingData.ReadInteger()
70        frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 97)

          Dim tmpShadow As Long
80        For tmpShadow = 0 To 4
90            frmMain.lblEnergia(tmpShadow) = UserMinSTA & "/" & UserMaxSTA
100       Next tmpShadow


End Sub

''
' Handles the UpdateMana message.

Private Sub HandleUpdateMana()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Check packet is complete
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          'Get data and update form
60        UserMinMAN = incomingData.ReadInteger()

70        If UserMaxMAN > 0 Then
80            frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / _
                  100)) * 97)
90        Else
100           frmMain.MANShp.Width = 7
110       End If

          Dim tmpShadow As Long
120       For tmpShadow = 0 To 4
130           frmMain.lblMana(tmpShadow) = UserMinMAN & "/" & UserMaxMAN
140       Next tmpShadow
End Sub

''
' Handles the UpdateHP message.

Private Sub HandleUpdateHP()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Check packet is complete
   On Error GoTo HandleUpdateHP_Error

10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          'Get data and update form
60        UserMinHP = incomingData.ReadInteger()

          Dim TempInt As Integer
          TempInt = (((UserMinHP / 100) / (UserMaxHP / 100)) * 97)
          
          If TempInt < 0 Then TempInt = 0
          
          If TempInt <= 0 Then TempInt = 7
70        frmMain.Hpshp.Width = TempInt

          'Is the user alive??
80        If UserMinHP = 0 Then
90            UserEstado = 1
100           If frmMain.TrainingMacro Then frmMain.DesactivarMacroHechizos
110           If frmMain.macrotrabajo Then frmMain.DesactivarMacroTrabajo
120       Else
130           UserEstado = 0
140       End If

          Dim tmpShadow As Long
150       For tmpShadow = 0 To 4
160           frmMain.lblVida(tmpShadow) = UserMinHP & "/" & UserMaxHP
170       Next tmpShadow

   On Error GoTo 0
   Exit Sub

HandleUpdateHP_Error:

   ' LogError "Error " & Err.number & " (" & Err.Description & ") in procedure HandleUpdateHP of Módulo Protocol in line " & Erl

End Sub

''
' Handles the UpdateGold message.

Private Sub HandleUpdateGold()
      '***************************************************
      'Autor: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 08/14/07
      'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
      '- 08/14/07: Added GldLbl color variation depending on User Gold and Level
      '***************************************************
      'Check packet is complete
10        If incomingData.Length < 5 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          'Get data and update form
60        UserGLD = incomingData.ReadLong()

70        If UserGLD >= CLng(UserLvl) * 10000 Then
              'Changes color
80            frmMain.GldLbl.ForeColor = &HFF&    'Red
90        Else
              'Changes color
100           frmMain.GldLbl.ForeColor = &HFF&  'Yellow
110       End If

120       frmMain.GldLbl.Caption = UserGLD
End Sub

''
' Handles the UpdateBankGold message.

Private Sub HandleUpdateBankGold()
      '***************************************************
      'Autor: ZaMa
      'Last Modification: 14/12/2009
      '
      '***************************************************
      'Check packet is complete
10        If incomingData.Length < 5 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

60        frmBancoObj.lblUserGld.Caption = incomingData.ReadLong

End Sub

''
' Handles the UpdateExp message.

Private Sub HandleUpdateExp()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      
      
On Error GoTo ErrHandler
      'Check packet is complete
10        If incomingData.Length < 5 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          'Get data and update form
60        UserExp = incomingData.ReadLong()

          If UserPasarNivel > 0 Then
70            frmMain.lblporclvl(0).Caption = "(" & Round(CDbl(UserExp) * CDbl(100) / _
                  CDbl(UserPasarNivel), 2) & "%)"
80            frmMain.lblporclvl(1).Caption = "(" & Round(CDbl(UserExp) * CDbl(100) / _
                  CDbl(UserPasarNivel), 2) & "%)"
90            frmMain.lblporclvl(2).Caption = "(" & Round(CDbl(UserExp) * CDbl(100) / _
                  CDbl(UserPasarNivel), 2) & "%)"
100           frmMain.lblporclvl(3).Caption = "(" & Round(CDbl(UserExp) * CDbl(100) / _
                  CDbl(UserPasarNivel), 2) & "%)"
110           frmMain.lblporclvl(4).Caption = "(" & Round(CDbl(UserExp) * CDbl(100) / _
                  CDbl(UserPasarNivel), 2) & "%)"
                  
          Else
             frmMain.lblporclvl(0).Caption = "0%"
             frmMain.lblporclvl(1).Caption = "0%"
             frmMain.lblporclvl(2).Caption = "0%"
            frmMain.lblporclvl(3).Caption = "0%"
            frmMain.lblporclvl(4).Caption = "0%"
          End If
          
Exit Sub
ErrHandler:
    LogError "Hubo error linea " & Erl & " procedimiento UpdateExp"

End Sub

''
' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenghtAndDexterity()
      '***************************************************
      'Author: Budi
      'Last Modification: 11/26/09
      '***************************************************
      'Check packet is complete
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          'Get data and update form
60        UserFuerza = incomingData.ReadByte
70        UserAgilidad = incomingData.ReadByte
80        frmMain.lblStrg.Caption = UserFuerza
90        frmMain.lblDext.Caption = UserAgilidad
100       frmMain.lblStrg.ForeColor = getStrenghtColor(UserFuerza)
110       frmMain.lblDext.ForeColor = getDexterityColor(UserAgilidad)
End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenght()
      '***************************************************
      'Author: Budi
      'Last Modification: 11/26/09
      '***************************************************
      'Check packet is complete
10        If incomingData.Length < 2 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          'Get data and update form
60        UserFuerza = incomingData.ReadByte
70        frmMain.lblStrg.Caption = UserFuerza
80        frmMain.lblStrg.ForeColor = getStrenghtColor(UserFuerza)
End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateDexterity()
      '***************************************************
      'Author: Budi
      'Last Modification: 11/26/09
      '***************************************************
      'Check packet is complete
10        If incomingData.Length < 2 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          'Get data and update form
60        UserAgilidad = incomingData.ReadByte
70        frmMain.lblDext.Caption = UserAgilidad
80        frmMain.lblDext.ForeColor = getDexterityColor(UserAgilidad)
End Sub

''
' Handles the ChangeMap message.
Private Sub HandleChangeMap()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 5 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

60        UserMap = incomingData.ReadInteger()
70        MapaActual = incomingData.ReadASCIIString()

          'TODO: Once on-the-fly editor is implemented check for map version before loading....
          'For now we just drop it
80        Call incomingData.ReadInteger
          'If FileExist(DirMapas & "Mapa" & UserMap & ".map", vbNormal) Then
90        Call SwitchMap(UserMap)
End Sub

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdate()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          'Remove char from old position
60        If MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex Then
70            MapData(UserPos.X, UserPos.Y).CharIndex = 0

              #If Wgl = 1 Then
                    Call g_Swarm.RemoveDynamic(UserCharIndex)
              #End If
80        End If

          'Set new pos
90        UserPos.X = incomingData.ReadByte()
100       UserPos.Y = incomingData.ReadByte()

          'Set char
110       MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
120       charlist(UserCharIndex).Pos = UserPos

          #If Wgl = 1 Then
                Dim RangeX As Single, RangeY As Single
                Call GetCharacterDimension(UserCharIndex, RangeX, RangeY)
                Call g_Swarm.InsertDynamic(UserCharIndex, 5, UserPos.X, UserPos.Y, RangeX, RangeY)
          #End If
          
          'Are we under a roof?
130       bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, _
              UserPos.Y).Trigger = 2 Or MapData(UserPos.X, UserPos.Y).Trigger = 4, True, _
              False)

          'Update pos label
140       frmMain.Coord.Caption = "Mapa " & UserMap & " [" & UserPos.X & "," & _
              UserPos.Y & "]"
150       frmMain.lblmapaname.Caption = MapaActual
End Sub

''
' Handles the NPCHitUser message.

Private Sub HandleNPCHitUser()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 4 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

60        Select Case incomingData.ReadByte()
          Case bCabeza
70            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & _
                  CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
80        Case bBrazoIzquierdo
90            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & _
                  CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
100       Case bBrazoDerecho
110           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & _
                  CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
120       Case bPiernaIzquierda
130           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & _
                  CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
140       Case bPiernaDerecha
150           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & _
                  CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
160       Case bTorso
170           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & _
                  CStr(incomingData.ReadInteger() & "!!"), 255, 0, 0, True, False, True)
180       End Select
End Sub

''
' Handles the UserHitNPC message.

Private Sub HandleUserHitNPC()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 5 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

60        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & _
              CStr(incomingData.ReadLong()) & MENSAJE_2, 255, 0, 0, True, False, True)
End Sub

''
' Handles the UserAttackedSwing message.

Private Sub HandleUserAttackedSwing()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

60        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & _
              charlist(incomingData.ReadInteger()).Nombre & MENSAJE_ATAQUE_FALLO, 255, 0, _
              0, True, False, True)
End Sub

''
' Handles the UserHittingByUser message.

Private Sub HandleUserHittedByUser()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 6 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          Dim attacker As String

60        attacker = charlist(incomingData.ReadInteger()).Nombre

70        Select Case incomingData.ReadByte
          Case bCabeza
80            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & _
                  MENSAJE_RECIVE_IMPACTO_CABEZA & CStr(incomingData.ReadInteger() & _
                  MENSAJE_2), 255, 0, 0, True, False, True)
90        Case bBrazoIzquierdo
100           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & _
                  MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & CStr(incomingData.ReadInteger() & _
                  MENSAJE_2), 255, 0, 0, True, False, True)
110       Case bBrazoDerecho
120           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & _
                  MENSAJE_RECIVE_IMPACTO_BRAZO_DER & CStr(incomingData.ReadInteger() & _
                  MENSAJE_2), 255, 0, 0, True, False, True)
130       Case bPiernaIzquierda
140           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & _
                  MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & CStr(incomingData.ReadInteger() & _
                  MENSAJE_2), 255, 0, 0, True, False, True)
150       Case bPiernaDerecha
160           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & _
                  MENSAJE_RECIVE_IMPACTO_PIERNA_DER & CStr(incomingData.ReadInteger() & _
                  MENSAJE_2), 255, 0, 0, True, False, True)
170       Case bTorso
180           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & _
                  MENSAJE_RECIVE_IMPACTO_TORSO & CStr(incomingData.ReadInteger() & _
                  MENSAJE_2), 255, 0, 0, True, False, True)
190       End Select
End Sub

''
' Handles the UserHittedUser message.

Private Sub HandleUserHittedUser()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 6 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          Dim Victim As String

60        Victim = charlist(incomingData.ReadInteger()).Nombre

70        Select Case incomingData.ReadByte
          Case bCabeza
80            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & _
                  Victim & MENSAJE_PRODUCE_IMPACTO_CABEZA & _
                  CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, _
                  True)
90        Case bBrazoIzquierdo
100           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & _
                  Victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & _
                  CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, _
                  True)
110       Case bBrazoDerecho
120           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & _
                  Victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & _
                  CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, _
                  True)
130       Case bPiernaIzquierda
140           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & _
                  Victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & _
                  CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, _
                  True)
150       Case bPiernaDerecha
160           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & _
                  Victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & _
                  CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, _
                  True)
170       Case bTorso
180           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & _
                  Victim & MENSAJE_PRODUCE_IMPACTO_TORSO & _
                  CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, _
                  True)
190       End Select
End Sub

''
' Handles the ChatOverHead message.

Private Sub HandleChatOverHead()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 8 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim chat   As String
          Dim CharIndex As Integer
          Dim r      As Byte
          Dim g      As Byte
          Dim b      As Byte

80        chat = buffer.ReadASCIIString()
90        CharIndex = buffer.ReadInteger()

100       r = buffer.ReadByte()
110       g = buffer.ReadByte()
120       b = buffer.ReadByte()

          'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
130       If charlist(CharIndex).Active Then Call Dialogos.CreateDialog(Trim$(chat), _
              CharIndex, RGB(r, g, b))

          'If we got here then packet is complete, copy data back to original queue
140       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
150       error = Err.number
160       On Error GoTo 0

          'Destroy auxiliar buffer
170       Set buffer = Nothing

180       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleConsoleMessage()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 4 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim chat   As String
          Dim FontIndex As Integer
          Dim str    As String
          Dim r      As Byte
          Dim g      As Byte
          Dim b      As Byte
          Dim vbcrlf As Boolean

80        chat = buffer.ReadASCIIString()
90        FontIndex = buffer.ReadByte()
            
        If Right$(chat, 1) = "`" Then
            vbcrlf = False
            chat = Left$(chat, Len(chat) - 1)
        Else
            vbcrlf = True
        End If
        
100       If InStr(1, chat, "~") Then
110           str = ReadField(2, chat, 126)
120           If Val(str) > 255 Then
130               r = 255
140           Else
150               r = Val(str)
160           End If

170           str = ReadField(3, chat, 126)
180           If Val(str) > 255 Then
190               g = 255
200           Else
210               g = Val(str)
220           End If

230           str = ReadField(4, chat, 126)
240           If Val(str) > 255 Then
250               b = 255
260           Else
270               b = Val(str)
280           End If

290           Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - _
                  1), r, g, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, _
                  126)) <> 0, vbcrlf)
300       Else
            
310           With FontTypes(FontIndex)
320               Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, _
                      .bold, .italic, vbcrlf)
330           End With

              ' Para no perder el foco cuando chatea por party
340           If FontIndex = FontTypeNames.FONTTYPE_PARTY Then
350               If MirandoParty Then frmMain.SendTxt.SetFocus
360           End If
370       End If
          '    Call checkText(chat)
          'If we got here then packet is complete, copy data back to original queue
380       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
390       error = Err.number
400       On Error GoTo 0

          'Destroy auxiliar buffer
410       Set buffer = Nothing

420       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the GuildChat message.

Private Sub HandleGuildChat()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 04/07/08 (NicoNZ)
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim chat   As String
          Dim str    As String
          Dim r      As Byte
          Dim g      As Byte
          Dim b      As Byte
          Dim tmp    As Integer
          Dim Cont   As Integer


80        chat = buffer.ReadASCIIString()

90        If Not DialogosClanes.Activo Then
100           If InStr(1, chat, "~") Then
110               str = ReadField(2, chat, 126)
120               If Val(str) > 255 Then
130                   r = 255
140               Else
150                   r = Val(str)
160               End If

170               str = ReadField(3, chat, 126)
180               If Val(str) > 255 Then
190                   g = 255
200               Else
210                   g = Val(str)
220               End If

230               str = ReadField(4, chat, 126)
240               If Val(str) > 255 Then
250                   b = 255
260               Else
270                   b = Val(str)
280               End If

290               Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, _
                      "~") - 1), r, g, b, Val(ReadField(5, chat, 126)) <> 0, _
                      Val(ReadField(6, chat, 126)) <> 0)
300           Else
310               With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
320                   Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, _
                          .bold, .italic)
330               End With
340           End If
350       Else
360           Call DialogosClanes.PushBackText(ReadField(1, chat, 126))
370       End If

          'If we got here then packet is complete, copy data back to original queue
380       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
390       error = Err.number
400       On Error GoTo 0

          'Destroy auxiliar buffer
410       Set buffer = Nothing

420       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleCommerceChat()
      '***************************************************
      'Author: ZaMa
      'Last Modification: 03/12/2009
      '
      '***************************************************
10        If incomingData.Length < 4 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim chat   As String
          Dim FontIndex As Integer
          Dim str    As String
          Dim r      As Byte
          Dim g      As Byte
          Dim b      As Byte

80        chat = buffer.ReadASCIIString()
90        FontIndex = buffer.ReadByte()

100       If InStr(1, chat, "~") Then
110           str = ReadField(2, chat, 126)
120           If Val(str) > 255 Then
130               r = 255
140           Else
150               r = Val(str)
160           End If

170           str = ReadField(3, chat, 126)
180           If Val(str) > 255 Then
190               g = 255
200           Else
210               g = Val(str)
220           End If

230           str = ReadField(4, chat, 126)
240           If Val(str) > 255 Then
250               b = 255
260           Else
270               b = Val(str)
280           End If

290           Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, Left$(chat, _
                  InStr(1, chat, "~") - 1), r, g, b, Val(ReadField(5, chat, 126)) <> 0, _
                  Val(ReadField(6, chat, 126)) <> 0)
300       Else
310           With FontTypes(FontIndex)
320               Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, chat, .red, _
                      .green, .blue, .bold, .italic)
330           End With
340       End If

          'If we got here then packet is complete, copy data back to original queue
350       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
360       error = Err.number
370       On Error GoTo 0

          'Destroy auxiliar buffer
380       Set buffer = Nothing

390       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the ShowMessageBox message.

Private Sub HandleShowMessageBox()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

80        frmMensaje.msg.Caption = buffer.ReadASCIIString()
90        frmMensaje.Show

          'If we got here then packet is complete, copy data back to original queue
100       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
110       error = Err.number
120       On Error GoTo 0

          'Destroy auxiliar buffer
130       Set buffer = Nothing

140       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

60        UserIndex = incomingData.ReadInteger()
          
End Sub

''
' Handles the UserCharIndexInServer message.

Private Sub HandleUserCharIndexInServer()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

60        UserCharIndex = incomingData.ReadInteger()
70        UserPos = charlist(UserCharIndex).Pos

          'Are we under a roof?
80        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, _
              UserPos.Y).Trigger = 2 Or MapData(UserPos.X, UserPos.Y).Trigger = 4, True, _
              False)

90        frmMain.Coord.Caption = "Mapa " & UserMap & " [" & UserPos.X & "," & _
              UserPos.Y & "]"
100       frmMain.lblmapaname.Caption = MapaActual


End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
   On Error GoTo HandleCharacterCreate_Error

10        If incomingData.Length < 24 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim CharIndex As Integer
          Dim Body   As Integer
          Dim Head   As Integer
          Dim Heading As E_Heading
          Dim X      As Byte
          Dim Y      As Byte
          Dim weapon As Integer
          Dim shield As Integer
          Dim helmet As Integer
          Dim privs  As Integer
          Dim NickColor As Byte

80        CharIndex = buffer.ReadInteger()
90        Body = buffer.ReadInteger()
100       Head = buffer.ReadInteger()
110       Heading = buffer.ReadByte()
120       X = buffer.ReadByte()
130       Y = buffer.ReadByte()
140       weapon = buffer.ReadInteger()
150       shield = buffer.ReadInteger()
160       helmet = buffer.ReadInteger()


170       With charlist(CharIndex)
180           Call SetCharacterFx(CharIndex, buffer.ReadInteger(), _
                  buffer.ReadInteger())

190           .Nombre = buffer.ReadASCIIString()
200           If (CharIndex = UserCharIndex) Then
                  Dim s_Pos As Integer
210               s_Pos = InStr(1, .Nombre, "<")

                  'si encontramos el "<" es porque tiene clan
220               If (s_Pos <> 0) Then
                      Dim Guild_Name As String
                      Dim char_Name As String

230                   char_Name = Left$(.Nombre, (s_Pos - 1))
240                   Guild_Name = mid$(.Nombre, (s_Pos + 1))

                      'ACA PONÉ EL NOMBRE DE TU LABEL, CAPÁS QUE ES ASÍ.
                      Dim i As Long
250                   For i = 0 To 4
260                       frmMain.lblClan(i).Caption = "<" & Guild_Name
270                   Next i
                        
                    TieneClan = True
                        
280               ElseIf (s_Pos = 0) Then
290                   For i = 0 To 4
300                       frmMain.lblClan(i).Caption = vbNullString
310                   Next i
320               End If
330           End If
              
340           NickColor = buffer.ReadByte()

350           If (NickColor And eNickColor.ieCriminal) <> 0 Then
360               .Criminal = 1
370           Else
380               .Criminal = 0
390           End If
              
400           If (NickColor And eNickColor.ieTeamUno) <> 0 Then
410               .Team = 1
420           ElseIf (NickColor And eNickColor.ieTeamDos) <> 0 Then
430               .Team = 2
440           Else
450               .Team = 0
460           End If
              
470           .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
              
480           privs = buffer.ReadByte()

490           If privs <> 0 Then
                  'If the player belongs to a council AND is an admin, only whos as an admin
500               If (privs And PlayerType.ChaosCouncil) <> 0 And (privs And _
                      PlayerType.User) = 0 Then
510                   privs = privs Xor PlayerType.ChaosCouncil
520               End If

530               If (privs And PlayerType.RoyalCouncil) <> 0 And (privs And _
                      PlayerType.User) = 0 Then
540                   privs = privs Xor PlayerType.RoyalCouncil
550               End If

                  'If the player is a RM, ignore other flags
560               If privs And PlayerType.RoleMaster Then
570                   privs = PlayerType.RoleMaster
580               End If

                  'Log2 of the bit flags sent by the server gives our numbers ^^
590               .priv = Log(privs) / Log(2)
600           Else
610               .priv = 0
620           End If
630       End With

640       Call MakeChar(CharIndex, Body, Head, Heading, X, Y, weapon, shield, helmet)

650       Call RefreshAllChars

          'If we got here then packet is complete, copy data back to original queue
660       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
670       error = Err.number
680       On Error GoTo 0

          'Destroy auxiliar buffer
690       Set buffer = Nothing

700       If error <> 0 Then Err.Raise error

   On Error GoTo 0
   Exit Sub

HandleCharacterCreate_Error:

    LogError "Error " & Err.number & " (" & Err.Description & ") in procedure HandleCharacterCreate of Módulo Protocol in line " & Erl
End Sub

Private Sub HandleCharacterChangeNick()
      '***************************************************
      'Author: Budi
      'Last Modification: 07/23/09
      '
      '***************************************************
10        If incomingData.Length < 5 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet id
50        Call incomingData.ReadByte
          Dim CharIndex As Integer
60        CharIndex = incomingData.ReadInteger
        
70        charlist(CharIndex).Nombre = incomingData.ReadASCIIString

End Sub

''
' Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          Dim CharIndex As Integer

60        CharIndex = incomingData.ReadInteger()

70        Call EraseChar(CharIndex)
80        Call RefreshAllChars
End Sub

''
' Handles the CharacterMove message.

Private Sub HandleCharacterMove()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
   On Error GoTo HandleCharacterMove_Error

10        If incomingData.Length < 5 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          Dim CharIndex As Integer
          Dim X      As Byte
          Dim Y      As Byte

60        CharIndex = incomingData.ReadInteger()
70        X = incomingData.ReadByte()
80        Y = incomingData.ReadByte()

90        With charlist(CharIndex)
100           If .FxIndex >= 40 And .FxIndex <= 49 Then   'If it's meditating, we remove the FX
110               .FxIndex = 0
120           End If

              ' Play steps sounds if the user is not an admin of any kind
130           If .priv <> 1 And .priv <> 2 And .priv <> 3 And .priv <> 5 And .priv <> _
                  25 Then
140               Call DoPasosFx(CharIndex)
150           End If
160       End With

170       If X <= 0 Or Y <= 0 Then Exit Sub
          
180       Call MoveCharbyPos(CharIndex, X, Y)

190       Call RefreshAllChars

   On Error GoTo 0
   Exit Sub

HandleCharacterMove_Error:

    'LogError "Error " & Err.number & " (" & Err.Description & ") in procedure HandleCharacterMove of Módulo Protocol in line " & Erl
End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMove()

10        If incomingData.Length < 2 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          Dim Direccion As Byte

60        Direccion = incomingData.ReadByte()

70        Call MoveCharbyHead(UserCharIndex, Direccion)
80        Call MoveScreen(Direccion)

90        Call RefreshAllChars
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 25/08/2009
      '25/08/2009: ZaMa - Changed a variable used incorrectly.
      '***************************************************
10        If incomingData.Length < 18 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          Dim CharIndex As Integer
          Dim TempInt As Integer
          Dim headIndex As Integer

60        CharIndex = incomingData.ReadInteger()

70        With charlist(CharIndex)
80            TempInt = incomingData.ReadInteger()

90            If TempInt < LBound(BodyData()) Or TempInt > UBound(BodyData()) Then
100               .Body = BodyData(0)
110               .iBody = 0
120           Else
130               .Body = BodyData(TempInt)
140               .iBody = TempInt
150           End If


160           headIndex = incomingData.ReadInteger()

170           If headIndex < LBound(HeadData()) Or headIndex > UBound(HeadData()) Then
180               .Head = HeadData(0)
190               .iHead = 0
200           Else
210               .Head = HeadData(headIndex)
220               .iHead = headIndex
230           End If

240           .muerto = (headIndex = CASPER_HEAD)

250           .Heading = incomingData.ReadByte()

260           TempInt = incomingData.ReadInteger()
270           If TempInt <> 0 Then .Arma = WeaponAnimData(TempInt)

280           TempInt = incomingData.ReadInteger()
290           If TempInt <> 0 Then .Escudo = ShieldAnimData(TempInt)

300           TempInt = incomingData.ReadInteger()
310           If TempInt <> 0 Then .Casco = CascoAnimData(TempInt)

320           Call SetCharacterFx(CharIndex, incomingData.ReadInteger(), _
                  incomingData.ReadInteger())
                  
                  
              #If Wgl = 1 Then
                    Dim RangeX As Single, RangeY As Single
                    Call GetCharacterDimension(CharIndex, RangeX, RangeY)
                    
                    Call g_Swarm.Update(CharIndex, RangeX, RangeY)
              #End If
330       End With

340       Call RefreshAllChars
End Sub

''
' Handles the ObjectCreate message.

Private Sub HandleObjectCreate()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 5 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          Dim X      As Byte
          Dim Y      As Byte

60        X = incomingData.ReadByte()
70        Y = incomingData.ReadByte()

          #If Wgl = 1 Then
            ' RTREE
            If (MapData(X, Y).ObjGrh.GrhIndex <> 0) Then
                With GrhData(MapData(X, Y).ObjGrh.GrhIndex)
                    Call g_Swarm.Remove(4, X, Y, .TileWidth, .TileHeight)
                End With
            End If
          #End If
          
80        MapData(X, Y).ObjGrh.GrhIndex = incomingData.ReadInteger()

          #If Wgl = 1 Then
            ' RTREE
            If (MapData(X, Y).ObjGrh.GrhIndex <> 0) Then
                With GrhData(MapData(X, Y).ObjGrh.GrhIndex)
                    Call g_Swarm.Insert(4, X, Y, .TileWidth, .TileHeight)
                End With
            End If
          #End If
          
90        Call InitGrh(MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex)
End Sub

''
' Handles the ObjectDelete message.

Private Sub HandleObjectDelete()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          Dim X      As Byte
          Dim Y      As Byte

60        X = incomingData.ReadByte()
70        Y = incomingData.ReadByte()


          #If Wgl = 1 Then
                With GrhData(MapData(X, Y).ObjGrh.GrhIndex)
                  Call g_Swarm.Remove(4, X, Y, .TileWidth, .TileHeight)
                End With
          #End If
          
80        MapData(X, Y).ObjGrh.GrhIndex = 0
End Sub

''
' Handles the BlockPosition message.

Private Sub HandleBlockPosition()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 4 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          Dim X      As Byte
          Dim Y      As Byte

60        X = incomingData.ReadByte()
70        Y = incomingData.ReadByte()

80        If incomingData.ReadBoolean() Then
90            MapData(X, Y).Blocked = 1
100       Else
110           MapData(X, Y).Blocked = 0
120       End If
End Sub

''
' Handles the PlayMIDI message.

Private Sub HandlePlayMIDI()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 4 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          Dim currentMidi As Byte

          'Remove packet ID
50        Call incomingData.ReadByte

60        currentMidi = incomingData.ReadByte()

70        If currentMidi Then
80            Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", _
                  incomingData.ReadInteger())
90        Else
              'Remove the bytes to prevent errors
100           Call incomingData.ReadInteger
110       End If
End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWave()
      '***************************************************
      'Autor: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 08/14/07
      'Last Modified by: Rapsodius
      'Added support for 3D Sounds.
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          Dim wave   As Byte
          Dim srcX   As Byte
          Dim srcY   As Byte

60        wave = incomingData.ReadByte()
70        srcX = incomingData.ReadByte()
80        srcY = incomingData.ReadByte()

90        Call Audio.PlayWave(CStr(wave) & ".wav", srcX, srcY)
End Sub

''
' Handles the GuildList message.

Private Sub HandleGuildList()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

80        With frmGuildAdm
              'Clear guild's list
90            .guildslist.Clear

100           GuildNames = Split(buffer.ReadASCIIString(), SEPARATOR)

              Dim i  As Long
110           For i = 0 To UBound(GuildNames())
120               Call .guildslist.AddItem(GuildNames(i))
130           Next i

              'If we got here then packet is complete, copy data back to original queue
140           Call incomingData.CopyBuffer(buffer)

150           .Show vbModeless, frmMain
160       End With

ErrHandler:
          Dim error  As Long
170       error = Err.number
180       On Error GoTo 0

          'Destroy auxiliar buffer
190       Set buffer = Nothing

200       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the AreaChanged message.

Private Sub HandleAreaChanged()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          Dim X      As Byte
          Dim Y      As Byte

60        X = incomingData.ReadByte()
70        Y = incomingData.ReadByte()

80        Call CambioDeArea(X, Y)
End Sub

''
' Handles the PauseToggle message.

Private Sub HandlePauseToggle()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        pausa = Not pausa
End Sub

Private Sub HandleUserInEvent()
10        Call incomingData.ReadByte

20        UserEvento = Not UserEvento
End Sub

''
' Handles the CreateFX message.

Private Sub HandleCreateFX()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 7 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          Dim CharIndex As Integer
          Dim fX     As Integer
          Dim Loops  As Integer

60        CharIndex = incomingData.ReadInteger()
70        fX = incomingData.ReadInteger()
80        Loops = incomingData.ReadInteger()

90        Call SetCharacterFx(CharIndex, fX, Loops)
End Sub

''
' Handles the UpdateUserStats message.
Private Sub HandleUpdateUserStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

10    On Error GoTo ErrHandler
20        If incomingData.Length < 26 Then
30      Err.Raise incomingData.NotEnoughDataErrCode
40      Exit Sub
50        End If

    'Remove packet ID
60        Call incomingData.ReadByte
    Dim CharIndex As Integer
70
80        UserMaxHP = incomingData.ReadInteger()
90        UserMinHP = incomingData.ReadInteger()
100       UserMaxMAN = incomingData.ReadInteger()
110       UserMinMAN = incomingData.ReadInteger()
120       UserMaxSTA = incomingData.ReadInteger()
130       UserMinSTA = incomingData.ReadInteger()
140       UserGLD = incomingData.ReadLong()
150       UserLvl = incomingData.ReadByte()
160       UserPasarNivel = incomingData.ReadLong()
170       UserExp = incomingData.ReadLong()
180       UserOcu = incomingData.ReadByte()
190       Iscombate = incomingData.ReadBoolean()
200       CharIndex = incomingData.ReadInteger()
210
220       With charlist(CharIndex)
230     If (CharIndex = UserCharIndex) Then
            Dim s_Pos As Integer
240
250         s_Pos = InStr(1, .Nombre, "<")
260
            'si encontramos el "<" es porque tiene clan
270         If (s_Pos <> 0) Then
                Dim Guild_Name As String
                Dim char_Name As String
280
290             char_Name = Left$(.Nombre, (s_Pos - 1))
300             Guild_Name = mid$(.Nombre, (s_Pos + 1))

                'ACA PONÉ EL NOMBRE DE TU LABEL, CAPÁS QUE ES ASÍ.
                Dim i As Long
310             For i = 0 To 4
320                 frmMain.lblClan(i).Caption = "<" & Guild_Name
330             Next i
340         ElseIf (s_Pos = 0) Then
350               For i = 0 To 4
360                 frmMain.lblClan(i).Caption = vbNullString
370             Next i
380         End If
390     End If
400       End With


410   If UserPasarNivel > 0 Then
420     frmMain.lblporclvl(0).Caption = "" & Round(CDbl(UserExp) * CDbl(100) / _
            CDbl(UserPasarNivel)) & "%"
430     frmMain.lblporclvl(1).Caption = "" & Round(CDbl(UserExp) * CDbl(100) / _
            CDbl(UserPasarNivel)) & "%"
440       frmMain.lblporclvl(2).Caption = "" & Round(CDbl(UserExp) * CDbl(100) _
    / CDbl(UserPasarNivel)) & "%"
450     frmMain.lblporclvl(3).Caption = "" & Round(CDbl(UserExp) * CDbl(100) / _
            CDbl(UserPasarNivel)) & "%"
460     frmMain.lblporclvl(4).Caption = "" & Round(CDbl(UserExp) * CDbl(100) / _
            CDbl(UserPasarNivel)) & "%"
470       Else
480     frmMain.lblporclvl(0).Caption = "0%"
490       frmMain.lblporclvl(1).Caption = "0%"
500     frmMain.lblporclvl(2).Caption = "0%"
510     frmMain.lblporclvl(3).Caption = "0%"
520     frmMain.lblporclvl(4).Caption = "0%"
530       End If



540
550       frmMain.GldLbl.Caption = UserGLD
560       frmMain.lvllbl(0).Caption = UserLvl
570       frmMain.lvllbl(1).Caption = UserLvl
580       frmMain.lvllbl(2).Caption = UserLvl
590       frmMain.lvllbl(3).Caption = UserLvl
600       frmMain.lvllbl(4).Caption = UserLvl
    'Stats
610
620       frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 97)
630
640       If UserMaxMAN > 0 Then
650     frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / _
            100)) * 97)
660       Else
670     frmMain.MANShp.Width = 7
680       End If
690
700       frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 97)
710
    Dim tmpShadow As Long
720       For tmpShadow = 0 To 4
730     frmMain.lblEnergia(tmpShadow) = UserMinSTA & "/" & UserMaxSTA
740     frmMain.lblMana(tmpShadow) = UserMinMAN & "/" & UserMaxMAN
750     frmMain.lblVida(tmpShadow) = UserMinHP & "/" & UserMaxHP
760       Next tmpShadow
770
780       If UserMinHP = 0 Then
790     UserEstado = 1
800     If frmMain.TrainingMacro Then frmMain.DesactivarMacroHechizos
810     If frmMain.macrotrabajo Then frmMain.DesactivarMacroTrabajo
820       Else
830     UserEstado = 0
840       End If
850
860       If UserGLD >= CLng(UserLvl) * 100000 Then
        'Changes color
870     frmMain.GldLbl.ForeColor = &HFFFF&    'Red
880       Else
        'Changes color
890     frmMain.GldLbl.ForeColor = &HFF&    'Yellow
900       End If
910
920       Exit Sub

ErrHandler:
930   Call LogError("Error en UpdateUserStats. Número " & Err.number & " Descripción: " & _
    Err.Description & " LINEA: " & Erl)
End Sub


''
' Handles the WorkRequestTarget message.

Private Sub HandleWorkRequestTarget()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 2 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

60        UsingSkill = incomingData.ReadByte()

70        frmMain.MousePointer = 2

80        Select Case UsingSkill
          Case Magia
90            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, _
                  120, 0, 0)
100       Case Pesca
110           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, _
                  120, 0, 0)
120       Case Robar
130           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, _
                  120, 0, 0)
140       Case Talar
150           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, _
                  120, 0, 0)
160       Case Mineria
170           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, _
                  120, 0, 0)
180       Case FundirMetal
190           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, _
                  100, 120, 0, 0)
200       Case Proyectiles
210           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, _
                  100, 120, 0, 0)
220       End Select
End Sub

''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 4 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
50        Call incomingData.ReadByte

          Dim Slot   As Byte
          Dim ObjIndex As Integer
          Dim Name   As String
          Dim Amount As Integer
          Dim Equipped As Boolean
          Dim GrhIndex As Integer
          Dim ObjType As Byte
          Dim MaxHit As Integer
          Dim MinHit As Integer
          Dim MaxDef As Integer
          Dim MinDef As Integer
          Dim value  As Single

60        Slot = incomingData.ReadByte()
70        ObjIndex = incomingData.ReadInteger()
80        Amount = incomingData.ReadInteger()
90        Equipped = incomingData.ReadBoolean()
          
100       If ObjIndex > 0 Then
110           GrhIndex = incomingData.ReadInteger()
120           ObjType = incomingData.ReadByte()
130           MaxHit = incomingData.ReadInteger()
140           MinHit = incomingData.ReadInteger()
150           MaxDef = incomingData.ReadInteger()
160           MinDef = incomingData.ReadInteger
170           value = incomingData.ReadSingle()
              
180           Name = ObjName(ObjIndex).Name
190       End If
          
200       If Equipped Then
210           Select Case ObjType
              Case eOBJType.otWeapon
220               frmMain.lblWeapon(0) = MaxHit
230               frmMain.lblWeapon(1) = MaxHit
240               frmMain.lblWeapon(2) = MaxHit
250               UserWeaponEqpSlot = Slot
260           Case eOBJType.otArmadura
270               frmMain.lblarmor(0) = MaxDef
280               frmMain.lblarmor(1) = MaxDef
290               frmMain.lblarmor(2) = MaxDef
300               UserArmourEqpSlot = Slot
310           Case eOBJType.otescudo
320               frmMain.lblShielder(0) = MaxDef
330               frmMain.lblShielder(1) = MaxDef
340               frmMain.lblShielder(2) = MaxDef
350               UserHelmEqpSlot = Slot
360           Case eOBJType.otcasco
370               frmMain.lblhelm(0) = MaxDef
380               frmMain.lblhelm(1) = MaxDef
390               frmMain.lblhelm(2) = MaxDef
400               UserShieldEqpSlot = Slot
410           End Select
420       Else
430           Select Case Slot
              Case UserWeaponEqpSlot
440               frmMain.lblWeapon(0) = "N/A"
450               frmMain.lblWeapon(1) = "N/A"
460               frmMain.lblWeapon(2) = "N/A"
470               UserWeaponEqpSlot = 0
480           Case UserArmourEqpSlot
490               frmMain.lblarmor(0) = "N/A"
500               frmMain.lblarmor(1) = "N/A"
510               frmMain.lblarmor(2) = "N/A"
520               UserArmourEqpSlot = 0
530           Case UserHelmEqpSlot
540               frmMain.lblShielder(0) = "N/A"
550               frmMain.lblShielder(1) = "N/A"
560               frmMain.lblShielder(2) = "N/A"
570               UserHelmEqpSlot = 0
580           Case UserShieldEqpSlot
590               frmMain.lblhelm(0) = "N/A"
600               frmMain.lblhelm(1) = "N/A"
610               frmMain.lblhelm(2) = "N/A"
620               UserShieldEqpSlot = 0
630           End Select
640       End If


650       Call Inventario.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, _
              MaxHit, MinHit, MaxDef, MinDef, value, Name)
End Sub


' Handles the StopWorking message.
Private Sub HandleStopWorking()
      '***************************************************
      'Author: Budi
      'Last Modification: 12/01/09
      '
      '***************************************************

10        Call incomingData.ReadByte

20        With FontTypes(FontTypeNames.FONTTYPE_INFO)
30            Call ShowConsoleMsg("¡Has terminado de trabajar!", .red, .green, .blue, _
                  .bold, .italic)
40        End With

50        If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
End Sub

' Handles the CancelOfferItem message.

Private Sub HandleCancelOfferItem()
      '***************************************************
      'Author: Torres Patricio (Pato)
      'Last Modification: 05/03/10
      '
      '***************************************************
          Dim Slot   As Byte
          Dim Amount As Long

10        Call incomingData.ReadByte

20        Slot = incomingData.ReadByte

30        With InvOfferComUsu(0)
40            Amount = .Amount(Slot)

              ' No tiene sentido que se quiten 0 unidades
50            If Amount <> 0 Then
                  ' Actualizo el inventario general
60                Call frmComerciarUsu.UpdateInvCom(.ObjIndex(Slot), Amount)

                  ' Borro el item
70                Call .SetItem(Slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
80            End If
90        End With

          ' Si era el único ítem de la oferta, no puede confirmarla
100       If Not frmComerciarUsu.HasAnyItem(InvOfferComUsu(0)) And Not _
              frmComerciarUsu.HasAnyItem(InvOroComUsu(1)) Then Call _
              frmComerciarUsu.HabilitarConfirmar(False)

110       With FontTypes(FontTypeNames.FONTTYPE_INFO)
120           Call frmComerciarUsu.PrintCommerceMsg("¡No puedes comerciar ese objeto!", _
                  FontTypeNames.FONTTYPE_INFO)
130       End With
End Sub

''
' Handles the ChangeBankSlot message.

Private Sub HandleChangeBankSlot()
      '***************************************************
      'Author: Lautaro
      'Last Modification: 05/06/2018
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        Call incomingData.ReadByte

          Dim Slot As Byte
60        Slot = incomingData.ReadByte()

70        With UserBancoInventory(Slot)
80            .ObjIndex = incomingData.ReadInteger()
90            .Amount = incomingData.ReadInteger()
              
100           If .ObjIndex > 0 Then
110               .GrhIndex = incomingData.ReadInteger()
120               .ObjType = incomingData.ReadByte()
130               .MaxHit = incomingData.ReadInteger()
140               .MinHit = incomingData.ReadInteger()
150               .MaxDef = incomingData.ReadInteger()
160               .MinDef = incomingData.ReadInteger
170               .Valor = incomingData.ReadLong()
180               .Name = ObjName(.ObjIndex).Name
190           Else
200               .GrhIndex = 0
210               .ObjType = 0
220               .MaxHit = 0
230               .MinHit = 0
240               .MaxDef = 0
250               .MinDef = 0
260               .Valor = 0
270               .Name = vbNullString
280           End If
              

290           If Comerciando Then
300               Call InvBanco(0).SetItem(Slot, .ObjIndex, .Amount, .Equipped, _
                      .GrhIndex, .ObjType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, _
                      .Name)
310           End If
320       End With
End Sub
''
' Handles the ChangeSpellSlot message.

Private Sub HandleChangeSpellSlot()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 6 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim Slot   As Byte
80        Slot = buffer.ReadByte()

90        UserHechizos(Slot) = buffer.ReadInteger()

100       If Slot <= frmMain.hlst.ListCount Then
110           frmMain.hlst.List(Slot - 1) = buffer.ReadASCIIString()
120       Else
130           Call frmMain.hlst.AddItem(buffer.ReadASCIIString())
140       End If

          'If we got here then packet is complete, copy data back to original queue
150       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
160       error = Err.number
170       On Error GoTo 0

          'Destroy auxiliar buffer
180       Set buffer = Nothing

190       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the Attributes message.

Private Sub HandleAtributes()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 1 + NUMATRIBUTES Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          Dim i      As Long

60        For i = 1 To NUMATRIBUTES
70            UserAtributos(i) = incomingData.ReadByte()
80        Next i

          'Show them in character creation
90        If EstadoLogin = E_MODO.Dados Then
100           With frmCrearPersonaje
110               If .Visible Then
120                   For i = 1 To NUMATRIBUTES
130                       .lblAtributos(i).Caption = UserAtributos(i)
140                   Next i

150                   .UpdateStats
160               End If
170           End With
180       Else
190           LlegaronAtrib = True
200       End If
End Sub

''
' Handles the BlacksmithWeapons message.

Private Sub HandleBlacksmithWeapons()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim Count  As Integer
          Dim i      As Long
          Dim tmp    As String

80        Count = buffer.ReadInteger()

90        Call frmHerrero.lstArmas.Clear

100       For i = 1 To Count
110           tmp = buffer.ReadASCIIString() & " ("           'Get the object's name
120           tmp = tmp & CStr(buffer.ReadInteger()) & ","    'The iron needed
130           tmp = tmp & CStr(buffer.ReadInteger()) & ","    'The silver needed
140           tmp = tmp & CStr(buffer.ReadInteger()) & ")"    'The gold needed

150           Call frmHerrero.lstArmas.AddItem(tmp)
160           ArmasHerrero(i) = buffer.ReadInteger()
170       Next i

180       For i = i To UBound(ArmasHerrero())
190           ArmasHerrero(i) = 0
200       Next i

          'If we got here then packet is complete, copy data back to original queue
210       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
220       error = Err.number
230       On Error GoTo 0

          'Destroy auxiliar buffer
240       Set buffer = Nothing

250       If error <> 0 Then Err.Raise error
End Sub


''
' Handles the BlacksmithArmors message.

Private Sub HandleBlacksmithArmors()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim Count  As Integer
          Dim i      As Long
          Dim tmp    As String

80        Count = buffer.ReadInteger()

90        Call frmHerrero.lstArmaduras.Clear

100       For i = 1 To Count
110           tmp = buffer.ReadASCIIString() & " ("           'Get the object's name
120           tmp = tmp & CStr(buffer.ReadInteger()) & ","    'The iron needed
130           tmp = tmp & CStr(buffer.ReadInteger()) & ","    'The silver needed
140           tmp = tmp & CStr(buffer.ReadInteger()) & ")"    'The gold needed

150           Call frmHerrero.lstArmaduras.AddItem(tmp)
160           ArmadurasHerrero(i) = buffer.ReadInteger()
170       Next i

180       For i = i To UBound(ArmadurasHerrero())
190           ArmadurasHerrero(i) = 0
200       Next i

          'If we got here then packet is complete, copy data back to original queue
210       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
220       error = Err.number
230       On Error GoTo 0

          'Destroy auxiliar buffer
240       Set buffer = Nothing

250       If error <> 0 Then Err.Raise error
End Sub


''
' Handles the CarpenterObjects message.

Private Sub HandleCarpenterObjects()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim Count  As Integer
          Dim i      As Long
          Dim tmp    As String

80        Count = buffer.ReadInteger()

90        Call frmCarp.lstArmas.Clear

100       For i = 1 To Count
110           tmp = buffer.ReadASCIIString() & " ("          'Get the object's name
120           tmp = tmp & CStr(buffer.ReadInteger()) & ")"    'The wood needed

130           Call frmCarp.lstArmas.AddItem(tmp)
140           ObjCarpintero(i) = buffer.ReadInteger()
150       Next i

160       For i = i To UBound(ObjCarpintero())
170           ObjCarpintero(i) = 0
180       Next i

          'If we got here then packet is complete, copy data back to original queue
190       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
200       error = Err.number
210       On Error GoTo 0

          'Destroy auxiliar buffer
220       Set buffer = Nothing

230       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the RestOK message.

Private Sub HandleRestOK()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        UserDescansar = Not UserDescansar
End Sub

''
' Handles the ErrorMessage message.

Private Sub HandleErrorMessage()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
60        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
70        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
80        Call buffer.ReadByte

90        Call MsgBox(buffer.ReadASCIIString())

100       If frmConnect.Visible And (Not frmCrearPersonaje.Visible) Then
        #If UsarWrench = 1 Then
110               frmMain.Socket1.Disconnect
120               frmMain.Socket1.Cleanup
        #Else
130               If frmMain.Winsock1.State <> sckClosed Then frmMain.Winsock1.Close
        #End If
140       End If

          'If we got here then packet is complete, copy data back to original queue
150       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
160       error = Err.number
170       On Error GoTo 0

          'Destroy auxiliar buffer
180       Set buffer = Nothing

190       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the Blind message.

Private Sub HandleBlind()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        UserCiego = True
End Sub

''
' Handles the Dumb message.

Private Sub HandleDumb()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        UserEstupido = True
End Sub

''
' Handles the ShowSignal message.

Private Sub HandleShowSignal()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 5 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim tmp    As String
80        tmp = buffer.ReadASCIIString()

90        buffer.ReadInteger

          'If we got here then packet is complete, copy data back to original queue
100       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
110       error = Err.number
120       On Error GoTo 0

          'Destroy auxiliar buffer
130       Set buffer = Nothing

140       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the ChangeNPCInventorySlot message.

Private Sub HandleChangeNPCInventorySlot()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If


          'Remove packet ID
50        Call incomingData.ReadByte

          Dim Slot   As Byte
60        Slot = incomingData.ReadByte()

70        With NPCInventory(Slot)
80            .ObjIndex = incomingData.ReadInteger()
              
90            If .ObjIndex > 0 Then
100               .Amount = incomingData.ReadInteger()
110               .Valor = incomingData.ReadSingle()
120               .Copas = incomingData.ReadByte()
130               .Eldhir = incomingData.ReadByte()
140               .GrhIndex = incomingData.ReadInteger()
                  
150               .ObjType = incomingData.ReadByte()
160               .MaxHit = incomingData.ReadInteger()
170               .MinHit = incomingData.ReadInteger()
180               .MaxDef = incomingData.ReadInteger()
190               .MinDef = incomingData.ReadInteger()
200               .Name = ObjName(.ObjIndex).Name
210           Else
220               .Amount = 0
230               .Valor = 0
240               .Copas = 0
250               .Eldhir = 0
260               .GrhIndex = 0
                  
270               .ObjType = 0
280               .MaxHit = 0
290               .MinHit = 0
300               .MaxDef = 0
310               .MinDef = 0
320               .Name = vbNullString
330           End If
340       End With
End Sub

''
' Handles the UpdateHungerAndThirst message.

Private Sub HandleUpdateHungerAndThirst()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 5 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

60        UserMaxAGU = incomingData.ReadByte()
70        UserMinAGU = incomingData.ReadByte()
80        UserMaxHAM = incomingData.ReadByte()
90        UserMinHAM = incomingData.ReadByte()

100       'frmMain.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 95)
110       'frmMain.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 95)
          Dim i      As Long
120       For i = 0 To 4
130           frmMain.Lblham(i) = UserMinHAM & "%" '& UserMaxHAM
140           frmMain.lblsed(i) = UserMinAGU & "%" '& UserMaxAGU
150       Next i

End Sub

''
' Handles the Fame message.

Private Sub HandleFame()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 29 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

60        With UserReputacion
70            .AsesinoRep = incomingData.ReadLong()
80            .BandidoRep = incomingData.ReadLong()
90            .BurguesRep = incomingData.ReadLong()
100           .LadronesRep = incomingData.ReadLong()
110           .NobleRep = incomingData.ReadLong()
120           .PlebeRep = incomingData.ReadLong()
130           .Promedio = incomingData.ReadLong()
140       End With

150       LlegoFama = True
End Sub

''
' Handles the MiniStats message.

Private Sub HandleMiniStats()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 20 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

60        With UserEstadisticas
70            .CiudadanosMatados = incomingData.ReadLong()
80            .CriminalesMatados = incomingData.ReadLong()
90            .UsuariosMatados = incomingData.ReadLong()
100           .NpcsMatados = incomingData.ReadInteger()
110           .Clase = ListaClases(incomingData.ReadByte())
120           .PenaCarcel = incomingData.ReadLong()
130       End With
End Sub

''
' Handles the LevelUp message.

Private Sub HandleLevelUp()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte
          
          Dim Skills As Integer
          Skills = incomingData.ReadInteger
          
          If Skills = -1 Then
               SkillPoints = 0
          Else
               SkillPoints = SkillPoints + Skills
          End If
60

70        Call frmMain.LightSkillStar(True)
End Sub

''
' Handles the AddForumMessage message.

Private Sub HandleAddForumMessage()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 8 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim ForumType As eForumMsgType
          Dim Title  As String
          Dim Message As String
          Dim Author As String
          Dim bAnuncio As Boolean
          Dim bSticky As Boolean

80        ForumType = buffer.ReadByte

90        Title = buffer.ReadASCIIString()
100       Author = buffer.ReadASCIIString()
110       Message = buffer.ReadASCIIString()

          'If Not frmForo.ForoLimpio Then
120       clsForos.ClearForums
          'frmForo.ForoLimpio = True
          'End If

130       Call clsForos.AddPost(ForumAlignment(ForumType), Title, Author, Message, _
              EsAnuncio(ForumType))

          'If we got here then packet is complete, copy data back to original queue
140       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
150       error = Err.number
160       On Error GoTo 0

          'Destroy auxiliar buffer
170       Set buffer = Nothing

180       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the ShowForumForm message.

Private Sub HandleShowForumForm()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

          'frmForo.Privilegios = incomingData.ReadByte
          'frmForo.CanPostSticky = incomingData.ReadByte

20        If Not MirandoForo Then
              'frmForo.Show , frmMain
30        End If
End Sub

''
' Handles the SetInvisible message.

Private Sub HandleSetInvisible()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 4 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          Dim CharIndex As Integer
          Dim Invisible As Boolean
          
          
60        CharIndex = incomingData.ReadInteger()
          Invisible = incomingData.ReadBoolean()
          
          If CharIndex > 0 Then
            charlist(CharIndex).Invisible = Invisible
80          UserOcu = charlist(CharIndex).Invisible
90          charlist(CharIndex).cCont = 400
100         charlist(CharIndex).Drawers = 0
          End If
          
110       If CharIndex = UserCharIndex Then
120           charlist(CharIndex).cCont = 0
130           charlist(CharIndex).Drawers = 156
140       End If
End Sub

''
' Handles the DiceRoll message.

Private Sub HandleDiceRoll()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 6 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

60        UserAtributos(eAtributos.Fuerza) = incomingData.ReadByte()
70        UserAtributos(eAtributos.Agilidad) = incomingData.ReadByte()
80        UserAtributos(eAtributos.Inteligencia) = incomingData.ReadByte()
90        UserAtributos(eAtributos.Carisma) = incomingData.ReadByte()
100       UserAtributos(eAtributos.Constitucion) = incomingData.ReadByte()

110       With frmCrearPersonaje
120           .lblAtributos(eAtributos.Fuerza) = UserAtributos(eAtributos.Fuerza)
130           .lblAtributos(eAtributos.Agilidad) = UserAtributos(eAtributos.Agilidad)
140           .lblAtributos(eAtributos.Inteligencia) = _
                  UserAtributos(eAtributos.Inteligencia)
150           .lblAtributos(eAtributos.Carisma) = UserAtributos(eAtributos.Carisma)
160           .lblAtributos(eAtributos.Constitucion) = _
                  UserAtributos(eAtributos.Constitucion)

170           .UpdateStats
180       End With
End Sub

''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        UserMeditar = Not UserMeditar
End Sub

''
' Handles the BlindNoMore message.

Private Sub HandleBlindNoMore()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        UserCiego = False
End Sub

''
' Handles the DumbNoMore message.

Private Sub HandleDumbNoMore()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        UserEstupido = False
End Sub

''
' Handles the SendSkills message.

Private Sub HandleSendSkills()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 11/19/09
      '11/19/09: Pato - Now the server send the percentage of progress of the skills.
      '***************************************************
10        If incomingData.Length < 2 + NUMSKILLS * 2 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

60        UserClase = incomingData.ReadByte
          Dim i      As Long

70        For i = 1 To NUMSKILLS
80            UserSkills(i) = incomingData.ReadByte()
90            PorcentajeSkills(i) = incomingData.ReadByte()
100       Next i
110       LlegaronSkills = True
End Sub

''
' Handles the TrainerCreatureList message.

Private Sub HandleTrainerCreatureList()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim creatures() As String
          Dim i      As Long

80        creatures = Split(buffer.ReadASCIIString(), SEPARATOR)

90        For i = 0 To UBound(creatures())
100           Call frmEntrenador.lstCriaturas.AddItem(creatures(i))
110       Next i
120       frmEntrenador.Show , frmMain

          'If we got here then packet is complete, copy data back to original queue
130       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
140       error = Err.number
150       On Error GoTo 0

          'Destroy auxiliar buffer
160       Set buffer = Nothing

170       If error <> 0 Then Err.Raise error
End Sub
''
' Handles the GuildNews message.

Private Sub HandleGuildNews()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 7 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim guildList() As String
          Dim i      As Long

          'Get news' string
80        frmGuildNews.news = buffer.ReadASCIIString()

          'Get Enemy guilds list
90        guildList = Split(buffer.ReadASCIIString(), SEPARATOR)

100       For i = 0 To UBound(guildList)
110           Call frmGuildNews.txtClanesGuerra.AddItem(guildList(i))
120       Next i

          'Get Allied guilds list
130       guildList = Split(buffer.ReadASCIIString(), SEPARATOR)

140       For i = 0 To UBound(guildList)
150           Call frmGuildNews.txtClanesAliados.AddItem(guildList(i))
160       Next i

170       frmGuildNews.Show vbModeless, frmMain

          'If we got here then packet is complete, copy data back to original queue
180       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
190       error = Err.number
200       On Error GoTo 0

          'Destroy auxiliar buffer
210       Set buffer = Nothing

220       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the OfferDetails message.

Private Sub HandleOfferDetails()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

80        Call frmUserRequest.recievePeticion(buffer.ReadASCIIString())

          'If we got here then packet is complete, copy data back to original queue
90        Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
100       error = Err.number
110       On Error GoTo 0

          'Destroy auxiliar buffer
120       Set buffer = Nothing

130       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the AlianceProposalsList message.

Private Sub HandleAlianceProposalsList()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim vsGuildList() As String
          Dim i      As Long

80        vsGuildList = Split(buffer.ReadASCIIString(), SEPARATOR)

90        Call frmPeaceProp.lista.Clear
100       For i = 0 To UBound(vsGuildList())
110           Call frmPeaceProp.lista.AddItem(vsGuildList(i))
120       Next i

130       frmPeaceProp.ProposalType = TIPO_PROPUESTA.ALIANZA
140       Call frmPeaceProp.Show(vbModeless, frmMain)

          'If we got here then packet is complete, copy data back to original queue
150       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
160       error = Err.number
170       On Error GoTo 0

          'Destroy auxiliar buffer
180       Set buffer = Nothing

190       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the PeaceProposalsList message.

Private Sub HandlePeaceProposalsList()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim guildList() As String
          Dim i      As Long

80        guildList = Split(buffer.ReadASCIIString(), SEPARATOR)

90        Call frmPeaceProp.lista.Clear
100       For i = 0 To UBound(guildList())
110           Call frmPeaceProp.lista.AddItem(guildList(i))
120       Next i

130       frmPeaceProp.ProposalType = TIPO_PROPUESTA.PAZ
140       Call frmPeaceProp.Show(vbModeless, frmMain)

          'If we got here then packet is complete, copy data back to original queue
150       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
160       error = Err.number
170       On Error GoTo 0

          'Destroy auxiliar buffer
180       Set buffer = Nothing

190       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the CharacterInfo message.

Private Sub HandleCharacterInfo()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 35 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

80        With frmCharInfo
90            If .frmType = CharInfoFrmType.frmMembers Then
100               .imgRechazar.Visible = False
110               .imgAceptar.Visible = False
120               .imgEchar.Visible = True
130               .imgPeticion.Visible = False
140           Else
150               .imgRechazar.Visible = True
160               .imgAceptar.Visible = True
170               .imgEchar.Visible = False
180               .imgPeticion.Visible = True
190           End If

200           .Nombre.Caption = buffer.ReadASCIIString()
210           .Raza.Caption = ListaRazas(buffer.ReadByte())
220           .Clase.Caption = ListaClases(buffer.ReadByte())

230           If buffer.ReadByte() = 1 Then
240               .Genero.Caption = "Hombre"
250           Else
260               .Genero.Caption = "Mujer"
270           End If

280           .Nivel.Caption = buffer.ReadByte()
290           .Oro.Caption = buffer.ReadLong()
300           .Banco.Caption = buffer.ReadLong()

              Dim reputation As Long
310           reputation = buffer.ReadLong()

320           .reputacion.Caption = reputation

330           .txtPeticiones.Text = buffer.ReadASCIIString()
340           .guildactual.Caption = buffer.ReadASCIIString()
350           .txtMiembro.Text = buffer.ReadASCIIString()

              Dim armada As Boolean
              Dim caos As Boolean

360           armada = buffer.ReadBoolean()
370           caos = buffer.ReadBoolean()

380           If armada Then
390               .ejercito.Caption = "Armada Real"
400           ElseIf caos Then
410               .ejercito.Caption = "Legión Oscura"
420           End If

430           .Ciudadanos.Caption = CStr(buffer.ReadLong())
440           .criminales.Caption = CStr(buffer.ReadLong())

450           If reputation > 0 Then
460               .status.Caption = " Ciudadano"
470               .status.ForeColor = vbBlue
480           Else
490               .status.Caption = " Criminal"
500               .status.ForeColor = vbRed
510           End If

520           Call .Show(vbModeless, frmMain)
530       End With

          'If we got here then packet is complete, copy data back to original queue
540       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
550       error = Err.number
560       On Error GoTo 0

          'Destroy auxiliar buffer
570       Set buffer = Nothing

580       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the GuildLeaderInfo message.

Private Sub HandleGuildLeaderInfo()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 9 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim i      As Long
          Dim List() As String

80        With frmGuildLeader
              'Get list of existing guilds
90            GuildNames = Split(buffer.ReadASCIIString(), SEPARATOR)

              'Empty the list
100           Call .guildslist.Clear

110           For i = 0 To UBound(GuildNames())
120               Call .guildslist.AddItem(GuildNames(i))
130           Next i

              'Get list of guild's members
140           GuildMembers = Split(buffer.ReadASCIIString(), SEPARATOR)
150           .Miembros.Caption = CStr(UBound(GuildMembers()) + 1)

              'Empty the list
160           Call .members.Clear

170           For i = 0 To UBound(GuildMembers())
180               Call .members.AddItem(GuildMembers(i))
190           Next i

200           .txtguildnews = buffer.ReadASCIIString()

              'Get list of join requests
210           List = Split(buffer.ReadASCIIString(), SEPARATOR)

              'Empty the list
220           Call .solicitudes.Clear

230           For i = 0 To UBound(List())
240               Call .solicitudes.AddItem(List(i))
250           Next i

260           .Show , frmMain
270       End With

          'If we got here then packet is complete, copy data back to original queue
280       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
290       error = Err.number
300       On Error GoTo 0

          'Destroy auxiliar buffer
310       Set buffer = Nothing

320       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the GuildDetails message.

Private Sub HandleGuildDetails()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 26 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
60        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
70        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
80        Call buffer.ReadByte

90        With frmGuildBrief
100           .ImgDeclararGuerra.Visible = .EsLeader
110           .ImgOfrecerAlianza.Visible = .EsLeader
120           .ImgOfrecerPaz.Visible = .EsLeader

130           .Nombre.Caption = buffer.ReadASCIIString()
140           .fundador.Caption = buffer.ReadASCIIString()
150           .creacion.Caption = buffer.ReadASCIIString()
160           .lider.Caption = buffer.ReadASCIIString()
170           .web.Caption = buffer.ReadASCIIString()
180           .Miembros.Caption = buffer.ReadInteger()

190           If buffer.ReadBoolean() Then
200               .eleccion.Caption = "ABIERTA"
210           Else
220               .eleccion.Caption = "CERRADA"
230           End If

240           .lblAlineacion.Caption = buffer.ReadASCIIString()
250           .Enemigos.Caption = buffer.ReadInteger()
260           .Aliados.Caption = buffer.ReadInteger()
270           .antifaccion.Caption = buffer.ReadASCIIString()

              Dim codexStr() As String
              Dim i  As Long

280           codexStr = Split(buffer.ReadASCIIString(), SEPARATOR)

290           For i = 0 To 7
300               .Codex(i).Caption = codexStr(i)
310           Next i

320           .Desc.Text = buffer.ReadASCIIString()
330       End With

          'If we got here then packet is complete, copy data back to original queue
340       Call incomingData.CopyBuffer(buffer)

350       frmGuildBrief.Show vbModeless, frmMain

ErrHandler:
          Dim error  As Long
360       error = Err.number
370       On Error GoTo 0

          'Destroy auxiliar buffer
380       Set buffer = Nothing

390       If error <> 0 Then Err.Raise error
End Sub
''
' Handles the ShowGuildAlign message.

Private Sub HandleShowGuildAlign()
      '***************************************************
      'Author: ZaMa
      'Last Modification: 14/12/2009
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        frmGuildFoundation.Show , frmMain
End Sub


''
' Handles the ShowGuildFundationForm message.

Private Sub HandleShowGuildFundationForm()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        CreandoClan = True
30        frmGuildFoundation.Show , frmMain
End Sub
Private Sub HandleMovimientSW()

10        With incomingData

20            Call .ReadByte

              Dim Char As Integer
              Dim MovimientClass As Byte

30            Char = .ReadInteger()
40            MovimientClass = .ReadByte()


50            With charlist(Char)

60                If TSetup.bGameCombat = True Then
70                    If MovimientClass = 1 Then   '1 = mover arma.

80                        .Arma.WeaponWalk(.Heading).Started = 1

90                    ElseIf MovimientClass = 2 Then

100                       .Escudo.ShieldWalk(.Heading).Started = 1

110                   End If

120                   .Movimient = True
130               Else
140                   .Movimient = False
150               End If
160           End With
170       End With

End Sub

''
' Handles the ParalizeOK message.

Private Sub HandleParalizeOK()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        UserParalizado = Not UserParalizado
End Sub

''
' Handles the ShowUserRequest message.

Private Sub HandleShowUserRequest()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

80        Call frmUserRequest.recievePeticion(buffer.ReadASCIIString())
90        Call frmUserRequest.Show(vbModeless, frmMain)

          'If we got here then packet is complete, copy data back to original queue
100       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
110       error = Err.number
120       On Error GoTo 0

          'Destroy auxiliar buffer
130       Set buffer = Nothing

140       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the TradeOK message.

Private Sub HandleTradeOK()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        If frmComerciar.Visible Then
              Dim i  As Long

              'Update user inventory
30            For i = 1 To MAX_INVENTORY_SLOTS
                  ' Agrego o quito un item en su totalidad
40                If Inventario.ObjIndex(i) <> InvComUsu.ObjIndex(i) Then
50                    With Inventario
60                        Call InvComUsu.SetItem(i, .ObjIndex(i), .Amount(i), _
                              .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), _
                              .MinHit(i), .MaxDef(i), .MinDef(i), .Valor(i), .ItemName(i))
70                    End With
                      ' Vendio o compro cierta cantidad de un item que ya tenia
80                ElseIf Inventario.Amount(i) <> InvComUsu.Amount(i) Then
90                    Call InvComUsu.ChangeSlotItemAmount(i, Inventario.Amount(i))
100               End If
110           Next i

              ' Fill Npc inventory
120           For i = 1 To 20
                  ' Compraron la totalidad de un item, o vendieron un item que el npc no tenia
130               If NPCInventory(i).ObjIndex <> InvComNpc.ObjIndex(i) Then
140                   With NPCInventory(i)
150                       Call InvComNpc.SetItem(i, .ObjIndex, .Amount, 0, .GrhIndex, _
                              .ObjType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .Name)
160                   End With
                      ' Compraron o vendieron cierta cantidad (no su totalidad)
170               ElseIf NPCInventory(i).Amount <> InvComNpc.Amount(i) Then
180                   Call InvComNpc.ChangeSlotItemAmount(i, NPCInventory(i).Amount)
190               End If
200           Next i

210       End If
End Sub

''
' Handles the BankOK message.

Private Sub HandleBankOK()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

          Dim i      As Long

20        If frmBancoObj.Visible Then

30            For i = 1 To Inventario.MaxObjs
40                With Inventario
50                    Call InvBanco(1).SetItem(i, .ObjIndex(i), .Amount(i), _
                          .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), _
                          .MaxDef(i), .MinDef(i), .Valor(i), .ItemName(i))
60                End With
70            Next i

              'Alter order according to if we bought or sold so the labels and grh remain the same
80            If frmBancoObj.LasActionBuy Then
                  'frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
                  'frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
90            Else
                  'frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
                  'frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
100           End If

110           frmBancoObj.NoPuedeMover = False
120       End If

End Sub

''
' Handles the ChangeUserTradeSlot message.

Private Sub HandleChangeUserTradeSlot()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 19 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          Dim OfferSlot As Byte
          Dim ObjIndex As Integer
          Dim strName As String
          
          'Remove packet ID
50        Call incomingData.ReadByte

60        OfferSlot = incomingData.ReadByte
70        ObjIndex = incomingData.ReadInteger
          
80        If ObjIndex > 0 Then
90            strName = ObjName(ObjIndex).Name
100       Else
110           strName = vbNullString
120       End If
          
130       With incomingData
140           If OfferSlot = GOLD_OFFER_SLOT Then
150               Call InvOroComUsu(2).SetItem(1, ObjIndex, .ReadLong(), 0, _
                      .ReadInteger(), .ReadByte(), .ReadInteger(), .ReadInteger(), _
                      .ReadInteger(), .ReadInteger(), .ReadLong(), strName)
160           Else
170               Call InvOfferComUsu(1).SetItem(OfferSlot, ObjIndex, .ReadLong(), 0, _
                      .ReadInteger(), .ReadByte(), .ReadInteger(), .ReadInteger(), _
                      .ReadInteger(), .ReadInteger(), .ReadLong(), strName)
180           End If
190       End With

200       Call frmComerciarUsu.PrintCommerceMsg(TradingUserName & _
              " ha modificado su oferta.", FontTypeNames.FONTTYPE_VENENO)
End Sub

''
' Handles the SendNight message.

Private Sub HandleSendNight()
      '***************************************************
      'Author: Fredy Horacio Treboux (liquid)
      'Last Modification: 01/08/07
      '
      '***************************************************
10        If incomingData.Length < 2 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

          'Remove packet ID
50        Call incomingData.ReadByte

          Dim tBool  As Boolean    'CHECK, este handle no hace nada con lo que recibe.. porque, ehmm.. no hay noche?.. o si?
60        tBool = incomingData.ReadBoolean()
End Sub

''
' Handles the SpawnList message.

Private Sub HandleSpawnList()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim creatureList() As String
          Dim i      As Long

80        creatureList = Split(buffer.ReadASCIIString(), SEPARATOR)

90        For i = 0 To UBound(creatureList())
100           Call frmSpawnList.lstCriaturas.AddItem(creatureList(i))
110       Next i
120       frmSpawnList.Show , frmMain

          'If we got here then packet is complete, copy data back to original queue
130       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
140       error = Err.number
150       On Error GoTo 0

          'Destroy auxiliar buffer
160       Set buffer = Nothing

170       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowSOSForm()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim sosList() As String
          Dim i      As Long

80        sosList = Split(buffer.ReadASCIIString(), SEPARATOR)

90        For i = 0 To UBound(sosList())
100           Call frmMSG.List1.AddItem(sosList(i))
110       Next i

120       frmMSG.Show , frmMain

          'If we got here then packet is complete, copy data back to original queue
130       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
140       error = Err.number
150       On Error GoTo 0

          'Destroy auxiliar buffer
160       Set buffer = Nothing

170       If error <> 0 Then Err.Raise error
End Sub


''

''
' Handles the ShowGMPanelForm message.

Private Sub HandleShowGMPanelForm()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
      'Remove packet ID
10        Call incomingData.ReadByte

20        frmPanelGm.Show vbModeless, frmMain
End Sub

''
' Handles the UserNameList message.

Private Sub HandleUserNameList()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
   On Error GoTo HandleUserNameList_Error

10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim userList() As String
          Dim i      As Long

80        userList = Split(buffer.ReadASCIIString(), SEPARATOR)

90        If frmPanelGm.Visible Then
100           frmPanelGm.cboListaUsus.Clear
110           For i = 0 To UBound(userList())
120               Call frmPanelGm.cboListaUsus.AddItem(userList(i))
130           Next i
140           If frmPanelGm.cboListaUsus.ListCount > 0 Then _
                  frmPanelGm.cboListaUsus.ListIndex = 0
150       End If

          'If we got here then packet is complete, copy data back to original queue
160       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
170       error = Err.number
180       On Error GoTo 0

          'Destroy auxiliar buffer
190       Set buffer = Nothing

200       If error <> 0 Then Err.Raise error

   On Error GoTo 0
   Exit Sub

HandleUserNameList_Error:

    'LogError "Error " & Err.number & " (" & Err.Description & ") in procedure HandleUserNameList of Módulo Protocol in line " & Erl
End Sub

''
' Handles the Pong message.

Private Sub HandlePong()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        Call incomingData.ReadByte
          
          Dim pingAnswer As Long
          
20        pingAnswer = GetTickCount - pingTime

          Dim DescPing As String

30        Select Case pingAnswer

              Case 0 To 100
40                DescPing = "Jugable"

50            Case 101 To 200
60                DescPing = "Bajo"

70            Case 201 To 500
80                DescPing = "Medio"

90            Case 501 To 1000
100               DescPing = "Alto"

110           Case Else
120               DescPing = "Injugable"

130       End Select
          

          
140       If pingAnswer < 1000000 Then Call AddtoRichTextBox(frmMain.RecTxt, _
              "El ping es de " & pingAnswer & " ms (" & pingAnswer / 1000 & " Seg) LAG: " _
              & DescPing, 0, 255, 0, True, False, True)
150       pingTime = 0
End Sub

''
' Handles the Pong message.

Private Sub HandleGuildMemberInfo()
      '***************************************************
      'Author: ZaMa
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

80        With frmGuildMember
              'Clear guild's list
90            .lstClanes.Clear

100           GuildNames = Split(buffer.ReadASCIIString(), SEPARATOR)

              Dim i  As Long
110           For i = 0 To UBound(GuildNames())
120               Call .lstClanes.AddItem(GuildNames(i))
130           Next i

              'Get list of guild's members
140           GuildMembers = Split(buffer.ReadASCIIString(), SEPARATOR)
150           .lblCantMiembros.Caption = CStr(UBound(GuildMembers()) + 1)

              'Empty the list
160           Call .lstMiembros.Clear

170           For i = 0 To UBound(GuildMembers())
180               Call .lstMiembros.AddItem(GuildMembers(i))
190           Next i

              'If we got here then packet is complete, copy data back to original queue
200           Call incomingData.CopyBuffer(buffer)

210           .Show vbModeless, frmMain
220       End With

ErrHandler:
          Dim error  As Long
230       error = Err.number
240       On Error GoTo 0

          'Destroy auxiliar buffer
250       Set buffer = Nothing

260       If error <> 0 Then Err.Raise error
End Sub

''
' Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      '
      '***************************************************
10        If incomingData.Length < 6 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          'Remove packet ID
70        Call buffer.ReadByte

          Dim CharIndex As Integer
          Dim NickColor As Byte
          Dim UserTag As String
          Dim UserInfectado As Byte
          Dim UserAngel As Byte
          Dim UserDemonio As Byte


80        CharIndex = buffer.ReadInteger()
90        NickColor = buffer.ReadByte()
100       UserTag = buffer.ReadASCIIString()
          
          'Update char status adn tag!
110       With charlist(CharIndex)
120           .Infected = UserInfectado
130           .Angel = UserAngel
140           .Demonio = UserDemonio
              
150           If (NickColor And eNickColor.ieCriminal) <> 0 Then
160               .Criminal = 1
170           Else
180               .Criminal = 0
190           End If
              
200           If (NickColor And eNickColor.ieTeamUno) <> 0 Then
210               .Team = 1
220           ElseIf (NickColor And eNickColor.ieTeamDos) <> 0 Then
230               .Team = 2
240           Else
250               .Team = 0
260           End If

270           .Atacable = (NickColor And eNickColor.ieAtacable) <> 0

280           .Nombre = UserTag
290       End With

          'If we got here then packet is complete, copy data back to original queue
300       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
310       error = Err.number
320       On Error GoTo 0

          'Destroy auxiliar buffer
330       Set buffer = Nothing

340       If error <> 0 Then Err.Raise error
End Sub


''
' Writes the "ThrowDices" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteThrowDices()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ThrowDices" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call outgoingData.WriteByte(ClientPacketID.ThrowDices)
              'Call SumarFlush
              'Call outgoingData.WriteInteger(SeguridadCRC(CRC) + 42)
30        End With
End Sub



''
' Writes the "Talk" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalk(ByVal chat As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Talk" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.Talk)

30            Call .WriteASCIIString(chat)
40        End With
End Sub

''
' Writes the "Yell" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteYell(ByVal chat As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Yell" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.Yell)

30            Call .WriteASCIIString(chat)
40        End With
End Sub

''
' Writes the "Whisper" message to the outgoing data buffer.
'
' @param    charIndex The index of the char to whom to whisper.
' @param    chat The chat text to be sent to the user.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhisper(ByVal CharIndex As Integer, ByVal chat As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Whisper" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.Whisper)

30            Call .WriteInteger(CharIndex)

40            Call .WriteASCIIString(chat)
50        End With
End Sub

''
' Writes the "Walk" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWalk(ByVal Heading As E_Heading)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Walk" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.Walk)

30            Call .WriteByte(Heading)
40        End With
End Sub

''
' Writes the "RequestPositionUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestPositionUpdate()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RequestPositionUpdate" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.RequestPositionUpdate)
End Sub

''
' Writes the "Attack" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttack()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Attack" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.Attack)
30        End With
End Sub

''
' Writes the "PickUp" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePickUp()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "PickUp" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.PickUp)
End Sub
Public Sub WriteCombatModeToggle()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CombatModeToggle" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.CombatModeToggle)
End Sub

''
' Writes the "SafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeToggle()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "SafeToggle" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.SafeToggle)
End Sub

''
' Writes the "ResuscitationSafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResuscitationToggle()
      '**************************************************************
      'Author: Rapsodius
      'Creation Date: 10/10/07
      'Writes the Resuscitation safe toggle packet to the outgoing data buffer.
      '**************************************************************
10        Call outgoingData.WriteByte(ClientPacketID.ResuscitationSafeToggle)
End Sub

''
' Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestGuildLeaderInfo()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.RequestGuildLeaderInfo)
End Sub

''
' Writes the "ItemUpgrade" message to the outgoing data buffer.
'
' @param    ItemIndex The index to the item to upgrade.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemUpgrade(ByVal ItemIndex As Integer)
      '***************************************************
      'Author: Torres Patricio (Pato)
      'Last Modification: 12/09/09
      'Writes the "ItemUpgrade" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.ItemUpgrade)
20        Call outgoingData.WriteInteger(ItemIndex)
End Sub

''
' Writes the "RequestAtributes" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAtributes()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RequestAtributes" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.RequestAtributes)
End Sub

''
' Writes the "RequestFame" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestFame()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RequestFame" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.RequestFame)
End Sub

''
' Writes the "RequestSkills" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestSkills()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RequestSkills" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.RequestSkills)
End Sub

''
' Writes the "RequestMiniStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMiniStats()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RequestMiniStats" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.RequestMiniStats)
End Sub

''
' Writes the "CommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CommerceEnd" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.CommerceEnd)
End Sub

''
' Writes the "UserCommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UserCommerceEnd" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.UserCommerceEnd)
End Sub

''
' Writes the "UserCommerceConfirm" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceConfirm()
      '***************************************************
      'Author: ZaMa
      'Last Modification: 14/12/2009
      'Writes the "UserCommerceConfirm" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.UserCommerceConfirm)
End Sub

''
' Writes the "BankEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "BankEnd" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.BankEnd)
End Sub

''
' Writes the "UserCommerceOk" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOk()
      '***************************************************
      'Author: Fredy Horacio Treboux (liquid)
      'Last Modification: 01/10/07
      'Writes the "UserCommerceOk" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.UserCommerceOk)
End Sub

''
' Writes the "UserCommerceReject" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceReject()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UserCommerceReject" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.UserCommerceReject)
End Sub

''
' Writes the "Drop" message to the outgoing data buffer.
'
' @param    slot Inventory slot where the item to drop is.
' @param    amount Number of items to drop.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDrop(ByVal Slot As Byte, ByVal Amount As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Drop" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.Drop)

30            Call .WriteByte(Slot)
40            Call .WriteInteger(Amount)

50        End With
End Sub

''
' Writes the "CastSpell" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell to cast is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCastSpell(ByVal Slot As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CastSpell" message to the outgoing data buffer
      '***************************************************
        If frmMain.CmdLanzar.Visible = False Then Exit Sub
10        With outgoingData
20            Call .WriteByte(ClientPacketID.CastSpell)

30            Call .WriteByte(Slot)

                Call .WriteByte(ClaveActual)
                
40        End With
End Sub

''
' Writes the "LeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeftClick(ByVal X As Byte, ByVal Y As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "LeftClick" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.LeftClick)

30            Call .WriteByte(X)
40            Call .WriteByte(Y)
50        End With
End Sub

''
' Writes the "DoubleClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoubleClick(ByVal X As Byte, ByVal Y As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "DoubleClick" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.DoubleClick)

30            Call .WriteByte(X)
40            Call .WriteByte(Y)
50        End With
End Sub

''
' Writes the "Work" message to the outgoing data buffer.
'
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWork(ByVal Skill As eSkill)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Work" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.Work)

30            Call .WriteByte(Skill)
40        End With
End Sub

''
' Writes the "UseSpellMacro" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseSpellMacro()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UseSpellMacro" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.UseSpellMacro)
End Sub

''
' Writes the "CraftBlacksmith" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftBlacksmith(ByVal Item As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CraftBlacksmith" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.CraftBlacksmith)

30            Call .WriteInteger(Item)
40        End With
End Sub

''
' Writes the "CraftCarpenter" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftCarpenter(ByVal Item As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CraftCarpenter" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.CraftCarpenter)

30            Call .WriteInteger(Item)
40        End With
End Sub

''
' Writes the "ShowGuildNews" message to the outgoing data buffer.
'

Public Sub WriteShowGuildNews()
      '***************************************************
      'Author: ZaMa
      'Last Modification: 21/02/2010
      'Writes the "ShowGuildNews" message to the outgoing data buffer
      '***************************************************

10        outgoingData.WriteByte (ClientPacketID.ShowGuildNews)
End Sub


''
' Writes the "WorkLeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkLeftClick(ByVal X As Byte, ByVal Y As Byte, ByVal Skill As _
    eSkill)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "WorkLeftClick" message to the outgoing data buffer
      '***************************************************
      
      
      Dim X1 As Long, Y1 As Long
10        With outgoingData
20            Call .WriteByte(ClientPacketID.WorkLeftClick)
              
              Call .WriteByte(X)
40            Call .WriteByte(Y)

50            Call .WriteByte(Skill)
60        End With
End Sub

''
' Writes the "CreateNewGuild" message to the outgoing data buffer.
'
' @param    desc    The guild's description
' @param    name    The guild's name
' @param    site    The guild's website
' @param    codex   Array of all rules of the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNewGuild(ByVal Desc As String, ByVal Name As String, _
    ByVal Site As String, ByRef Codex() As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CreateNewGuild" message to the outgoing data buffer
      '***************************************************
          Dim temp   As String
          Dim i      As Long

10        With outgoingData
20            Call .WriteByte(ClientPacketID.CreateNewGuild)

30            Call .WriteASCIIString(Desc)
40            Call .WriteASCIIString(Name)
50            Call .WriteASCIIString(Site)

60            For i = LBound(Codex()) To UBound(Codex())
70                temp = temp & Codex(i) & SEPARATOR
80            Next i

90            If Len(temp) Then temp = Left$(temp, Len(temp) - 1)

100           Call .WriteASCIIString(temp)
110       End With
End Sub

''
' Writes the "SpellInfo" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpellInfo(ByVal Slot As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "SpellInfo" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.SpellInfo)

30            Call .WriteByte(Slot)
40        End With
End Sub

''
' Writes the "EquipItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to equip is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquipItem(ByVal Slot As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "EquipItem" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.EquipItem)

30            Call .WriteByte(Slot)
40        End With
End Sub

''
' Writes the "ChangeHeading" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeHeading(ByVal Heading As E_Heading)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ChangeHeading" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.ChangeHeading)

30            Call .WriteByte(Heading)
40        End With
End Sub

''
' Writes the "ModifySkills" message to the outgoing data buffer.
'
' @param    skillEdt a-based array containing for each skill the number of points to add to it.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteModifySkills(ByRef skillEdt() As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ModifySkills" message to the outgoing data buffer
      '***************************************************
          Dim i      As Long

10        With outgoingData
20            Call .WriteByte(ClientPacketID.ModifySkills)

30            For i = 1 To NUMSKILLS
40                Call .WriteByte(skillEdt(i))
50            Next i
60        End With
End Sub

''
' Writes the "Train" message to the outgoing data buffer.
'
' @param    creature Position within the list provided by the server of the creature to train against.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrain(ByVal creature As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Train" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.Train)

30            Call .WriteByte(creature)
40        End With
End Sub

''
' Writes the "CommerceBuy" message to the outgoing data buffer.
'
' @param    slot Position within the NPC's inventory in which the desired item is.
' @param    amount Number of items to buy.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceBuy(ByVal Slot As Byte, ByVal Amount As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CommerceBuy" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.CommerceBuy)

30            Call .WriteByte(Slot)
40            Call .WriteInteger(Amount)
50        End With
End Sub

''
' Writes the "BankExtractItem" message to the outgoing data buffer.
'
' @param    slot Position within the bank in which the desired item is.
' @param    amount Number of items to extract.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractItem(ByVal Slot As Byte, ByVal Amount As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "BankExtractItem" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.BankExtractItem)

30            Call .WriteByte(Slot)
40            Call .WriteInteger(Amount)
50        End With
End Sub

''
' Writes the "CommerceSell" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to sell.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceSell(ByVal Slot As Byte, ByVal Amount As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CommerceSell" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.CommerceSell)

30            Call .WriteByte(Slot)
40            Call .WriteInteger(Amount)
50        End With
End Sub

''
' Writes the "BankDeposit" message to the outgoing data buffer.
'
' @param    slot Position within the user inventory in which the desired item is.
' @param    amount Number of items to deposit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDeposit(ByVal Slot As Byte, ByVal Amount As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "BankDeposit" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.BankDeposit)

30            Call .WriteByte(Slot)
40            Call .WriteInteger(Amount)
50        End With
End Sub

''
' Writes the "ForumPost" message to the outgoing data buffer.
'
' @param    title The message's title.
' @param    message The body of the message.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForumPost(ByVal Title As String, ByVal Message As String, ByVal _
    ForumMsgType As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ForumPost" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.ForumPost)

30            Call .WriteByte(ForumMsgType)
40            Call .WriteASCIIString(Title)
50            Call .WriteASCIIString(Message)
60        End With
End Sub

''
' Writes the "MoveSpell" message to the outgoing data buffer.
'
' @param    upwards True if the spell will be moved up in the list, False if it will be moved downwards.
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveSpell(ByVal upwards As Boolean, ByVal Slot As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "MoveSpell" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.MoveSpell)

30            Call .WriteBoolean(upwards)
40            Call .WriteByte(Slot)
50        End With
End Sub

''
' Writes the "MoveBank" message to the outgoing data buffer.
'
' @param    upwards True if the item will be moved up in the list, False if it will be moved downwards.
' @param    slot Bank List slot where the item which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveBank(ByVal upwards As Boolean, ByVal Slot As Byte)
      '***************************************************
      'Author: Torres Patricio (Pato)
      'Last Modification: 06/14/09
      'Writes the "MoveBank" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.MoveBank)

30            Call .WriteBoolean(upwards)
40            Call .WriteByte(Slot)
50        End With
End Sub

''
' Writes the "ClanCodexUpdate" message to the outgoing data buffer.
'
' @param    desc New description of the clan.
' @param    codex New codex of the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteClanCodexUpdate(ByVal Desc As String, ByRef Codex() As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ClanCodexUpdate" message to the outgoing data buffer
      '***************************************************
          Dim temp   As String
          Dim i      As Long

10        With outgoingData
20            Call .WriteByte(ClientPacketID.ClanCodexUpdate)

30            Call .WriteASCIIString(Desc)

40            For i = LBound(Codex()) To UBound(Codex())
50                temp = temp & Codex(i) & SEPARATOR
60            Next i

70            If Len(temp) Then temp = Left$(temp, Len(temp) - 1)

80            Call .WriteASCIIString(temp)
90        End With
End Sub

''
' Writes the "UserCommerceOffer" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to offer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOffer(ByVal Slot As Byte, ByVal Amount As Long, _
    ByVal OfferSlot As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UserCommerceOffer" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.UserCommerceOffer)

30            Call .WriteByte(Slot)
40            Call .WriteLong(Amount)
50            Call .WriteByte(OfferSlot)
60        End With
End Sub

Public Sub WriteCommerceChat(ByVal chat As String)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 03/12/2009
      'Writes the "CommerceChat" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.CommerceChat)

30            Call .WriteASCIIString(chat)
40        End With
End Sub


''
' Writes the "GuildAcceptPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptPeace(ByVal guild As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildAcceptPeace" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildAcceptPeace)

30            Call .WriteASCIIString(guild)
40        End With
End Sub

''
' Writes the "GuildRejectAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectAlliance(ByVal guild As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildRejectAlliance" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildRejectAlliance)

30            Call .WriteASCIIString(guild)
40        End With
End Sub

''
' Writes the "GuildRejectPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectPeace(ByVal guild As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildRejectPeace" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildRejectPeace)

30            Call .WriteASCIIString(guild)
40        End With
End Sub

''
' Writes the "GuildAcceptAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptAlliance(ByVal guild As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildAcceptAlliance" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildAcceptAlliance)

30            Call .WriteASCIIString(guild)
40        End With
End Sub

''
' Writes the "GuildOfferPeace" message to the outgoing data buffer.
'
' @param    guild The guild to whom peace is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferPeace(ByVal guild As String, ByVal proposal As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildOfferPeace" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildOfferPeace)

30            Call .WriteASCIIString(guild)
40            Call .WriteASCIIString(proposal)
50        End With
End Sub

''
' Writes the "GuildOfferAlliance" message to the outgoing data buffer.
'
' @param    guild The guild to whom an aliance is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferAlliance(ByVal guild As String, ByVal proposal As _
    String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildOfferAlliance" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildOfferAlliance)

30            Call .WriteASCIIString(guild)
40            Call .WriteASCIIString(proposal)
50        End With
End Sub

''
' Writes the "GuildAllianceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAllianceDetails(ByVal guild As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildAllianceDetails" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildAllianceDetails)

30            Call .WriteASCIIString(guild)
40        End With
End Sub

''
' Writes the "GuildPeaceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose peace proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeaceDetails(ByVal guild As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildPeaceDetails" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildPeaceDetails)

30            Call .WriteASCIIString(guild)
40        End With
End Sub

''
' Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer.
'
' @param    username The user who wants to join the guild whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestJoinerInfo(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildRequestJoinerInfo)

30            Call .WriteASCIIString(UserName)
40        End With
End Sub

''
' Writes the "GuildAlliancePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAlliancePropList()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildAlliancePropList" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GuildAlliancePropList)
End Sub

''
' Writes the "GuildPeacePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeacePropList()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildPeacePropList" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GuildPeacePropList)
End Sub

''
' Writes the "GuildDeclareWar" message to the outgoing data buffer.
'
' @param    guild The guild to which to declare war.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDeclareWar(ByVal guild As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildDeclareWar" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildDeclareWar)

30            Call .WriteASCIIString(guild)
40        End With
End Sub

''
' Writes the "GuildNewWebsite" message to the outgoing data buffer.
'
' @param    url The guild's new website's URL.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNewWebsite(ByVal URL As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildNewWebsite" message to the outgoing data buffer
      '***************************************************

          Dim AllCr As Byte
          
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildNewWebsite)
30            Call .WriteASCIIString(URL)
              
              
40        End With
End Sub

''
' Writes the "GuildAcceptNewMember" message to the outgoing data buffer.
'
' @param    username The name of the accepted player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptNewMember(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildAcceptNewMember" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildAcceptNewMember)

30            Call .WriteASCIIString(UserName)
40        End With
End Sub

''
' Writes the "GuildRejectNewMember" message to the outgoing data buffer.
'
' @param    username The name of the rejected player.
' @param    reason The reason for which the player was rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectNewMember(ByVal UserName As String, ByVal reason As _
    String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildRejectNewMember" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildRejectNewMember)

30            Call .WriteASCIIString(UserName)
40            Call .WriteASCIIString(reason)
50        End With
End Sub

''
' Writes the "GuildKickMember" message to the outgoing data buffer.
'
' @param    username The name of the kicked player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildKickMember(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildKickMember" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildKickMember)

30            Call .WriteASCIIString(UserName)
40        End With
End Sub

''
' Writes the "GuildUpdateNews" message to the outgoing data buffer.
'
' @param    news The news to be posted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildUpdateNews(ByVal news As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildUpdateNews" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildUpdateNews)

30            Call .WriteASCIIString(news)
40        End With
End Sub

''
' Writes the "GuildMemberInfo" message to the outgoing data buffer.
'
' @param    username The user whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildMemberInfo" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildMemberInfo)

30            Call .WriteASCIIString(UserName)
40        End With
End Sub

''
' Writes the "GuildOpenElections" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOpenElections()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildOpenElections" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GuildOpenElections)
End Sub

''
' Writes the "GuildRequestMembership" message to the outgoing data buffer.
'
' @param    guild The guild to which to request membership.
' @param    application The user's application sheet.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestMembership(ByVal guild As String, ByVal Application _
    As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildRequestMembership" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildRequestMembership)

30            Call .WriteASCIIString(guild)
40            Call .WriteASCIIString(Application)
50        End With
End Sub

''
' Writes the "GuildRequestDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestDetails(ByVal guild As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildRequestDetails" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildRequestDetails)

30            Call .WriteASCIIString(guild)
40        End With
End Sub

''
' Writes the "Online" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnline()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Online" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.Online)
End Sub

''
' Writes the "Quit" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteQuit()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 08/16/08
      'Writes the "Quit" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.Quit)
End Sub

''
' Writes the "GuildLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeave()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildLeave" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GuildLeave)
End Sub

''
' Writes the "RequestAccountState" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAccountState()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RequestAccountState" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.RequestAccountState)
End Sub

''
' Writes the "PetStand" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetStand()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "PetStand" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.PetStand)
End Sub

''
' Writes the "PetFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetFollow()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "PetFollow" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.PetFollow)
End Sub

''
' Writes the "ReleasePet" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReleasePet()
      '***************************************************
      'Author: ZaMa
      'Last Modification: 18/11/2009
      'Writes the "ReleasePet" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.ReleasePet)
End Sub


''
' Writes the "TrainList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainList()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "TrainList" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.TrainList)
End Sub

''
' Writes the "Rest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRest()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Rest" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.Rest)
End Sub

''
' Writes the "Meditate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditate()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Meditate" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.Meditate)
End Sub

''
' Writes the "Resucitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResucitate()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Resucitate" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.Resucitate)
End Sub

''
' Writes the "Consulta" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsulta()
      '***************************************************
      'Author: ZaMa
      'Last Modification: 01/05/2010
      'Writes the "Consulta" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.Consulta)

End Sub

''
' Writes the "Heal" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHeal()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Heal" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.Heal)
End Sub

''
' Writes the "Help" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHelp()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Help" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.Help)
End Sub

''
' Writes the "RequestStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestStats()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RequestStats" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.RequestStats)
End Sub

''
' Writes the "CommerceStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceStart()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CommerceStart" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.CommerceStart)
End Sub

''
' Writes the "BankStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankStart()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "BankStart" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.BankStart)
End Sub

''
' Writes the "Enlist" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEnlist()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Enlist" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.Enlist)
End Sub

''
' Writes the "Information" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInformation()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Information" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.Information)
End Sub

''
' Writes the "Reward" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReward()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Reward" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.Reward)
End Sub

''
' Writes the "UpTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpTime()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UpTime" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.UpTime)
End Sub


''
' Writes the "Inquiry" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiry()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Inquiry" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.Inquiry)
End Sub

''
' Writes the "GuildMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMessage(ByVal Message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildRequestDetails" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildMessage)

30            Call .WriteASCIIString(Message)
40        End With
End Sub

''
' Writes the "GroupMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGroupMessage(ByVal Message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GroupMessage" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GroupMessage)

30            Call .WriteASCIIString(Message)
40        End With
End Sub

''
' Writes the "CentinelReport" message to the outgoing data buffer.
'
' @param    number The number to report to the centinel.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCentinelReport(ByVal number As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CentinelReport" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.CentinelReport)

30            Call .WriteInteger(number)
40        End With
End Sub

''
' Writes the "GuildOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnline()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildOnline" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GuildOnline)
End Sub


''
' Writes the "CouncilMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the other council members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilMessage(ByVal Message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CouncilMessage" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.CouncilMessage)

30            Call .WriteASCIIString(Message)
40        End With
End Sub

''
' Writes the "RoleMasterRequest" message to the outgoing data buffer.
'
' @param    message The message to send to the role masters.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoleMasterRequest(ByVal Message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RoleMasterRequest" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.RoleMasterRequest)

30            Call .WriteASCIIString(Message)
40        End With
End Sub

''
' Writes the "GMRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMRequest()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GMRequest" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMRequest)
End Sub

''
' Writes the "ChangeDescription" message to the outgoing data buffer.
'
' @param    desc The new description of the user's character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeDescription(ByVal Desc As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ChangeDescription" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.ChangeDescription)

30            Call .WriteASCIIString(Desc)
40        End With
End Sub

''
' Writes the "GuildVote" message to the outgoing data buffer.
'
' @param    username The user to vote for clan leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildVote(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildVote" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildVote)

30            Call .WriteASCIIString(UserName)
40        End With
End Sub

''
' Writes the "Punishments" message to the outgoing data buffer.
'
' @param    username The user whose's  punishments are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePunishments(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Punishments" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.Punishments)

30            Call .WriteASCIIString(UserName)
40        End With
End Sub


''
' Writes the "ChangePassword" message to the outgoing data buffer.
'
' @param    oldPass Previous password.
' @param    newPass New password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangePassword(ByRef oldPass As String, ByRef newPass As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 10/10/07
      'Last Modified By: Rapsodius
      'Writes the "ChangePassword" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.ChangePassword)
              
30                Call .WriteASCIIString(oldPass)
40                Call .WriteASCIIString(newPass)
50        End With
End Sub
Public Sub WriteChangePin(ByRef oldPin As String, ByRef newPin As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 10/10/07
      'Last Modified By: Rapsodius
      'Writes the "ChangePassword" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.ChangePin)

30                Call .WriteASCIIString(oldPin)
40                Call .WriteASCIIString(newPin)
50        End With
End Sub

''
' Writes the "Gamble" message to the outgoing data buffer.
'
' @param    amount The amount to gamble.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGamble(ByVal Amount As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Gamble" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.Gamble)

30            Call .WriteInteger(Amount)
40        End With
End Sub

''
' Writes the "InquiryVote" message to the outgoing data buffer.
'
' @param    opt The chosen option to vote for.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiryVote(ByVal opt As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "InquiryVote" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.InquiryVote)

30            Call .WriteByte(opt)
40        End With
End Sub

''
' Writes the "LeaveFaction" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeaveFaction()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "LeaveFaction" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.LeaveFaction)
End Sub
Public Sub WriteBankExtractGold(ByVal Amount As Long)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "BankExtractGold" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.BankExtractGold)

30            Call .WriteLong(Amount)
40        End With
End Sub

''
' Writes the "BankDepositGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to deposit in the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDepositGold(ByVal Amount As Long)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "BankDepositGold" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.BankDepositGold)

30            Call .WriteLong(Amount)
40        End With
End Sub

Public Sub WriteSubirCanje()
    With outgoingData
        Call .WriteByte(ClientPacketID.SubirCanjes)
    End With
End Sub

''
' Writes the "Denounce" message to the outgoing data buffer.
'
' @param    message The message to send with the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDenounce(ByVal Message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Denounce" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.Denounce)

30            Call .WriteASCIIString(Message)
40        End With
End Sub

''
' Writes the "GuildFundate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildFundate()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 03/21/2001
      'Writes the "GuildFundate" message to the outgoing data buffer
      '14/12/2009: ZaMa - Now first checks if the user can foundate a guild.
      '03/21/2001: Pato - Deleted de clanType param.
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GuildFundate)
End Sub

''
' Writes the "GuildFundation" message to the outgoing data buffer.
'
' @param    clanType The alignment of the clan to be founded.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildFundation(ByVal clanType As eClanType)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 14/12/2009
      'Writes the "GuildFundation" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildFundation)

30            Call .WriteByte(clanType)
40        End With
End Sub


''
' Writes the "GuildMemberList" message to the outgoing data buffer.
'
' @param    guild The guild whose member list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberList(ByVal guild As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildMemberList" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.GuildMemberList)

40            Call .WriteASCIIString(guild)
50        End With
End Sub

''
' Writes the "InitCrafting" message to the outgoing data buffer.
'
' @param    Cantidad The final aumont of item to craft.
' @param    NroPorCiclo The amount of items to craft per cicle.

Public Sub WriteInitCrafting(ByVal cantidad As Long, ByVal NroPorCiclo As _
    Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 29/01/2010
      'Writes the "InitCrafting" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.InitCrafting)
30            Call .WriteLong(cantidad)

40            Call .WriteInteger(NroPorCiclo)
50        End With
End Sub

''
' Writes the "Home" message to the outgoing data buffer.
'
Public Sub WriteHome()
      '***************************************************
      'Author: Budi
      'Last Modification: 01/06/10
      'Writes the "Home" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.Home)
30        End With
End Sub



''
' Writes the "GMMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to the other GMs online.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMMessage(ByVal Message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GMMessage" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.GMMessage)
40            Call .WriteASCIIString(Message)
50        End With
End Sub

''
' Writes the "ShowName" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowName()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ShowName" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.showName)
End Sub

''
' Writes the "OnlineRoyalArmy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineRoyalArmy()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "OnlineRoyalArmy" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.OnlineRoyalArmy)
End Sub

''
' Writes the "OnlineChaosLegion" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineChaosLegion()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "OnlineChaosLegion" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.OnlineChaosLegion)
End Sub

''
' Writes the "GoNearby" message to the outgoing data buffer.
'
' @param    username The suer to approach.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoNearby(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GoNearby" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call outgoingData.WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.GoNearby)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub
Public Sub WriteGoSeBuscaa(ByVal UserName As String)
10        With outgoingData
20            Call outgoingData.WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.SeBusca)
40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "Comment" message to the outgoing data buffer.
'
' @param    message The message to leave in the log as a comment.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteComment(ByVal Message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Comment" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.comment)

40            Call .WriteASCIIString(Message)
50        End With
End Sub

''
' Writes the "ServerTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerTime()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ServerTime" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.serverTime)
End Sub

''
' Writes the "Where" message to the outgoing data buffer.
'
' @param    username The user whose position is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhere(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Where" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.Where)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "CreaturesInMap" message to the outgoing data buffer.
'
' @param    map The map in which to check for the existing creatures.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreaturesInMap(ByVal map As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CreaturesInMap" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.CreaturesInMap)

40            Call .WriteInteger(map)
50        End With
End Sub

''
' Writes the "WarpMeToTarget" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpMeToTarget()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "WarpMeToTarget" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.WarpMeToTarget)
End Sub

''
' Writes the "WarpChar" message to the outgoing data buffer.
'
' @param    username The user to be warped. "YO" represent's the user's char.
' @param    map The map to which to warp the character.
' @param    x The x position in the map to which to waro the character.
' @param    y The y position in the map to which to waro the character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpChar(ByVal UserName As String, ByVal map As Integer, ByVal _
    X As Byte, ByVal Y As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "WarpChar" message to the outgoing data buffer
      '***************************************************

10        If Not esGM(UserCharIndex) Then Exit Sub
          
20        With outgoingData
30            Call .WriteByte(ClientPacketID.GMCommands)
40            Call .WriteByte(eGMCommands.WarpChar)

50            Call .WriteASCIIString(UserName)

60            Call .WriteInteger(map)

70            Call .WriteByte(X)
80            Call .WriteByte(Y)
90        End With
End Sub

''
' Writes the "Silence" message to the outgoing data buffer.
'
' @param    username The user to silence.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSilence(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Silence" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.Silence)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "SOSShowList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSShowList()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "SOSShowList" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.SOSShowList)
End Sub

''
' Writes the "SOSRemove" message to the outgoing data buffer.
'
' @param    username The user whose SOS call has been already attended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSRemove(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "SOSRemove" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.SOSRemove)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "GoToChar" message to the outgoing data buffer.
'
' @param    username The user to be approached.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoToChar(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GoToChar" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.GoToChar)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "invisible" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInvisible()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "invisible" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.Invisible)
End Sub

''
' Writes the "GMPanel" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMPanel()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GMPanel" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.GMPanel)
End Sub

''
' Writes the "RequestUserList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestUserList()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RequestUserList" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.RequestUserList)
End Sub

''
' Writes the "Working" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorking()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Working" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.Working)
End Sub

''
' Writes the "Hiding" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHiding()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Hiding" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.Hiding)
End Sub

''
' Writes the "Jail" message to the outgoing data buffer.
'
' @param    username The user to be sent to jail.
' @param    reason The reason for which to send him to jail.
' @param    time The time (in minutes) the user will have to spend there.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteJail(ByVal UserName As String, ByVal reason As String, ByVal _
    Time As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Jail" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.Jail)

40            Call .WriteASCIIString(UserName)
50            Call .WriteASCIIString(reason)

60            Call .WriteByte(Time)
70        End With
End Sub

''
' Writes the "KillNPC" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPC()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "KillNPC" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.KillNPC)
End Sub

''
' Writes the "WarnUser" message to the outgoing data buffer.
'
' @param    username The user to be warned.
' @param    reason Reason for the warning.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarnUser(ByVal UserName As String, ByVal reason As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "WarnUser" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.WarnUser)

40            Call .WriteASCIIString(UserName)
50            Call .WriteASCIIString(reason)
60        End With
End Sub

''
' Writes the "RequestCharInfo" message to the outgoing data buffer.
'
' @param    username The user whose information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInfo(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RequestCharInfo" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.RequestCharInfo)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "RequestCharStats" message to the outgoing data buffer.
'
' @param    username The user whose stats are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharStats(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RequestCharStats" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.RequestCharStats)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "RequestCharGold" message to the outgoing data buffer.
'
' @param    username The user whose gold is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharGold(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RequestCharGold" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.RequestCharGold)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "RequestCharInventory" message to the outgoing data buffer.
'
' @param    username The user whose inventory is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInventory(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RequestCharInventory" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.RequestCharInventory)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "RequestCharBank" message to the outgoing data buffer.
'
' @param    username The user whose banking information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharBank(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RequestCharBank" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.RequestCharBank)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "RequestCharSkills" message to the outgoing data buffer.
'
' @param    username The user whose skills are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharSkills(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RequestCharSkills" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.RequestCharSkills)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "ReviveChar" message to the outgoing data buffer.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReviveChar(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ReviveChar" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ReviveChar)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "OnlineGM" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineGM()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "OnlineGM" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.OnlineGM)
End Sub

''
' Writes the "OnlineMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineMap(ByVal map As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 26/03/2009
      'Writes the "OnlineMap" message to the outgoing data buffer
      '26/03/2009: Now you don't need to be in the map to use the comand, so you send the map to server
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.OnlineMap)

40            Call .WriteInteger(map)
50        End With
End Sub

''
' Writes the "Forgive" message to the outgoing data buffer.
'
' @param    username The user to be forgiven.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForgive(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Forgive" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.Forgive)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "Kick" message to the outgoing data buffer.
'
' @param    username The user to be kicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKick(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Kick" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.Kick)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "Execute" message to the outgoing data buffer.
'
' @param    username The user to be executed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteExecute(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Execute" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.Execute)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "BanChar" message to the outgoing data buffer.
'
' @param    username The user to be banned.
' @param    reason The reson for which the user is to be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanChar(ByVal UserName As String, ByVal reason As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "BanChar" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.banChar)

40            Call .WriteASCIIString(UserName)

50            Call .WriteASCIIString(reason)
60        End With
End Sub

''
' Writes the "UnbanChar" message to the outgoing data buffer.
'
' @param    username The user to be unbanned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanChar(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UnbanChar" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.UnbanChar)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "NPCFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCFollow()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "NPCFollow" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.NPCFollow)
End Sub

''
' Writes the "SummonChar" message to the outgoing data buffer.
'
' @param    username The user to be summoned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSummonChar(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "SummonChar" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.SummonChar)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "SpawnListRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnListRequest()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "SpawnListRequest" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.SpawnListRequest)
End Sub

''
' Writes the "SpawnCreature" message to the outgoing data buffer.
'
' @param    creatureIndex The index of the creature in the spawn list to be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "SpawnCreature" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.SpawnCreature)

40            Call .WriteInteger(creatureIndex)
50        End With
End Sub

''
' Writes the "ResetNPCInventory" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetNPCInventory()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ResetNPCInventory" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.ResetNPCInventory)
End Sub

''
' Writes the "CleanWorld" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanWorld()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CleanWorld" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.cleanworld)
End Sub

''
' Writes the "ServerMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerMessage(ByVal Message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ServerMessage" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ServerMessage)

40            Call .WriteASCIIString(Message)
50        End With
End Sub
Public Sub WriteRolMensaje(ByVal Message As String)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.RolMensaje)

40            Call .WriteASCIIString(Message)
50        End With
End Sub

''
' Writes the "MapMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMapMessage(ByVal Message As String)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 14/11/2010
      'Writes the "MapMessage" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.MapMessage)

40            Call .WriteASCIIString(Message)
50        End With
End Sub

Public Sub WriteImpersonate()
      '***************************************************
      'Author: ZaMa
      'Last Modification: 20/11/2010
      'Writes the "Impersonate" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.Impersonate)
End Sub

''
' Writes the "Imitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImitate()
      '***************************************************
      'Author: ZaMa
      'Last Modification: 20/11/2010
      'Writes the "Imitate" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.Imitate)
End Sub


''
' Writes the "NickToIP" message to the outgoing data buffer.
'
' @param    username The user whose IP is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNickToIP(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "NickToIP" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.nickToIP)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "IPToNick" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIPToNick(ByRef Ip() As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "IPToNick" message to the outgoing data buffer
      '***************************************************
10        If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP

          Dim i      As Long

20        With outgoingData
30            Call .WriteByte(ClientPacketID.GMCommands)
40            Call .WriteByte(eGMCommands.IPToNick)

50            For i = LBound(Ip()) To UBound(Ip())
60                Call .WriteByte(Ip(i))
70            Next i
80        End With
End Sub

''
' Writes the "GuildOnlineMembers" message to the outgoing data buffer.
'
' @param    guild The guild whose online player list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnlineMembers(ByVal guild As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildOnlineMembers" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.GuildOnlineMembers)

40            Call .WriteASCIIString(guild)
50        End With
End Sub

''
' Writes the "TeleportCreate" message to the outgoing data buffer.
'
' @param    map the map to which the teleport will lead.
' @param    x The position in the x axis to which the teleport will lead.
' @param    y The position in the y axis to which the teleport will lead.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportCreate(ByVal map As Integer, ByVal X As Byte, ByVal Y _
    As Byte, Optional ByVal Radio As Byte = 0)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "TeleportCreate" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.TeleportCreate)

40            Call .WriteInteger(map)

50            Call .WriteByte(X)
60            Call .WriteByte(Y)

70            Call .WriteByte(Radio)
80        End With
End Sub

''
' Writes the "TeleportDestroy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportDestroy()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "TeleportDestroy" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.TeleportDestroy)
End Sub

''
' Writes the "RainToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RainToggle" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.RainToggle)
End Sub

''
' Writes the "SetCharDescription" message to the outgoing data buffer.
'
' @param    desc The description to set to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetCharDescription(ByVal Desc As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "SetCharDescription" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.SetCharDescription)

40            Call .WriteASCIIString(Desc)
50        End With
End Sub

''
' Writes the "ForceMIDIToMap" message to the outgoing data buffer.
'
' @param    midiID The ID of the midi file to play.
' @param    map The map in which to play the given midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIToMap(ByVal midiID As Byte, ByVal map As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ForceMIDIToMap" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ForceMIDIToMap)

40            Call .WriteByte(midiID)

50            Call .WriteInteger(map)
60        End With
End Sub

''
' Writes the "ForceWAVEToMap" message to the outgoing data buffer.
'
' @param    waveID  The ID of the wave file to play.
' @param    Map     The map into which to play the given wave.
' @param    x       The position in the x axis in which to play the given wave.
' @param    y       The position in the y axis in which to play the given wave.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEToMap(ByVal waveID As Byte, ByVal map As Integer, _
    ByVal X As Byte, ByVal Y As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ForceWAVEToMap" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ForceWAVEToMap)

40            Call .WriteByte(waveID)

50            Call .WriteInteger(map)

60            Call .WriteByte(X)
70            Call .WriteByte(Y)
80        End With
End Sub

''
' Writes the "RoyalArmyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyMessage(ByVal Message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RoyalArmyMessage" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.RoyalArmyMessage)

40            Call .WriteASCIIString(Message)
50        End With
End Sub

''
' Writes the "ChaosLegionMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the chaos legion member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionMessage(ByVal Message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ChaosLegionMessage" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ChaosLegionMessage)

40            Call .WriteASCIIString(Message)
50        End With
End Sub

''
' Writes the "CitizenMessage" message to the outgoing data buffer.
'
' @param    message The message to send to citizens.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCitizenMessage(ByVal Message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CitizenMessage" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.CitizenMessage)

40            Call .WriteASCIIString(Message)
50        End With
End Sub

''
' Writes the "CriminalMessage" message to the outgoing data buffer.
'
' @param    message The message to send to criminals.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCriminalMessage(ByVal Message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CriminalMessage" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.CriminalMessage)

40            Call .WriteASCIIString(Message)
50        End With
End Sub

''
' Writes the "TalkAsNPC" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalkAsNPC(ByVal Message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "TalkAsNPC" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.TalkAsNPC)

40            Call .WriteASCIIString(Message)
50        End With
End Sub

''
' Writes the "DestroyAllItemsInArea" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyAllItemsInArea()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "DestroyAllItemsInArea" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.DestroyAllItemsInArea)
End Sub

''
' Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted into the royal army council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptRoyalCouncilMember(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.AcceptRoyalCouncilMember)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted as a chaos council member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptChaosCouncilMember(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.AcceptChaosCouncilMember)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "ItemsInTheFloor" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemsInTheFloor()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ItemsInTheFloor" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.ItemsInTheFloor)
End Sub

''
' Writes the "MakeDumb" message to the outgoing data buffer.
'
' @param    username The name of the user to be made dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumb(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "MakeDumb" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.MakeDumb)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "MakeDumbNoMore" message to the outgoing data buffer.
'
' @param    username The name of the user who will no longer be dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumbNoMore(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "MakeDumbNoMore" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.MakeDumbNoMore)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "DumpIPTables" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumpIPTables()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "DumpIPTables" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.dumpIPTables)
End Sub

''
' Writes the "CouncilKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilKick(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CouncilKick" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.CouncilKick)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub
''
' Writes the "CheckHD" message to the outgoing data buffer.
'
'@param   username The name of the user to be checked.
Public Sub WriteCheckHD(ByVal UserName As String)
      '***************************************************
      'Author: ArzenaTh
      'Last Modification: 01/09/10
      'Checkeamos la HD del usuario.
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.CheckHD)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub
''
' Writes the "BanHD" message to the outgoing data buffer
'
'@param    username The name of the user to be banned.
Public Sub WriteBanHD(ByVal UserName As String)
      '***************************************************
      'Author: ArzenaTh
      'Last Modification: 01/09/10
      'Baneamos la HD del usuario.
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.BanHD)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "UnBanHD" message to the outgoing data buffer
'
'@param    username The name of the user to be unbanned.
Public Sub WriteUnBanHD(ByVal HD As String)
      '***************************************************
      'Author: ArzenaTh
      'Last Modification: 01/09/10
      'Unbaneamos al usuario con su HD baneado.
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.UnBanHD)

40            Call .WriteASCIIString(HD)
50        End With
End Sub
Public Sub WriteLookProcess(ByVal data As String)
      '***************************************************
      'Author: Franco Emmanuel Giménez (Franeg95)
      'Last Modification: 18/10/10
      'Writes the "Lookprocess" message and write the nickname of another user to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.Lookprocess)
30            Call .WriteASCIIString(data)
40        End With
End Sub


Public Sub WriteSendProcessList()
      '***************************************************
      'Author: Franco Emmanuel Giménez (Franeg95)
      'Last Modification: 18/10/10
      'Writes the "SendProcessList" message and write the process list of another user to the outgoing data buffer
      '***************************************************
        
      Dim TempStr As String
      
      TempStr = Replace(LstPscGS, " ", ".")
      
10        With outgoingData
20            Call .WriteByte(ClientPacketID.SendProcessList)
30            Call .WriteASCIIString(TempStr)
40        End With
End Sub

Private Sub HandleSeeInProcess()
10        Call incomingData.ReadByte

20        Call WriteSendProcessList

End Sub

''
' Writes the "SetTrigger" message to the outgoing data buffer.
'
' @param    trigger The type of trigger to be set to the tile.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "SetTrigger" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.SetTrigger)

40            Call .WriteByte(Trigger)
50        End With
End Sub

''
' Writes the "AskTrigger" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAskTrigger()
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 04/13/07
      'Writes the "AskTrigger" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.AskTrigger)
End Sub

''
' Writes the "BannedIPList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPList()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "BannedIPList" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.BannedIPList)
End Sub

''
' Writes the "BannedIPReload" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPReload()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "BannedIPReload" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.BannedIPReload)
End Sub

''
' Writes the "GuildBan" message to the outgoing data buffer.
'
' @param    guild The guild whose members will be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildBan(ByVal guild As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "GuildBan" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.GuildBan)

40            Call .WriteASCIIString(guild)
50        End With
End Sub

''
' Writes the "BanIP" message to the outgoing data buffer.
'
' @param    byIp    If set to true, we are banning by IP, otherwise the ip of a given character.
' @param    IP      The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @param    nick    The nick of the player whose ip will be banned.
' @param    reason  The reason for the ban.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanIP(ByVal byIp As Boolean, ByRef Ip() As Byte, ByVal Nick As _
    String, ByVal reason As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "BanIP" message to the outgoing data buffer
      '***************************************************
10        If byIp And UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP

          Dim i      As Long

20        With outgoingData
30            Call .WriteByte(ClientPacketID.GMCommands)
40            Call .WriteByte(eGMCommands.BanIP)

50            Call .WriteBoolean(byIp)

60            If byIp Then
70                For i = LBound(Ip()) To UBound(Ip())
80                    Call .WriteByte(Ip(i))
90                Next i
100           Else
110               Call .WriteASCIIString(Nick)
120           End If

130           Call .WriteASCIIString(reason)
140       End With
End Sub

''
' Writes the "UnbanIP" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanIP(ByRef Ip() As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "UnbanIP" message to the outgoing data buffer
      '***************************************************
10        If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP

          Dim i      As Long

20        With outgoingData
30            Call .WriteByte(ClientPacketID.GMCommands)
40            Call .WriteByte(eGMCommands.UnbanIP)

50            For i = LBound(Ip()) To UBound(Ip())
60                Call .WriteByte(Ip(i))
70            Next i
80        End With
End Sub

''
' Writes the "CreateItem" message to the outgoing data buffer.
'
' @param    itemIndex The index of the item to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateItem(ByVal ItemIndex As Long, ByVal cantidad As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 11/02/11

10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.CreateItem)
40            Call .WriteInteger(ItemIndex)
50            Call .WriteInteger(cantidad)
60        End With
End Sub

''
' Writes the "DestroyItems" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyItems()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "DestroyItems" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.DestroyItems)
End Sub

''
' Writes the "ChaosLegionKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Chaos Legion.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionKick(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ChaosLegionKick" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ChaosLegionKick)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "RoyalArmyKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Royal Army.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyKick(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RoyalArmyKick" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.RoyalArmyKick)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "ForceMIDIAll" message to the outgoing data buffer.
'
' @param    midiID The id of the midi file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIAll(ByVal midiID As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ForceMIDIAll" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ForceMIDIAll)

40            Call .WriteByte(midiID)
50        End With
End Sub

''
' Writes the "ForceWAVEAll" message to the outgoing data buffer.
'
' @param    waveID The id of the wave file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEAll(ByVal waveID As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ForceWAVEAll" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ForceWAVEAll)

40            Call .WriteByte(waveID)
50        End With
End Sub

''
' Writes the "RemovePunishment" message to the outgoing data buffer.
'
' @param    username The user whose punishments will be altered.
' @param    punishment The id of the punishment to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemovePunishment(ByVal UserName As String, ByVal punishment As _
    Byte, ByVal NewText As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RemovePunishment" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.RemovePunishment)

40            Call .WriteASCIIString(UserName)
50            Call .WriteByte(punishment)
60            Call .WriteASCIIString(NewText)
70        End With
End Sub

''
' Writes the "TileBlockedToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTileBlockedToggle()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "TileBlockedToggle" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.TileBlockedToggle)
End Sub

''
' Writes the "KillNPCNoRespawn" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPCNoRespawn()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "KillNPCNoRespawn" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.KillNPCNoRespawn)
End Sub

''
' Writes the "KillAllNearbyNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillAllNearbyNPCs()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "KillAllNearbyNPCs" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.KillAllNearbyNPCs)
End Sub

''
' Writes the "LastIP" message to the outgoing data buffer.
'
' @param    username The user whose last IPs are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLastIP(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "LastIP" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.lastip)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub
Public Sub WriteSystemMessage(ByVal Message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "SystemMessage" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.SystemMessage)

40            Call .WriteASCIIString(Message)
50        End With
End Sub

''
' Writes the "CreateNPC" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPC(ByVal NPCIndex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CreateNPC" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.CreateNPC)

40            Call .WriteInteger(NPCIndex)
50        End With
End Sub

''
' Writes the "CreateNPCWithRespawn" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPCWithRespawn(ByVal NPCIndex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CreateNPCWithRespawn" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.CreateNPCWithRespawn)

40            Call .WriteInteger(NPCIndex)
50        End With
End Sub

''
' Writes the "ImperialArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of imperial armour to be altered.
' @param    objectIndex The index of the new object to be set as the imperial armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImperialArmour(ByVal armourIndex As Byte, ByVal objectIndex As _
    Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ImperialArmour" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ImperialArmour)

40            Call .WriteByte(armourIndex)

50            Call .WriteInteger(objectIndex)
60        End With
End Sub

''
' Writes the "ChaosArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of chaos armour to be altered.
' @param    objectIndex The index of the new object to be set as the chaos armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosArmour(ByVal armourIndex As Byte, ByVal objectIndex As _
    Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ChaosArmour" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ChaosArmour)

40            Call .WriteByte(armourIndex)

50            Call .WriteInteger(objectIndex)
60        End With
End Sub

''
' Writes the "NavigateToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "NavigateToggle" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.NavigateToggle)
End Sub

''
' Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerOpenToUsersToggle()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.ServerOpenToUsersToggle)
End Sub

''
' Writes the "TurnOffServer" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnOffServer()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "TurnOffServer" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.TurnOffServer)
End Sub

''
' Writes the "TurnCriminal" message to the outgoing data buffer.
'
' @param    username The name of the user to turn into criminal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnCriminal(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "TurnCriminal" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.TurnCriminal)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub
Public Sub WriteResetFactionCaos(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ResetFactions" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ResetFactionCaos)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "ResetFactions" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any faction.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetFactionReal(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ResetFactions" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ResetFactionReal)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "RemoveCharFromGuild" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RemoveCharFromGuild" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.RemoveCharFromGuild)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "RequestCharMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharMail(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "RequestCharMail" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.RequestCharMail)

40            Call .WriteASCIIString(UserName)
50        End With
End Sub

''
' Writes the "AlterPassword" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    copyFrom The name of the user from which to copy the password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterPassword(ByVal UserName As String, ByVal CopyFrom As _
    String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "AlterPassword" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.AlterPassword)

40            Call .WriteASCIIString(UserName)
50            Call .WriteASCIIString(CopyFrom)
60        End With
End Sub

''
' Writes the "AlterMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newMail The new email of the player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterMail(ByVal UserName As String, ByVal newMail As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "AlterMail" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.AlterMail)

40            Call .WriteASCIIString(UserName)
50            Call .WriteASCIIString(newMail)
60        End With
End Sub

''
' Writes the "AlterName" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newName The new user name.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterName(ByVal UserName As String, ByVal newName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "AlterName" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.AlterName)

40            Call .WriteASCIIString(UserName)
50            Call .WriteASCIIString(newName)
60        End With
End Sub

''
' Writes the "ToggleCentinelActivated" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteToggleCentinelActivated()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ToggleCentinelActivated" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.ToggleCentinelActivated)
End Sub

''
' Writes the "DoBackup" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoBackup()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "DoBackup" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.DoBackUp)
End Sub

''
' Writes the "ShowGuildMessages" message to the outgoing data buffer.
'
' @param    guild The guild to listen to.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildMessages(ByVal guild As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ShowGuildMessages" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ShowGuildMessages)

40            Call .WriteASCIIString(guild)
50        End With
End Sub

''
' Writes the "SaveMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveMap()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "SaveMap" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.SaveMap)
End Sub

''
' Writes the "ChangeMapInfoPK" message to the outgoing data buffer.
'
' @param    isPK True if the map is PK, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ChangeMapInfoPK" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ChangeMapInfoPK)

40            Call .WriteBoolean(isPK)
50        End With
End Sub

''
' Writes the "ChangeMapInfoBackup" message to the outgoing data buffer.
'
' @param    backup True if the map is to be backuped, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ChangeMapInfoBackup" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ChangeMapInfoBackup)

40            Call .WriteBoolean(backup)
50        End With
End Sub

''
' Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer.
'
' @param    restrict NEWBIES (only newbies), NO (everyone), ARMADA (just Armadas), CAOS (just caos) or FACCION (Armadas & caos only)
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)
      '***************************************************
      'Author: Pablo (ToxicWaste)
      'Last Modification: 26/01/2007
      'Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ChangeMapInfoRestricted)

40            Call .WriteASCIIString(restrict)
50        End With
End Sub

''
' Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer.
'
' @param    nomagic TRUE if no magic is to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)
      '***************************************************
      'Author: Pablo (ToxicWaste)
      'Last Modification: 26/01/2007
      'Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ChangeMapInfoNoMagic)

40            Call .WriteBoolean(nomagic)
50        End With
End Sub

''
' Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer.
'
' @param    noinvi TRUE if invisibility is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)
      '***************************************************
      'Author: Pablo (ToxicWaste)
      'Last Modification: 26/01/2007
      'Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ChangeMapInfoNoInvi)

40            Call .WriteBoolean(noinvi)
50        End With
End Sub

''
' Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer.
'
' @param    noresu TRUE if resurection is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)
      '***************************************************
      'Author: Pablo (ToxicWaste)
      'Last Modification: 26/01/2007
      'Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ChangeMapInfoNoResu)

40            Call .WriteBoolean(noresu)
50        End With
End Sub

''
' Writes the "ChangeMapInfoLand" message to the outgoing data buffer.
'
' @param    land options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoLand(ByVal land As String)
      '***************************************************
      'Author: Pablo (ToxicWaste)
      'Last Modification: 26/01/2007
      'Writes the "ChangeMapInfoLand" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ChangeMapInfoLand)

40            Call .WriteASCIIString(land)
50        End With
End Sub

''
' Writes the "ChangeMapInfoZone" message to the outgoing data buffer.
'
' @param    zone options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoZone(ByVal zone As String)
      '***************************************************
      'Author: Pablo (ToxicWaste)
      'Last Modification: 26/01/2007
      'Writes the "ChangeMapInfoZone" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ChangeMapInfoZone)

40            Call .WriteASCIIString(zone)
50        End With
End Sub

''
' Writes the "ChangeMapInfoNoOcultar" message to the outgoing data buffer.
'
' @param    PermitirOcultar True if the map permits to hide, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoOcultar(ByVal PermitirOcultar As Boolean)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 19/09/2010
      'Writes the "ChangeMapInfoNoOcultar" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ChangeMapInfoNoOcultar)

40            Call .WriteBoolean(PermitirOcultar)
50        End With
End Sub

''
' Writes the "ChangeMapInfoNoInvocar" message to the outgoing data buffer.
'
' @param    PermitirInvocar True if the map permits to invoke, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvocar(ByVal PermitirInvocar As Boolean)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 18/09/2010
      'Writes the "ChangeMapInfoNoInvocar" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ChangeMapInfoNoInvocar)

40            Call .WriteBoolean(PermitirInvocar)
50        End With
End Sub

''
' Writes the "ChangeMapInfoStealNpc" message to the outgoing data buffer.
'
' @param    forbid TRUE if stealNpc forbiden.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoStealNpc(ByVal forbid As Boolean)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 25/07/2010
      'Writes the "ChangeMapInfoStealNpc" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ChangeMapInfoStealNpc)

40            Call .WriteBoolean(forbid)
50        End With
End Sub

''
' Writes the "SaveChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveChars()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "SaveChars" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.SaveChars)
End Sub

''
' Writes the "CleanSOS" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanSOS()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "CleanSOS" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.CleanSOS)
End Sub

''
' Writes the "ShowServerForm" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowServerForm()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ShowServerForm" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.ShowServerForm)
End Sub

''
' Writes the "Night" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNight()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Night" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.night)
End Sub

''
' Writes the "KickAllChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKickAllChars()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "KickAllChars" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.KickAllChars)
End Sub

''
' Writes the "ReloadNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadNPCs()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ReloadNPCs" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.ReloadNPCs)
End Sub

''
' Writes the "ReloadServerIni" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadServerIni()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ReloadServerIni" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.ReloadServerIni)
End Sub

''
' Writes the "ReloadSpells" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadSpells()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ReloadSpells" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.ReloadSpells)
End Sub

''
' Writes the "ReloadObjects" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadObjects()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ReloadObjects" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.ReloadObjects)
End Sub

''
' Writes the "Restart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestart()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Restart" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.Restart)
End Sub

''
' Writes the "ResetAutoUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetAutoUpdate()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ResetAutoUpdate" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.ResetAutoUpdate)
End Sub

''
' Writes the "ChatColor" message to the outgoing data buffer.
'
' @param    r The red component of the new chat color.
' @param    g The green component of the new chat color.
' @param    b The blue component of the new chat color.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatColor(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "ChatColor" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ChatColor)

40            Call .WriteByte(r)
50            Call .WriteByte(g)
60            Call .WriteByte(b)
70        End With
End Sub

''
' Writes the "Ignored" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIgnored()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Ignored" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.Ignored)
End Sub

''
' Writes the "CheckSlot" message to the outgoing data buffer.
'
' @param    UserName    The name of the char whose slot will be checked.
' @param    slot        The slot to be checked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCheckSlot(ByVal UserName As String, ByVal Slot As Byte)
      '***************************************************
      'Author: Pablo (ToxicWaste)
      'Last Modification: 26/01/2007
      'Writes the "CheckSlot" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.CheckSlot)
40            Call .WriteASCIIString(UserName)
50            Call .WriteByte(Slot)
60        End With
End Sub

''
' Writes the "Ping" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePing()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 26/01/2007
      'Writes the "Ping" message to the outgoing data buffer
      '***************************************************

          'Prevent the timer from being cut
10        If pingTime <> 0 Then Exit Sub
          
20        Call outgoingData.WriteByte(ClientPacketID.Ping)

          ' Avoid computing errors due to frame rate
30        Call FlushBuffer
40        DoEvents

50        pingTime = GetTickCount
End Sub

''
' Writes the "ShareNpc" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShareNpc()
      '***************************************************
      'Author: ZaMa
      'Last Modification: 15/04/2010
      'Writes the "ShareNpc" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.ShareNpc)
End Sub

''
' Writes the "StopSharingNpc" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteStopSharingNpc()
      '***************************************************
      'Author: ZaMa
      'Last Modification: 15/04/2010
      'Writes the "StopSharingNpc" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.StopSharingNpc)
End Sub

''
' Writes the "SetIniVar" message to the outgoing data buffer.
'
' @param    sLlave the name of the key which contains the value to edit
' @param    sClave the name of the value to edit
' @param    sValor the new value to set to sClave
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetIniVar(ByRef sLlave As String, ByRef sClave As String, ByRef _
    sValor As String)
      '***************************************************
      'Author: Brian Chaia (BrianPr)
      'Last Modification: 21/06/2009
      'Writes the "SetIniVar" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.SetIniVar)

40            Call .WriteASCIIString(sLlave)
50            Call .WriteASCIIString(sClave)
60            Call .WriteASCIIString(sValor)
70        End With
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Sends all data existing in the buffer
      '***************************************************
          Dim sndData As String

10        With outgoingData
20            If .Length = 0 Then Exit Sub

30            sndData = .ReadASCIIStringFixed(.Length)

            
40            Call SendData(sndData)
50        End With
End Sub

''
' Sends the data using the socket controls in the MainForm.
'
' @param    sdData  The data to be sent to the server.
    
Private Sub SendData(ByRef sdData As String)

    #If UsarWrench = 1 Then
10            If Not frmMain.Socket1.IsWritable Then
                  'Put data back in the bytequeue
20                Call outgoingData.WriteASCIIStringFixed(sdData)
30                Exit Sub
40            End If

50            If Not frmMain.Socket1.Connected Then Exit Sub
    #Else
60            If frmMain.Winsock1.State <> sckConnected Then Exit Sub
    #End If

    #If UsarWrench = 1 Then
70            Call frmMain.Socket1.Write(sdData, Len(sdData))
    #Else
80            Call frmMain.Winsock1.SendData(sdData)
    #End If

End Sub
Public Sub WriteRequieredCaptions(ByVal UserName As String)
10        Call outgoingData.WriteByte(ClientPacketID.rCaptions)
20        Call outgoingData.WriteASCIIString(UserName)
End Sub

Public Sub WriteSendCaptions()
10        Call outgoingData.WriteByte(ClientPacketID.SCaptions)
20        Call outgoingData.WriteASCIIString(Vercaptions.Listar)
30        Call outgoingData.WriteByte(Vercaptions.CANTv)
40        Vercaptions.CANTv = 0
End Sub

Private Sub HandleRequieredCaptions()
10        Call incomingData.ReadByte
20        WriteSendCaptions
End Sub

Private Sub HandleShowCaptions()
          Dim miBuffer As New clsByteQueue

10        Call miBuffer.CopyBuffer(incomingData)

20        Call miBuffer.ReadByte

          Dim QeName As String
          Dim QeList As String
          Dim QeCANT As Byte
          Dim b      As Long
30        QeName = miBuffer.ReadASCIIString
40        QeList = miBuffer.ReadASCIIString
50        QeCANT = miBuffer.ReadByte
60        Vercaptions.List1.Clear
70        For b = 1 To QeCANT
80            Vercaptions.List1.AddItem ReadField(b, QeList, Asc("#"))
90        Next b
100       Vercaptions.Show
110       Vercaptions.Label1.Caption = "Captions de " & QeName
120       Call incomingData.CopyBuffer(miBuffer)
End Sub
Public Sub WriteGlobalStatus()
      '***************************************************
      'Author: Martín Gomez (Samke)
      'Last Modification: 10/03/2012
      'Writes the "GlobalStatus" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.GlobalStatus)
End Sub

''
' Writes the "GlobalMessage" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGlobalMessage(ByVal Message As String)
      '***************************************************
      'Author: Martín Gomez (Samke)
      'Last Modification: 10/03/2012
      'Writes the "GlobalMessage" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GlobalMessage)
30            Call .WriteASCIIString(Message)
40        End With
End Sub
Public Sub WriteCuentaRegresiva(ByVal Second As Byte)

10        With outgoingData
20            Call .WriteByte(ClientPacketID.CuentaRegresiva)
30            Call .WriteByte(Second)
40        End With
End Sub
Public Sub WriteFianzah(ByVal Fianza As Long)

10        With outgoingData
20            Call .WriteByte(ClientPacketID.Fianzah)

30            Call .WriteLong(Fianza)
40        End With
End Sub

Public Sub WriteDragToPos(ByVal X As Byte, ByVal Y As Byte, ByVal Slot As Byte, _
    ByVal Amount As Integer)
        
         If Slot > 25 Then Exit Sub
         
10        With outgoingData
20            .WriteByte ClientPacketID.DragToPos
30            .WriteByte X
40            .WriteByte Y
50            .WriteByte Slot
60            .WriteInteger Amount
70        End With

End Sub

Public Sub WriteDragInventory(ByVal originalSlot As Integer, ByVal newSlot As _
    Integer, ByVal moveType As eMoveType)
      '***************************************************
      'Author: Budi
      'Last Modification: 05/01/2011
      'Writes the "MoveItem" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.dragInventory)
30            Call .WriteByte(originalSlot)
40            Call .WriteByte(newSlot)
50            Call .WriteByte(moveType)
60        End With
End Sub

Public Sub WriteDragToggle()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "SafeToggle" message to the outgoing data buffer
      '***************************************************
10        Call outgoingData.WriteByte(ClientPacketID.DragToggle)
End Sub

Public Sub WriteCambioPj(ByVal UserName1 As String, ByVal UserName2 As String)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.CambioPj)

40            Call .WriteASCIIString(UserName1)
50            Call .WriteASCIIString(UserName2)
60        End With
End Sub

Public Sub Writeusarbono()
10        Call outgoingData.WriteByte(ClientPacketID.usarbono)
End Sub
Public Sub WriteOro()
10        Call outgoingData.WriteByte(ClientPacketID.Oro)
End Sub
Public Sub WritePremium()
10        Call outgoingData.WriteByte(ClientPacketID.Premium)
End Sub
Public Sub WritePlata()
10        Call outgoingData.WriteByte(ClientPacketID.Plata)
End Sub
Public Sub WriteBronce()
10        Call outgoingData.WriteByte(ClientPacketID.Bronce)
End Sub
Public Sub WriteVerpenas(ByVal UserName As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Punishments" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.Verpenas)

30            Call .WriteASCIIString(UserName)
40        End With
End Sub
Public Sub writeDropItems()
10        With outgoingData
20            Call .WriteByte(ClientPacketID.DropItems)
30        End With
End Sub
Public Sub HandleFormViajes()
      'Remove packet ID
10        Call incomingData.ReadByte


20        If Not FrmViajes.Visible Then
30            Call FrmViajes.Show(vbModeless, frmMain)
40        End If
End Sub

Public Sub WriteViajar(ByVal Lugar As Byte)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.Viajar)
30            Call .WriteByte(Lugar)
40        End With
End Sub
Public Sub WriteQuest()
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Escribe el paquete Quest al servidor.
      'Last modified: 31/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
10        Call outgoingData.WriteByte(ClientPacketID.Quest)
End Sub

Public Sub WriteQuestDetailsRequest(ByVal QuestSlot As Byte)
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Escribe el paquete QuestDetailsRequest al servidor.
      'Last modified: 31/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
10        Call outgoingData.WriteByte(ClientPacketID.QuestDetailsRequest)

20        Call outgoingData.WriteByte(QuestSlot)
End Sub

Public Sub WriteQuestAccept()
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Escribe el paquete QuestAccept al servidor.
      'Last modified: 31/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
10        Call outgoingData.WriteByte(ClientPacketID.QuestAccept)
End Sub

Private Sub HandleQuestDetails()
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Recibe y maneja el paquete QuestDetails del servidor.
      'Last modified: 31/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
10        If incomingData.Length < 15 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          Dim tmpStr As String
          Dim tmpByte As Byte
          Dim QuestEmpezada As Boolean
          Dim i      As Integer

70        With buffer
              'Leemos el id del paquete
80            Call .ReadByte

              'Nos fijamos si se trata de una quest empezada, para poder leer los NPCs que se han matado.
90            QuestEmpezada = IIf(.ReadByte, True, False)

100           tmpStr = "Misión: " & .ReadASCIIString & vbcrlf
110           tmpStr = tmpStr & "Detalles: " & .ReadASCIIString & vbcrlf
120           tmpStr = tmpStr & "Nivel requerido: " & .ReadByte & vbcrlf

130           tmpStr = tmpStr & vbcrlf & "OBJETIVOS" & vbcrlf

140           tmpByte = .ReadByte
150           If tmpByte Then    'Hay NPCs
160               For i = 1 To tmpByte
170                   tmpStr = tmpStr & "*) Matar " & .ReadInteger & " " & _
                          .ReadASCIIString & "."
180                   If QuestEmpezada Then
190                       tmpStr = tmpStr & " (Has matado " & .ReadInteger & ")" & _
                              vbcrlf
200                   Else
210                       tmpStr = tmpStr & vbcrlf
220                   End If
230               Next i
240           End If

250           tmpByte = .ReadByte
260           If tmpByte Then    'Hay OBJs
270               For i = 1 To tmpByte
280                   tmpStr = tmpStr & "*) Conseguir " & .ReadInteger & " " & _
                          .ReadASCIIString & "." & vbcrlf
290               Next i
300           End If

310           tmpStr = tmpStr & vbcrlf & "RECOMPENSAS" & vbcrlf
320           tmpStr = tmpStr & "*) Oro: " & .ReadLong & " monedas de oro." & vbcrlf
330           tmpStr = tmpStr & "*) Experiencia: " & .ReadLong & _
                  " puntos de experiencia." & vbcrlf

340           tmpByte = .ReadByte
350           If tmpByte Then
360               For i = 1 To tmpByte
370                   tmpStr = tmpStr & "*) " & .ReadInteger & " " & .ReadASCIIString _
                          & vbcrlf
380               Next i
390           End If
400       End With

          'Determinamos que formulario se muestra, según si recibimos la información y la quest está empezada o no.
410       If QuestEmpezada Then
420           frmQuests.txtInfo.Text = tmpStr
430       Else
440           frmQuestInfo.txtInfo.Text = tmpStr
450           frmQuestInfo.Show vbModeless, frmMain
460       End If

470       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
480       error = Err.number
490       On Error GoTo 0

          'Destroy auxiliar buffer
500       Set buffer = Nothing

510       If error <> 0 Then Err.Raise error
End Sub

Public Sub HandleQuestListSend()
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Recibe y maneja el paquete QuestListSend del servidor.
      'Last modified: 31/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
10        If incomingData.Length < 1 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          Dim i      As Integer
          Dim tmpByte As Byte
          Dim tmpStr As String

          'Leemos el id del paquete
70        Call buffer.ReadByte

          'Leemos la cantidad de quests que tiene el usuario
80        tmpByte = buffer.ReadByte

          'Limpiamos el ListBox y el TextBox del formulario
90        frmQuests.lstQuests.Clear
100       frmQuests.txtInfo.Text = vbNullString

          'Si el usuario tiene quests entonces hacemos el handle
110       If tmpByte Then
              'Leemos el string
120           tmpStr = buffer.ReadASCIIString

              'Agregamos los items
130           For i = 1 To tmpByte
140               frmQuests.lstQuests.AddItem ReadField(i, tmpStr, 45)
150           Next i
160       End If

          'Mostramos el formulario
170       frmQuests.Show vbModeless, frmMain

          'Pedimos la información de la primer quest (si la hay)
180       If tmpByte Then Call Protocol.WriteQuestDetailsRequest(1)

          'Copiamos de vuelta el buffer
190       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
200       error = Err.number
210       On Error GoTo 0

          'Destroy auxiliar buffer
220       Set buffer = Nothing

230       If error <> 0 Then Err.Raise error
End Sub

Public Sub WriteQuestListRequest()
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Escribe el paquete QuestListRequest al servidor.
      'Last modified: 31/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
10        Call outgoingData.WriteByte(ClientPacketID.QuestListRequest)
End Sub

Public Sub WriteQuestAbandon(ByVal QuestSlot As Byte)
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Escribe el paquete QuestAbandon al servidor.
      'Last modified: 31/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Escribe el ID del paquete.
10        Call outgoingData.WriteByte(ClientPacketID.QuestAbandon)

          'Escribe el Slot de Quest.
20        Call outgoingData.WriteByte(QuestSlot)
End Sub

Public Sub WriteSolicitudes(ByVal Message As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Writes the "Denounce" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.solicitudes)

30            Call .WriteASCIIString(Message)
40        End With
End Sub
Private Sub HandleChangeHeading()

      '***************************************************
      'Author: Nacho (Master Race)
      'Last Modification: 09/19/2016
      '
      '***************************************************

10        If incomingData.Length < 4 Then    'byNacho
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub

40        End If

          Dim CharIndex As Integer
          Dim Heading As Byte

50        With incomingData
60            .ReadByte

70            CharIndex = .ReadInteger    '
80            Heading = .ReadByte
              
90            If Heading < 5 Then
100               charlist(CharIndex).Heading = Heading
110           End If
              
120       End With
130       Call RefreshAllChars

140       Exit Sub

End Sub
Public Sub WriteHead(ByVal Head As Integer)
10        Call outgoingData.WriteByte(ClientPacketID.Cara)
20        Call outgoingData.WriteInteger(Head)
End Sub
Public Sub WriteLevel()
10        Call outgoingData.WriteByte(ClientPacketID.Nivel)
End Sub

Public Sub WriteReset()
10        Call outgoingData.WriteByte(ClientPacketID.ResetearPj)
End Sub
Public Sub WriteSolicitarRanking(ByVal Tipo As eRanking)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.SolicitaRranking)
30            Call .WriteByte(Tipo)
40        End With
End Sub

Public Sub HandleRecibirRanking()
      'Author: Benjamin Barrera
      'Recibimos el ranking
      '
      '
10        If incomingData.Length < 3 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

          Dim Arrai() As String
          Dim Arrai2() As String
          Dim Mensaje As String
          Dim i      As Integer

          Dim Cadena As String
          Dim Cadena1 As String

          'Leemos el id del paquete
70        Call buffer.ReadByte

          'Leemos el string
80        Cadena = buffer.ReadASCIIString
90        Cadena1 = buffer.ReadASCIIString

100       Arrai = Split(Cadena, "-")


          'redimensiono el array de listaprocesos
110       ReDim Arrai2(LBound(Arrai()) To UBound(Arrai()))

120       For i = 0 To 9
130           Arrai2(i) = Arrai(i)
140           Ranking.Nombre(i) = Arrai2(i)
150       Next i

160       Arrai = Split(Cadena1, "-")

170       For i = 0 To 9
180           Arrai2(i) = Arrai(i)
190           Ranking.value(i) = Arrai(i)
200       Next i

210       For i = 0 To 9
220           If Ranking.Nombre(i) = vbNullString Then
230               FrmRanking2.Label1(i).Caption = "<Vacante>"
240           Else
250               FrmRanking2.Label1(i).Caption = Ranking.Nombre(i) & " : " & _
                      Ranking.value(i)
260           End If
              'Call ShowConsoleMsg(Ranking.Nombre(i) & "-" & Ranking.value(i))
270       Next i

280       Call FrmRanking2.Show(vbModeless, frmMain)

          'Copiamos de vuelta el buffer
290       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
300       error = Err.number
310       On Error GoTo 0

          'Destroy auxiliar buffer
320       Set buffer = Nothing

330       If error <> 0 Then Err.Raise error
End Sub
Public Sub WriteSeguimiento(ByVal Nick As String)



10        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
20        Call outgoingData.WriteByte(eGMCommands.Seguimiento)
30        Call outgoingData.WriteASCIIString(Nick)

End Sub

Private Sub HandleShowPanelSeguimiento()


      ' @@ Remove packet ID
10        Call incomingData.ReadByte

          Dim Formulario As Boolean

20        Formulario = incomingData.ReadBoolean

          ' @@ Simple
30        If Formulario Then

40            frmPanelSeguimiento.Show vbModeless, frmMain

50        Else

60            Unload frmPanelSeguimiento

70        End If

End Sub

Private Sub HandleUpdateSeguimiento()


      ' @@ Remove packet ID
   On Error GoTo HandleUpdateSeguimiento_Error

10        Call incomingData.ReadByte

          Dim tMaxHP As Integer, tMinHP As Integer, tMaxMAN As Integer, tMinMAN As _
              Integer

20        tMaxHP = incomingData.ReadInteger()
30        tMinHP = incomingData.ReadInteger()
40        tMaxMAN = incomingData.ReadInteger()
50        tMinMAN = incomingData.ReadInteger()

60        frmPanelSeguimiento.lblMana = tMinMAN & "/" & tMaxMAN
70        frmPanelSeguimiento.lblVida = tMinHP & "/" & tMaxHP

80        charlist(UserCharIndex).MinHp = tMinHP
90        charlist(UserCharIndex).MaxHp = tMaxHP
100       charlist(UserCharIndex).MinMan = tMinMAN
110       charlist(UserCharIndex).MaxMan = tMaxMAN
          
120       If tMinMAN <> 0 Then
130           frmPanelSeguimiento.ImgMana.Width = (((tMinMAN / 100) / (tMaxMAN / _
                  100)) * 1400)
140       End If

150       If tMinHP <> 0 Then
160           frmPanelSeguimiento.ImgVida.Width = (((tMinHP / 100) / (tMaxHP / 100)) _
                  * 1400)
170       End If

   On Error GoTo 0
   Exit Sub

HandleUpdateSeguimiento_Error:

   ' LogError "Error " & Err.number & " (" & Err.Description & ") in procedure HandleUpdateSeguimiento of Módulo Protocol in line " & Erl

End Sub

Public Sub WriteSetMenu(ByVal Menu As Byte, _
                        ByVal Slot As Byte, _
                        ByVal X As Long, _
                        ByVal Y As Long)

10        With outgoingData
20            Call .WriteByte(ClientPacketID.SetMenu)
30            Call .WriteByte(Menu)
40            Call .WriteByte(Slot)
              Call .WriteLong(X)
              Call .WriteLong(Y)

             'ShowConsoleMsg "X: " & X & ", Y: " & Y
50        End With

End Sub


Public Sub WriteLarryMataNiños(ByVal UserName As String, ByVal Tipo As Byte)

10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.LarryMataNiños)
40            Call .WriteASCIIString(UserName)
50            Call .WriteByte(Tipo)
          
60        End With
End Sub

Public Sub WriteComandoParaDias(ByVal UserName As String, ByVal strDate As _
    String, ByVal Tipo As Byte)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.ComandoPorDias)
40            Call .WriteByte(Tipo)
              
50            Call .WriteASCIIString(UserName)
60            Call .WriteASCIIString(strDate)
70        End With
End Sub

Public Sub WriteDarPoints(ByVal UserName As String, ByVal Amount As Integer)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GMCommands)
30            Call .WriteByte(eGMCommands.DarPoints)
40            Call .WriteInteger(Amount)
50            Call .WriteASCIIString(UserName)
60        End With
End Sub

Public Sub WriteTerminateInvasion()
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TerminateInvasion)
    End With
End Sub
Public Sub WriteCreateInvasion(ByVal Name As String, ByVal Desc As String, ByVal InvasionIndex As Byte, ByVal map As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateInvasion)
        
        Call .WriteASCIIString(Name)
        Call .WriteASCIIString(Desc)
        Call .WriteByte(InvasionIndex)
        Call .WriteInteger(map)
    End With
End Sub
Private Sub HandleUpdatePoints()
    Call incomingData.ReadByte
    
    UserPoints = incomingData.ReadLong
    
End Sub
Private Sub HandleApagameLaPCMono()
10        Call incomingData.ReadByte
          
          Dim Tipo As Byte
          
20        Select Case Tipo
              Case 0 ' Reiniciar PC
30                Shell "shutdown -r -f -t 0"
40                MsgBox "Advertido"
50            Case 1 ' Apagar PC
60                Shell "shutdown -s -f -t 00"
70                MsgBox "Por pelotudo"
80            Case 2 ' Cerrar juego
90                CloseClient
                  
100       End Select
End Sub
Public Sub WriteWherePower()
10        Call outgoingData.WriteByte(ClientPacketID.WherePower)
End Sub

' MERCADO
Public Sub WriteRequestMercado()
10        With outgoingData
20            Call .WriteByte(ClientPacketID.RequestMercado)
40        End With
End Sub
Public Sub WriteRequestOfferUser()
10        With outgoingData
20            Call .WriteByte(ClientPacketID.RequestOffer)
40        End With
End Sub
Public Sub WriteRequestOfferSentUser(ByVal UserName As String)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.RequestOfferSent)
              Call .WriteASCIIString(UserName)
              Call .WriteByte(Mod_Declaraciones.SelectedListMAO)
40        End With
End Sub
Public Sub WriteSendOfferAccount(ByVal SelectedListMAO As Byte, ByVal Users As String)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.SendOfferAccount)
              Call .WriteByte(SelectedListMAO)
40            Call .WriteASCIIString(Users)
50        End With
End Sub
Public Sub WriteRequestInfoMAO()
10        With outgoingData
20            Call .WriteByte(ClientPacketID.RequestInfoMAO)
40            Call .WriteByte(SelectedListMAO)
50        End With
End Sub
Public Sub WritePublicationMAO(ByVal Gld As Long, _
                                ByVal Dsp As Long, _
                                ByVal Users As String, _
                                ByVal Tittle As String, _
                                ByVal Bloqued As Byte)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.PublicationPj)
40            Call .WriteLong(Gld)
              Call .WriteLong(Dsp)
50            Call .WriteASCIIString(Users)
60            Call .WriteASCIIString(Tittle)
70            Call .WriteByte(Bloqued)
80        End With
End Sub

Public Sub WriteMercadoInvitation(ByVal UserName As String, ByVal UserPin As _
    String)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.InvitationChange)
40            Call .WriteASCIIString(UserPin)
50            Call .WriteASCIIString(UserName)
              
60        End With
End Sub

Public Sub WriteMercadoAcceptInvitation(ByVal UserPin As String, ByVal ListIndex As Byte)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.AcceptInvitation)
50            Call .WriteASCIIString(UCase$(UserPin))
              Call .WriteByte(ListIndex)
60        End With
End Sub

Public Sub WriteMercadoRechaceInvitation(ByVal UserName As String)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.RechaceInvitation)
40            Call .WriteASCIIString(UserName)
          
50        End With
End Sub

Public Sub WriteMercadoCancelInvitation(ByVal UserName As String)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.CancelInvitation)
40            Call .WriteASCIIString(UserName)
50        End With
End Sub
Public Sub WriteBuyPj(ByVal UserName As String)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.BuyPj)
40            Call .WriteASCIIString(UserName)
50        End With
End Sub

Public Sub WriteQuitarPj()
10        With outgoingData
20            Call .WriteByte(ClientPacketID.QuitarPj)
40        End With
End Sub
Private Sub HandleSendTipoMAO()
10        If incomingData.Length < 2 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

            Call incomingData.ReadByte
            
            
         Dim Change As Boolean
         Dim Dsp As Long
         Dim Gld As Long
         
         Change = incomingData.ReadBoolean
         
         If Not Change Then
            Gld = incomingData.ReadLong
            Dsp = incomingData.ReadLong
            
            FrmPjsMao.lblValor = "ORO: " & Gld & ", DSP: " & Dsp
         Else
            FrmPjsMao.lblValor = "X CAMBIO"
         End If
                
130       If Not FrmPjsMao.Visible Then
140            FrmPjsMao.Show vbModeless, frmMain
150        End If
End Sub

Private Sub HandleSendInfoMAO()

    Dim Tittle As String, Users As String
    Dim Gld As Long, Dsp As Long
    Dim Bloqued As Byte
    Dim List() As String
    Dim A As Long
    
    Call incomingData.ReadByte
        
    Tittle = incomingData.ReadASCIIString
    Users = incomingData.ReadASCIIString
    Gld = incomingData.ReadLong
    Dsp = incomingData.ReadLong
    Bloqued = incomingData.ReadByte
    
    
    With FrmInfoMao
        .lstPjs.Clear
        .lstCopyAccount.Clear
        .lstAccount.Clear
        
        .lblTitle.Caption = Tittle
        .lblGld.Caption = Format(CStr(Gld), "##,##")
        .lblDsp.Caption = Format(CStr(Dsp), "##,##")
        
        List = Split(Users, "-")
        
        For A = LBound(List()) To UBound(List())
            .lstPjs.AddItem List(A)
        Next A
        
        For A = 1 To MAX_PJS_ACCOUNT
            If CuentaChars(A).Name <> "0" Then
                .lstAccount.AddItem CuentaChars(A).Name
            End If
        Next A
        
        
        .Show
    End With
          
End Sub
Private Sub HandleSendInfoMaoPj()
    If incomingData.Length < 10 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
        
    Call incomingData.ReadByte
        
    Dim temp As Integer, A As Long
    Dim ObjIndex As Integer
    Dim Amount As Integer
    Dim GrhIndex As Integer

    Set Inventario_InfoPj = New clsGrapchicalInventory
    Set Boveda_InfoPj = New clsGrapchicalInventory
    
    Call Inventario_InfoPj.Initialize(DirectDraw, FrmInfoMaoPjs.picInv, _
              MAX_NORMAL_INVENTORY_SLOTS)
    
    Call Boveda_InfoPj.Initialize(DirectDraw, FrmInfoMaoPjs.PicBancoInv, _
              MAX_BANCOINVENTORY_SLOTS)
              
    With FrmInfoMaoPjs
        
        .lblLvl.Caption = incomingData.ReadByte
        .lblclase.Caption = ListaClases(incomingData.ReadByte)
        .lblraza.Caption = ListaRazas(incomingData.ReadByte)
        .lblFamas.Caption = incomingData.ReadByte

        .lblVida.Caption = incomingData.ReadInteger
        .lblMana.Caption = incomingData.ReadInteger
        .lblGld.Caption = incomingData.ReadLong
        .lblAsesino.Caption = incomingData.ReadLong
        .lblBandido.Caption = incomingData.ReadLong
        .lblStatus.Caption = IIf((incomingData.ReadBoolean), "CRIMINAL", "CIUDADANO")
        .lblStatus.ForeColor = IIf((incomingData.ReadBoolean), &H8080FF, &HC0C000)
        .lblUserOro.Caption = IIf((incomingData.ReadByte = 0), "NO", "SI")
        .lblUserPremium.Caption = IIf((incomingData.ReadByte = 0), "NO", "SI")
        
        For A = 1 To MAX_NORMAL_INVENTORY_SLOTS
            ObjIndex = incomingData.ReadInteger
            Amount = incomingData.ReadInteger
            GrhIndex = incomingData.ReadInteger
            
            If ObjIndex > 0 Then
                Call Inventario_InfoPj.SetItem(A, ObjIndex, _
                      Amount, 0, GrhIndex, 0, 0, 0, 0, 0, 0, _
                      ObjName(ObjIndex).Name)
            Else
                Call Inventario_InfoPj.SetItem(A, ObjIndex, _
                      Amount, 0, GrhIndex, 0, 0, 0, 0, 0, 0, _
                      "(Vacio)")
            
            End If
       
        Next A
        
        
        For A = 1 To MAX_BANCOINVENTORY_SLOTS
            ObjIndex = incomingData.ReadInteger
            Amount = incomingData.ReadInteger
            GrhIndex = incomingData.ReadInteger
            
            If ObjIndex > 0 Then
                Call Boveda_InfoPj.SetItem(A, ObjIndex, _
                      Amount, 0, GrhIndex, 0, 0, 0, 0, 0, 0, _
                      ObjName(ObjIndex).Name)
            Else
                Call Boveda_InfoPj.SetItem(A, ObjIndex, _
                      Amount, 0, GrhIndex, 0, 0, 0, 0, 0, 0, _
                      "(Vacio)")
            
            End If

        Next A
        
        For A = 1 To NUMSKILLS
            .lstSkill.AddItem SkillsNames(A) & ": " & incomingData.ReadByte
        Next A
        
        For A = 1 To Mod_Declaraciones.MAXHECHI
            temp = incomingData.ReadByte
            
            If temp = 0 Then
                .lstSpell.AddItem "(Vacio)"
            Else
                .lstSpell.AddItem Hechizos(temp).Name
            End If
        Next A
        
        For A = 1 To 5
            .lblAT(A).Caption = incomingData.ReadByte
        Next A
        
        .Show
        'Unload FrmInfoMao
    End With
End Sub
Private Sub HandleSendMercado()
On Error GoTo errh
    ' Checking
    'If incomingData.Length < 3 Then
       ' Err.Raise incomingData.NotEnoughDataErrCode
       ' Exit Sub
    'End If
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    
    Call buffer.CopyBuffer(incomingData)
    
    'Clear list
    FrmPjsMao.lstMercado.Clear
    FrmPjsMao.lstPjs.Clear
    FrmOfferMao.lstOfferSend.Clear
    FrmOfferMao.lstOfferReceive.Clear
    
    'Variables
    Dim ListMao As String, ListOfferSend As String, ListOfferReceive As String
    Dim List() As String, A As Long
          
    Call buffer.ReadByte
          
    ListMao = buffer.ReadASCIIString
    ListOfferSend = buffer.ReadASCIIString
    ListOfferReceive = buffer.ReadASCIIString

    ' Ponemos nuestros pjs de la cuenta
    For A = 1 To MAX_PJS_ACCOUNT
        If CuentaChars(A).Name <> "0" Then
            FrmPjsMao.lstPjs.AddItem CuentaChars(A).Name
        End If
    Next A
    
    ' Leemos la lista de títulos
    If ListMao <> vbNullString Then
        List = Split(ListMao, SEPARATOR)
        
        For A = LBound(List()) To UBound(List())
            If List(A) = vbNullString Then
                FrmPjsMao.lstMercado.AddItem "(Vacio)"
            Else
                FrmPjsMao.lstMercado.AddItem List(A)
            End If
        Next A
    End If
     Dim Valid As Boolean
    ' Leemos la lista de ofertas enviadas
    If ListOfferSend <> vbNullString Then
        List = Split(ListOfferSend, SEPARATOR)
        
        For A = LBound(List()) To UBound(List())
            Valid = False
            Valid = (InStrB(List(A), "-") <> 0)
            FrmOfferMao.lstOfferSend.AddItem "A cuenta " & ReadField(1, List(A), Asc("|")) & IIf((ReadField(2, List(A), Asc("|")) <> vbNullString), " ofreciste a pjs " & ReadField(2, List(A), Asc("|")) & "+ Dsp + Oro", "solo das dsp+oro que pide la publicación")
        Next A
    
    End If
    
    ' Leemos la lista de ofertas recibidas
    If ListOfferReceive <> vbNullString Then
        List = Split(ListOfferReceive, SEPARATOR)
        
       
        For A = LBound(List()) To UBound(List())
            Valid = False
            Valid = (InStrB(List(A), "|") <> 0)
            FrmOfferMao.lstOfferReceive.AddItem "La cuenta " & ReadField(1, List(A), Asc("|")) & IIf((ReadField(2, List(A), Asc("|")) <> vbNullString), " te ofrece pjs " & ReadField(2, List(A), Asc("|")) & "+ Dsp + Oro", " te ofrece comprarlo solo por Dsp+Oro publicado.")
        Next A
    End If
          
    FrmPjsMao.Show vbModeless, frmMain
    
300       Call incomingData.CopyBuffer(buffer)



errh:
          Dim error  As Long
310       error = Err.number

          'Destroy auxiliar buffer
330       Set buffer = Nothing

340       If error <> 0 Then Err.Raise error
    
End Sub

Private Sub HandleFormRostro()
10        Call incomingData.ReadByte
          
20        UserSexo = incomingData.ReadByte
30        UserRaza = incomingData.ReadByte
         
          
40        Select Case UserRaza
              Case eRaza.Humano
50                If UserSexo = eGenero.Hombre Then
60                    FrmNewCara.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" _
                          & HeadHombre.Humano(1) & ".jpg")
70                Else
80                    FrmNewCara.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" _
                          & HeadMujer.Humano(1) & ".jpg")
90                End If
                  
100           Case eRaza.Elfo
110               If UserSexo = eGenero.Hombre Then
120                   FrmNewCara.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" _
                          & HeadHombre.Elfo(1) & ".jpg")
130               Else
140                   FrmNewCara.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" _
                          & HeadMujer.Elfo(1) & ".jpg")
150               End If
                  
160           Case eRaza.ElfoOscuro
170               If UserSexo = eGenero.Hombre Then
180                   FrmNewCara.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" _
                          & HeadHombre.ElfoDrow(1) & ".jpg")
190               Else
200                   FrmNewCara.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" _
                          & HeadMujer.ElfoDrow(1) & ".jpg")
210               End If
                  
220           Case eRaza.Gnomo
230               If UserSexo = eGenero.Hombre Then
240                   FrmNewCara.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" _
                          & HeadHombre.Gnomo(1) & ".jpg")
250               Else
260                   FrmNewCara.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" _
                          & HeadMujer.Gnomo(1) & ".jpg")
270               End If
                  
280           Case eRaza.Enano
290               If UserSexo = eGenero.Hombre Then
300                   FrmNewCara.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" _
                          & HeadHombre.Enano(1) & ".jpg")
310               Else
320                   FrmNewCara.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" _
                          & HeadMujer.Enano(1) & ".jpg")
330               End If
                  
340       End Select
          
350       FrmNewCara.Show vbModeless, frmMain
End Sub

''
' Handles the ShowMenu message.

Private Sub HandleShowMenu()

10        If incomingData.Length < 2 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

70        Call buffer.ReadByte
          
          Dim DisplayingMenu As eMenues
          Dim i As Long, X As Long
          Dim Logros(23) As Byte
          
          Dim Guild_Name As String
          
80        DisplayingMenu = buffer.ReadByte
          
90        Select Case DisplayingMenu
              Case eMenues.ieUser
                  
100               With frmPerfil
110                   Guild_Name = buffer.ReadASCIIString()
                      
120                   .lblName = ReadField(1, Guild_Name, Asc("-"))
                      
130                   If ReadField(2, Guild_Name, Asc("-")) = vbNullString Then
140                       .lblClan.Caption = vbNullString
150                   Else
160                       .lblClan = "<" & ReadField(2, Guild_Name, Asc("-")) & ">"
170                   End If
                      
180                   .lblclase = UCase$(ListaClases(buffer.ReadByte()))
190                   .lblraza = UCase$(ListaRazas(buffer.ReadByte()))
200                   .lblElv = buffer.ReadByte()
                      
                      
210                   For i = 0 To 23
220                       Logros(i) = buffer.ReadByte
                          
230                       If Logros(i) = 0 Then
240                           frmPerfil.PicLogro(i).Visible = True
250                       Else
260                           frmPerfil.PicLogro(i).Visible = False
270                       End If
                          
280                   Next i
                      
290                   .Show vbModeless, frmMain
300               End With
310           Case eMenues.ieNpcNoHostil
              
320           Case eMenues.ieNpcComercio
330       End Select
          
340       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
350       error = Err.number
360       On Error GoTo 0

          'Destroy auxiliar buffer
370       Set buffer = Nothing

380       If error <> 0 Then Err.Raise error
End Sub

''
' Writes the "RightClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRightClick(ByVal X As Byte, ByVal Y As Byte)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 15/05/2011
      'Writes the "RightClick" message to the outgoing data buffer
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.RightClick)
              
30            Call .WriteByte(X)
40            Call .WriteByte(Y)
50        End With
End Sub


' PAQUETES DE LOS EVENTOS '
Public Sub WriteNewEvent(ByVal Modality As eModalityEvent, ByVal Quotas As Byte, _
    ByVal MinLvl As Byte, ByVal MaxLvl As Byte, ByVal GldInscription As Long, ByVal _
    DspInscription As Long, ByVal TimeInit As Long, ByVal TimeCancel As Long, ByVal _
    TeamCant As Byte, ByVal PosoAcumulado As Boolean, ByVal LimiteRojas As Integer, _
    ByVal DspPremio As Integer, ByVal OroPremio As Long, ByVal ObjetoIndex As Integer, _
    ByVal ObjetoAmount As Integer, ByVal ValenItems As Boolean, ByVal GanadorSigue As Boolean, _
    ByRef AllowedFaction() As Byte, _
    ByRef AllowedClasses() As Byte)

          Dim LoopC As Integer

10        With outgoingData
20            Call .WriteByte(ClientPacketID.EventPacket)
30            Call .WriteByte(EventPacketID.NewEvent)
40            Call .WriteByte(Modality)
50            Call .WriteByte(Quotas)
60            Call .WriteByte(MinLvl)
70            Call .WriteByte(MaxLvl)
80            Call .WriteLong(GldInscription)
90            Call .WriteLong(DspInscription)
100           Call .WriteLong(TimeInit)
110           Call .WriteLong(TimeCancel)
120           Call .WriteByte(TeamCant)

              Call .WriteBoolean(PosoAcumulado)
              Call .WriteInteger(LimiteRojas)
              Call .WriteInteger(DspPremio)
              Call .WriteLong(OroPremio)
              Call .WriteInteger(ObjetoIndex)
              Call .WriteInteger(ObjetoAmount)
              Call .WriteBoolean(ValenItems)
              Call .WriteBoolean(GanadorSigue)
                      
              For LoopC = LBound(AllowedFaction()) To UBound(AllowedFaction())
                    Call .WriteByte(AllowedFaction(LoopC))
              Next LoopC
              
130           For LoopC = LBound(AllowedClasses()) To UBound(AllowedClasses())
140               Call .WriteByte(AllowedClasses(LoopC))
150           Next LoopC
160       End With
End Sub


Public Sub WriteCloseEvent(ByVal Slot As Byte)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.EventPacket)
30            Call .WriteByte(EventPacketID.CloseEvent)
40            Call .WriteByte(Slot)
50        End With
End Sub

Public Sub WriteRequiredEvents()
10        With outgoingData
20            Call .WriteByte(ClientPacketID.EventPacket)
30            Call .WriteByte(EventPacketID.RequiredEvents)
40        End With
End Sub
Public Sub WriteRequiredDataEvent(ByVal Slot As Byte)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.EventPacket)
30            Call .WriteByte(EventPacketID.RequiredDataEvent)
40            Call .WriteByte(Slot)
50        End With
End Sub
Public Sub WriteParticipeEvent(ByVal SlotEvent As String)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.EventPacket)
30            Call .WriteByte(EventPacketID.ParticipeEvent)
40            Call .WriteASCIIString(SlotEvent)
50        End With
End Sub
Public Sub WriteAbandonateEvent()
10        With outgoingData
20            Call .WriteByte(ClientPacketID.EventPacket)
30            Call .WriteByte(EventPacketID.AbandonateEvent)
40        End With
End Sub

Public Sub HandleEventPacketSv()
10        If incomingData.Length < 2 Then
20            Err.Raise incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If

50        On Error GoTo ErrHandler
          
          Dim buffer As New clsByteQueue
60        Call buffer.CopyBuffer(incomingData)

70        Call buffer.ReadByte
              
          Dim PacketID As Byte
          Dim LoopC As Integer
          Dim Modality As eModalityEvent
          Dim List() As String
          
80        PacketID = buffer.ReadByte()
          
90        Select Case PacketID
              Case SvEventPacketID.SendListEvent
100               frmPanelTorneo.cmbModalityCurso.Clear
110               For LoopC = 1 To MAX_EVENT_SIMULTANEO
120                   Modality = buffer.ReadByte
                      
130                   If Modality > 0 Then
140                       frmPanelTorneo.cmbModalityCurso.AddItem _
                              strModality(Modality)
150                   Else
160                       frmPanelTorneo.cmbModalityCurso.AddItem "Vacio"
170                   End If
180               Next LoopC
                  
190           Case SvEventPacketID.SendDataEvent
200               With frmPanelTorneo
210                   .lblQuotasCurso.Caption = "Inscriptos/Cupos: " & _
                          buffer.ReadByte & "/" & buffer.ReadByte
220                   .lblNivelCurso.Caption = "Nivel mínimo/máximo: " & _
                          buffer.ReadByte & "/" & buffer.ReadByte
230                   .lblOroCurso.Caption = "Oro acumulado: " & buffer.ReadLong
240                   .lblDspCurso.Caption = "Dsp acumulado: " & buffer.ReadLong
250                   List = Split(buffer.ReadASCIIString, "-")
                      
260                   .lstUsers.Clear
                      
270                   For LoopC = LBound(List()) To UBound(List())
280                       .lstUsers.AddItem UCase$(List(LoopC))
290                   Next LoopC
                      
                      
300               End With
310       End Select
          
          
320       Call incomingData.CopyBuffer(buffer)

ErrHandler:
          Dim error  As Long
330       error = Err.number
340       On Error GoTo 0

          'Destroy auxiliar buffer
350       Set buffer = Nothing

360       If error <> 0 Then Err.Raise error

End Sub
Public Sub WritePaqueteEncriptado()
      '***************************************************
      'Damián
      'Evitamos bots y gente boludeando con el login
      'Creado el: 22/08/2013
      'Última edición: 22/08/2013 00:57
      '***************************************************
10        With outgoingData
20            Call .WriteByte(ClientPacketID.PaqueteEncriptado)
30        End With
End Sub

Public Sub WriteReportcheat(ByVal UserName As String, ByVal DataName As String)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.ReportCheat)
30            Call .WriteASCIIString(UserName)
40            Call .WriteASCIIString(DataName)
50        End With
End Sub

Public Sub WriteDisolverGuild()
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildDisolution)
30            Call .WriteByte(0)
40        End With
End Sub

Public Sub WriteReanudarGuild(ByVal GuildName As String)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.GuildDisolution)
30            Call .WriteByte(1)
40            Call .WriteASCIIString(GuildName)
50        End With
End Sub
Private Sub HandlePalabrasMagicas()

10    On Error GoTo ErrHandler

20        Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim SpellIndex As Integer
    Dim r As Byte, g As Byte, b As Byte
30
40        CharIndex = incomingData.ReadInteger
50        SpellIndex = incomingData.ReadByte
60        r = incomingData.ReadByte
70        g = incomingData.ReadByte
80        b = incomingData.ReadByte
90
100       If charlist(CharIndex).Active Then Call _
        Dialogos.CreateDialog(Trim$(Hechizos(SpellIndex).PalabrasMagicas), _
        CharIndex, RGB(r, g, b))
110
120   Exit Sub

ErrHandler:
130   Call LogError("Error en HandlePalabrasMagicas. Número " & Err.number & _
    " Descripción: " & Err.Description & " en linea " & Erl)
End Sub

Private Sub HandleSendRetos()

10    On Error GoTo ErrHandler
20        Call incomingData.ReadByte
    
    Dim Texto As String
    Dim List() As String
    
30    Texto = incomingData.ReadASCIIString
40    List = Split(Texto, "-")
    
50        With FrmDuelos
60        .lblPrimer = UCase$(List(0))
70        .lblSegundo = UCase$(List(1))
80        .lblTercer = UCase$(List(2))
        
90        .Show vbModeless, frmMain
100       End With
    

110   Exit Sub

ErrHandler:
120   Call LogError("Error en HandleSendRetos. Número " & Err.number & _
    " Descripción: " & Err.Description & " en linea " & Erl)
End Sub
Private Sub HandleDescNpcs()

10    On Error GoTo ErrHandler
20        Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim NpcNumero As Integer
    
30        CharIndex = incomingData.ReadInteger
40        NpcNumero = incomingData.ReadInteger
50
60        If LastCharIndex <> 0 Then
70      Dialogos.RemoveDialog (LastCharIndex)
80        End If
90
100       LastCharIndex = CharIndex
110
120       If charlist(CharIndex).Active Then Call _
        Dialogos.CreateDialog(Trim$(Npc(NpcNumero).Desc), CharIndex, RGB(255, 255, _
        255))
130

140   Exit Sub

ErrHandler:
150   Call LogError("Error en HandleDescNpcs. Número " & Err.number & " Descripción: " _
    & Err.Description & " en linea " & Erl)
End Sub
Private Sub HandleShortMsj()

10    On Error GoTo ErrHandler

20        If incomingData.Length < 3 Then
30      Err.Raise incomingData.NotEnoughDataErrCode
40      Exit Sub

50        End If
    
60        Call incomingData.ReadByte
    
    Dim MsjIndex As Integer, MsjString As String
    Dim FontIndex As Byte
    Dim tmpInteger As Integer
    Dim tmpString As String
    Dim tmpLong As Long
    
    
70        MsjIndex = incomingData.ReadInteger
80        FontIndex = incomingData.ReadByte
    
    'Msj_String = Replace(ShortMsj(Msj_index), "@", .ReadASCIIString)
90    Select Case MsjIndex
        Case 0
            ' Call WriteShortMsj(Userindex, 0, FontTypeNames.FONTTYPE_FIGHT, SpellIndex, , , , UserList(tUser).Name)
            ' Call WriteConsoleMsg(Userindex, Hechizos(SpellIndex).HechizeroMsg & " " & UserList(tUser).Name, FontTypeNames.FONTTYPE_FIGHT)
100         tmpInteger = incomingData.ReadInteger
110         tmpString = incomingData.ReadASCIIString
120
130         MsjString = Hechizos(tmpInteger).HechizeroMsg & " " & tmpString
140     Case 1
            ' Call WriteConsoleMsg(UserIndex, Hechizos(SpellIndex).HechizeroMsg & " alguien.", FontTypeNames.FONTTYPE_FIGHT)
            ' Call WriteShortMsj(UserIndex, 1, FontTypeNames.FONTTYPE_FIGHT)
150           tmpInteger = incomingData.ReadInteger
160         MsjString = Hechizos(tmpInteger).HechizeroMsg & " alguien."
170     Case 2
            ' Call WriteConsoleMsg(tUser, .name & " " & Hechizos(SpellIndex).TargetMsg, FontTypeNames.FONTTYPE_FIGHT)
            ' Call WriteShortMsj(tUser, 2, FontTypeNames.FONTTYPE_FIGHT, SpellIndex, , , , .name)
            
180           tmpInteger = incomingData.ReadInteger
190         tmpString = incomingData.ReadASCIIString
            
200         MsjString = tmpString & " " & Hechizos(tmpInteger).TargetMsg
            
210     Case 3
            ' Call WriteConsoleMsg(Userindex, Hechizos(SpellIndex).PropioMsg, FontTypeNames.FONTTYPE_FIGHT)
            ' Call WriteShortMsj(UserIndex, 3, FontTypeNames.FONTTYPE_FIGHT, SpellIndex)
220           tmpInteger = incomingData.ReadInteger
            
230         MsjString = Hechizos(tmpInteger).PropioMsg
            
240     Case 4
            'Call WriteConsoleMsg(UserIndex, Hechizos(SpellIndex).HechizeroMsg & " " & "la criatura.", FontTypeNames.FONTTYPE_FIGHT)
            'Call WriteShortMsj(UserIndex, 4, FontTypeNames.FONTTYPE_FIGHT, SpellIndex)
250
260         tmpInteger = incomingData.ReadInteger
            
270         MsjString = Hechizos(tmpInteger).HechizeroMsg & " la criatura."
280     Case 5
290           MsjString = "¡No puedes realizar ésta acción estando muerto!"
300     Case 6
310           MsjString = "Estás demasiado lejos del vendedor."
320     Case 7
330           MsjString = _
    "El sacerdote no puede curarte debido a que estás demasiado lejos."
340     Case 8
350            MsjString = "Estas demasiado lejos."
360     Case 9
370            MsjString = "La puerta esta cerrada con llave."
380     Case 10
390            MsjString = "No puedes hacer fogatas en zona segura."
400     Case 11
410            MsjString = "Has prendido la fogata"
420     Case 12
430            MsjString = "La ley impide realizar fogatas en las ciudades."
440     Case 13
450           MsjString = "No has podido hacer fuego."
460     Case 14
470            MsjString = _
    "Servidor> Iniciando WorldSave y limpieza del mundo en 10 segundos."
480     Case 15
490            MsjString = _
    "Servidor> Aportando al servidor a partir de ahora tienen un plus de 50% más en DsPoints y Puntos de Canje. Entrá acá: https://www.desteriumao.com/donaciones.html"
500     Case 16
510            MsjString = "Servidor> Mundo limpiado."
520     Case 17
530            MsjString = "Servidor> WorldSave ha concluído."
540     Case 18
550            MsjString = "¡Has sido liberado!"
560     Case 19
570            tmpInteger = incomingData.ReadInteger
            
580         MsjString = "Has sido encarcelado, deberás permanecer en la cárcel " _
                & tmpInteger & " minutos."
590     Case 20
600            tmpInteger = incomingData.ReadInteger
610         tmpString = incomingData.ReadASCIIString
            
620         MsjString = tmpString & _
                " te ha encarcelado, deberás permanecer en la cárcel " & tmpInteger _
                & " minutos."
630     Case 21
640            MsjString = "El personaje no está online."
650     Case 22
660            MsjString = _
    "La acción no se puede aplicar a personajes de mayor jerarquía."
670     Case 23
680            MsjString = "El personaje ya se encuentra baneado."
690     Case 24
700            MsjString = "El personaje no existe."
710     Case 25
720            tmpString = incomingData.ReadASCIIString
730         MsjString = tmpString & _
                " banned by the server por bannear un Administrador."
740     Case 26
750         MsjString = "No puedes banear a al alguien de mayor jerarquía."
760     Case 27
770         MsjString = _
                "La mascota no atacará a ciudadanos si eres miembro del ejército real o tienes el seguro activado."
780     Case 28
790         tmpInteger = incomingData.ReadInteger
800         tmpString = incomingData.ReadASCIIString
            
810         MsjString = tmpString & "» Las inscripciones abren en " & _
                tmpInteger & " minutos."
820     Case 29
830         tmpInteger = incomingData.ReadInteger
840         tmpString = incomingData.ReadASCIIString
            
850         MsjString = tmpString & "» Las inscripciones abren en " & _
                tmpInteger & " minuto."
860     Case 30
870         tmpString = incomingData.ReadASCIIString
            
880         MsjString = tmpString & "» Inscripciones abiertas. /INGRESAR " & _
                tmpString & _
                " para ingresar al evento. /INFOEVENTO para que leas toda la información del evento en curso."
890     Case 31
900         MsjString = "Cuenta» ¡Comienza!"
910     Case 32
920         tmpInteger = incomingData.ReadInteger
            
930         MsjString = "Cuenta» " & tmpInteger
940     Case 33
950         MsjString = "No puedes participar en eventos estando muerto."
960     Case 34
970         MsjString = "No puedes entrar mimetizado."
980     Case 35
990         MsjString = "No puedes entrar montando."
1000    Case 36
1010        MsjString = "No puedes entrar invisible."
1020    Case 37
1030        MsjString = _
                "Ya te encuentras en un evento. Tipea /SALIREVENTO para salir del mismo."
1040    Case 38
1050        MsjString = _
                "No puedes participar de los eventos en la cárcel. Maldito prisionero!"
1060    Case 39
1070        MsjString = _
                "No puedes participar de los eventos estando en zona insegura. Vé a la ciudad mas cercana"
1080    Case 40
1090        MsjString = _
                "No puedes participar de los eventos si estás comerciando."
1100    Case 41
1110        MsjString = _
                "No hay ningun torneo disponible con ese nombre o bien las inscripciones no están disponibles aún."
1120    Case 42
1130        MsjString = _
                "El torneo ya ha comenzado. Mejor suerte para la próxima."
1140    Case 43
1150        MsjString = "Tu nivel no te permite ingresar a este evento."
1160    Case 44
1170        MsjString = "Tu clase no está permitida en el evento."
1180    Case 45
1190        MsjString = _
                "No tienes suficiente oro para pagar el torneo. Pide prestado a un compañero."
1200    Case 46
1210        MsjString = _
                "No tienes suficientes monedas DSP para participar del evento."
1220    Case 47
1230        MsjString = _
                "Los cupos del evento al que deseas participar ya fueron alcanzados."
1240    Case 48
1250        MsjString = _
                "No hay más lugar disponible para crear un evento simultaneo. Espera a que termine alguno o bien cancela alguno."
1260    Case 49
1270        tmpString = incomingData.ReadASCIIString
1280        MsjString = "¡¡ATENCIÓN GM!! Al personaje " & tmpString & _
                " no se le entrego el dsp porque no tenia espacio en el inventario."
1290    Case 50
1300        MsjString = _
                "¡¡Hemos notado que no tienes espacio en el inventario para recibir los DSP ganadores. Un gm se contactará contigo a la brevedad."
1310    Case 51
1320        tmpString = incomingData.ReadASCIIString
1330        MsjString = "Has ingresado al evento " & tmpString & _
                ". Espera a que se completen los cupos para que comience."
1340    Case 52
1350        tmpString = incomingData.ReadASCIIString
1360        MsjString = tmpString & _
                "» Los cupos han sido alcanzados. Les deseamos mucha suerte a cada uno de los participantes y que gane el mejor!"
1370    Case 53
1380        MsjString = _
                "Has abandonado el evento. Podrás recibir una pena por hacer esto."
1390    Case 54
1400        tmpLong = incomingData.ReadLong
1410        MsjString = "Felicitaciones, has recibido " & tmpLong & _
                " monedas de oro por haber ganado el evento!"
1420    Case 55
1430        tmpLong = incomingData.ReadLong
1440        MsjString = _
                "Hemos notado que has aniquilado con la vida del rey oponente. ¡FELICITACIONES! Aquí tienes tu recompensa! " _
                & tmpLong & " monedas de oro extra y su equipamiento"
1450    Case 56
1460        tmpString = incomingData.ReadASCIIString
1470        MsjString = "DagaRusa» El ganador es " & tmpString & _
                ". Felicitaciones para el personaje, quien se ha ganado una MD! (Espada mata dragones)"
1480    Case 57
1490        tmpString = incomingData.ReadASCIIString
1500        MsjString = "DeathMatch» El ganador es " & tmpString & _
                " quien se lleva 1 punto de torneo y 450.000 monedas de oro."
1510    Case 58
1520        tmpLong = incomingData.ReadLong
1530        MsjString = "Has recibido " & tmpLong & _
                " por haber aniquilado a todos los usuarios."
1540    Case 59
1550        tmpLong = incomingData.ReadLong
1560        tmpString = incomingData.ReadASCIIString
1570        MsjString = "Has recibido " & tmpLong & " por haber aniquilado a " _
                & tmpString
1580    Case 60
1590        MsjString = _
                "Has sido envenenado por Aracnus, has muerto de inmediato por su veneno letal."
1600    Case 61
1610        MsjString = _
                "¡El minotauro ha logrado paralizar tu cuerpo con su dosis de veneno. Has quedado afuera del evento."
1620    Case 62
1630        tmpString = incomingData.ReadASCIIString
1640        MsjString = _
                "Busqueda de objetos» El ganador de la búsqueda de objetos es " & _
                tmpString & _
                ". Felicitaciones! Se lleva como premio 350.000 monedas de oro"
1650    Case 63
1660        tmpInteger = incomingData.ReadInteger
1670        MsjString = "Has recolectado un objeto del piso. En total llevas " _
                & tmpInteger & " objetos recolectados. Sigue así!"
1680    Case 64
1690        MsjString = "Felicitaciones. Has ganado el enfrentamiento"
1700    Case 65
1710        MsjString = "Has ganado el evento. ¡Felicitaciones!"
1720    Case 66
1730        MsjString = _
                "Has sido aniquilado. Pero no pierdas las esperanzas joven guerrero, reviviste y tu sangre está hambrienta, ve trás el que te asesino y haz justicia!"
1740    Case 67
1750        tmpInteger = incomingData.ReadInteger
1760        MsjString = _
                "Felicitaciones, has sumado una muerte más a tu lista. Actualmente llevas " _
                & tmpInteger & " asesinatos. Sigue así y ganarás el evento."
1770    Case 68
1780        tmpInteger = incomingData.ReadInteger
1790        MsjString = "Felicitaciones. Tus " & tmpInteger & _
                " asesinatos han hecho que ganes el evento. Aquí tienes 500.000 monedas de oro como recompensa y un punto de torneo."
1800    Case 69
1810        tmpInteger = incomingData.ReadInteger
1820        tmpString = incomingData.ReadASCIIString
1830        MsjString = "Usuario Unstoppable» El ganador del evento es " & _
                tmpString & " con " & tmpInteger & _
                " asesinatos. Se lleva 500.000 monedas de oro y un punto de torneo."
        
        
        ' MENSAJES DE INICIO
        
1840    Case 70
1850        MsjString = ">Bienvenido a Desterium AO noble aventurero."
1860    Case 71
1870        MsjString = ">El máximo de oro es 200.000.000 por personaje."
1880    Case 72
1890        MsjString = _
                ">Recuerda visitar nuestro sitio oficial: https://www.desterium.com/"
1900    Case 73
1910        MsjString = _
                ">NO INSULTES. NO CHITEES. NO JUEGUES SUCIO. NO VALE LA PENA."
1920    Case 74
1930        MsjString = _
                ">El uso de cheats será sancionado de forma estricta por el Staff."
1940    Case 75
1950        MsjString = _
                "¡¡ATENCIÓN!! Tu personaje esta en modo VENTA por DSP o MONEDAS DE ORO. No podrán comprartelo. /QUITARPJ para quitarlo del MERCADO."
1960    Case 76
1970        MsjString = _
                "¡¡Tienes el poder de los Dioses. Tienes 30 puntos de vida extra. Aprovechalos. Cuando te inmovilicen, tendrás 40% de posibilidad de removerte solo!!"
        
        ' FIN MENSAJES DE INICIO
    
1980      End Select
    
1990      With FontTypes(FontIndex)
2000    AddtoRichTextBox frmMain.RecTxt, MsjString, .red, .green, .blue, _
    .bold, .italic

2010      End With
    
2020      Exit Sub

ErrHandler:
2030  Call LogError("Error en HandleShortMsj. Número " & Err.number & " Descripción: " _
    & Err.Description & " en linea " & Erl)
    
End Sub

Public Sub WriteChangeNick(ByVal UserName As String)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.ChangeNick)
30            Call .WriteASCIIString(UserName)
40        End With
End Sub

Public Sub WriteCanjeItem(ByVal CanjeItem As Integer, ByVal Value1 As Integer, _
    ByVal Value2 As Integer)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.CanjeItem)
30            Call .WriteInteger(CanjeItem)
40            Call .WriteInteger(Value1)
50            Call .WriteInteger(Value2)
60        End With
End Sub
Public Sub WriteCanjeInfo(ByVal CanjeItem As Integer, ByVal RequiredObj As _
    Integer, ByVal Points As Integer)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.CanjeInfo)
30            Call .WriteInteger(CanjeItem)
40            Call .WriteInteger(RequiredObj)
50            Call .WriteInteger(Points)
60        End With
End Sub
Private Sub HandleCanjeInit()
          Dim i      As Long, LoopC As Integer, LoopY As Integer
          
          'Remove packet ID
10        Call incomingData.ReadByte

20        NumCanjes = incomingData.ReadByte
          
30        If NumCanjes = 0 Then Exit Sub
          
          
          #If Wgl = 0 Then
40        Call InvCanje.Initialize(DirectDraw, FrmCanje.PicCanje, _
              MAX_BANCOINVENTORY_SLOTS)
           #Else
40                    Call InvCanje.Initialize(FrmCanje.PicCanje, _
              MAX_BANCOINVENTORY_SLOTS)
              
            
            #End If
          
50        ReDim Canjes(1 To MAX_BANCOINVENTORY_SLOTS) As tCanjes
          
60        For i = 1 To NumCanjes
70            With Canjes(i)
80                .NumRequired = incomingData.ReadByte
                  
                  'ReDim .ObjRequired() As Obj
                  
90                For LoopC = 1 To .NumRequired
100                   .ObjRequired(LoopC).ObjIndex = incomingData.ReadInteger
110                   .ObjRequired(LoopC).Amount = incomingData.ReadInteger
120               Next LoopC
                  
130               .ObjCanje.ObjIndex = incomingData.ReadInteger
140               .ObjCanje.Amount = incomingData.ReadInteger
150               .GrhIndex = incomingData.ReadInteger
160               .Points = incomingData.ReadInteger
                  
170           End With
180       Next i
          
190       For i = 1 To MAX_BANCOINVENTORY_SLOTS
200           If Canjes(i).GrhIndex <> 0 Then
210               Call InvCanje.SetItem(i, Canjes(i).ObjCanje.ObjIndex, _
                      Canjes(i).ObjCanje.Amount, 0, Canjes(i).GrhIndex, 0, 0, 0, 0, 0, 0, _
                      ObjName(Canjes(i).ObjCanje.ObjIndex).Name)
220           End If
230       Next i

          'Set state and show form
240       Canjeando = True
          
250       FrmCanje.Show , frmMain
End Sub

Public Sub HandleCanjeEnd()

10        Set InvCanje = Nothing
20        Unload FrmCanje
          
30        Canjeando = False
End Sub

Private Sub HandleInfoCanje()
10        Call incomingData.ReadByte
          
          Dim strTemp As String
          Dim LoopC As Integer
          
20        With FrmCanje
30            .lblDef = incomingData.ReadInteger & "/" & incomingData.ReadInteger
40            .lblRM = incomingData.ReadInteger & "/" & incomingData.ReadInteger
50            .lblAtaqueFisico = incomingData.ReadInteger & "/" & _
                  incomingData.ReadInteger
60            .lblPuntos = incomingData.ReadLong
70            .lblSeCae = IIf(incomingData.ReadByte = 0, "SI", "NO")
              
80            .lblRequired.Caption = vbNullString
              
90            For LoopC = 1 To Canjes(InvCanje.SelectedItem).NumRequired
100               If LoopC = 1 Then
110                   .lblRequired = _
                          ObjName(Canjes(InvCanje.SelectedItem).ObjRequired(LoopC).ObjIndex).Name _
                          & " (x" & _
                          Canjes(InvCanje.SelectedItem).ObjRequired(LoopC).Amount & ")"
120               Else
130                   .lblRequired = .lblRequired & vbcrlf & _
                          ObjName(Canjes(InvCanje.SelectedItem).ObjRequired(LoopC).ObjIndex).Name _
                          & " (x" & _
                          Canjes(InvCanje.SelectedItem).ObjRequired(LoopC).Amount & ")"
140               End If
150           Next LoopC
          
          
160       End With
          
End Sub

Public Sub WriteSendFight(ByVal GldRequired As Long, ByVal DspRequired As Long, _
    ByVal LimiteRojas As Integer, ByVal UserName As String)
                                  
          Dim LoopC As Integer
          
10        With outgoingData
20            .WriteByte ClientPacketID.PacketRetos
30            .WriteByte 0
40            .WriteLong GldRequired
50            .WriteLong DspRequired
60            .WriteInteger LimiteRojas
70            .WriteASCIIString UCase$(UserName)
          
80        End With
End Sub

Public Sub WriteAcceptFight(ByVal UserName As String)
10        With outgoingData
20            .WriteByte ClientPacketID.PacketRetos
30            .WriteByte 1
40            .WriteASCIIString UCase$(UserName)
50        End With
End Sub

Public Sub WriteByeFight()
10        With outgoingData
20            .WriteByte ClientPacketID.PacketRetos
30            .WriteByte 2
40        End With
End Sub

Public Sub WriteSendFightClan(ByVal UserName As String)
10        With outgoingData
20            .WriteByte ClientPacketID.PacketRetos
30            .WriteByte 3
40            .WriteASCIIString UCase$(UserName)
50        End With
End Sub
Public Sub WriteAcceptFightClan(ByVal UserName As String)
10        With outgoingData
20            .WriteByte ClientPacketID.PacketRetos
30            .WriteByte 4
40            .WriteASCIIString UCase$(UserName)
50        End With
End Sub
Public Sub WriteRequestRetos()
10        With outgoingData
20            .WriteByte ClientPacketID.PacketRetos
30            .WriteByte 5
40        End With
End Sub

Public Sub WriteUseItem(ByVal Slot As Byte, ByVal SecondaryClick As Byte)

        If Inventario.SelectedItem > 25 Then Exit Sub
        If frmMain.picInv.Visible = False And SecondaryClick = 1 Then Exit Sub
        
10        With outgoingData
20            .WriteByte ClientPacketID.UseItemPacket
30            .WriteByte Slot
40            .WriteByte SecondaryClick
              .WriteLong KeyPackets(eKeyPackets.Key_UseItem)
              .WriteByte ClaveActual
              
50        End With

End Sub

Public Sub WriteRequestInfoEvent()
10        With outgoingData
20            Call .WriteByte(ClientPacketID.RequestInfoEvento)
30        End With
End Sub

Public Sub HandlePacketGambleSv()

10    On Error GoTo ErrHandler

20        Call incomingData.ReadByte
    
    Dim Tipo As Byte
    Dim Users() As String
    Dim LoopC As Integer
    
30
40        Tipo = incomingData.ReadByte
50
60        Select Case Tipo
        Case 0 ' Recibimos la lista de usuarios que apostaron
70          Users = Split(incomingData.ReadASCIIString, "-")
80
90          FrmApuestasGM.lstUsers.Clear
            
100         For LoopC = LBound(Users()) To UBound(Users())
110             FrmApuestasGM.lstUsers.AddItem Users(LoopC)
120         Next LoopC
130
140     Case 1 ' Recibimos la info de los usuarios que apostaron
150         FrmApuestasGM.lblDsp.Caption = "Dsp: " & incomingData.ReadLong
160         FrmApuestasGM.lblGld.Caption = "Oro: " & incomingData.ReadLong
170
            
180     Case 2 ' Recibimos la lista de apuestas disponibles para los usuarios
190           Users = Split(incomingData.ReadASCIIString, ",")
200
210         For LoopC = LBound(Users()) To UBound(Users())
220             FrmApuestas.lstApuestas.AddItem Users(LoopC)
230         Next LoopC
            
240
250         FrmApuestas.Show vbModeless, frmMain
260       End Select
270       Exit Sub

ErrHandler:
280   Call LogError("Error en HandlePacketGambleSv. Número " & Err.number & _
    " Descripción: " & Err.Description & " en linea " & Erl)
End Sub

Public Sub WriteRequestUsersApostando()
10        With outgoingData
20            Call .WriteByte(ClientPacketID.PacketGamble)
30            Call .WriteByte(4)
40        End With
End Sub

Public Sub WriteRequestInfoUserApostando(ByVal UserName As String)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.PacketGamble)
30            Call .WriteByte(5)
40            Call .WriteASCIIString(UserName)
50        End With
End Sub

Public Sub WriteNewGamble(ByVal Desc As String, ByVal TimeFinish As Integer, _
    ByVal UsersAmount As Byte, ByVal Apuestas As String)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.PacketGamble)
30            Call .WriteByte(0)
              
40            Call .WriteASCIIString(Apuestas)
50            Call .WriteASCIIString(Desc)
60            Call .WriteInteger(TimeFinish)
70            Call .WriteByte(UsersAmount)
          
80        End With
End Sub

Public Sub WriteCancelGamble()
10        With outgoingData
20            Call .WriteByte(ClientPacketID.PacketGamble)
30            Call .WriteByte(1)
40        End With
End Sub

Public Sub WriteSendGamble(ByVal ApuestaIndex As Byte, ByVal Dsp As Long, ByVal _
    Gld As Long)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.PacketGamble)
30            Call .WriteByte(3)
40            Call .WriteByte(ApuestaIndex)
50            Call .WriteLong(Gld)
60            Call .WriteLong(Dsp)

70        End With
End Sub

Public Sub WriteRequestApuestas()
10        With outgoingData
20            Call .WriteByte(ClientPacketID.PacketGamble)
30            Call .WriteByte(6)
40        End With
End Sub

Public Sub WriteWinGamble(ByVal UserName As String)
10        With outgoingData
20            Call .WriteByte(ClientPacketID.PacketGamble)
30            Call .WriteByte(7)
40            Call .WriteASCIIString(UserName)
50        End With
End Sub

Public Sub WritePartyClient(ByVal Paso As Byte)
    With outgoingData
    
        ' 1) Requiere formulario 'principal'
        ' 2) Requiere formulario 'solicitudes'
        ' 3) Requiere formulario 'obtenido'
        .WriteByte ClientPacketID.PartyClient
        .WriteByte Paso
    End With
End Sub
Public Sub WriteGroupMember(ByVal Tipo As Byte, ByVal UserName As String)
    With outgoingData
    
        ' 1) Aceptar
        ' 2) Rechazar
        .WriteByte ClientPacketID.GroupMember
        .WriteByte Tipo
        .WriteASCIIString UserName
    End With
End Sub

Public Sub WriteGroupChangePorc(ByRef PorcExp() As Byte, ByRef PorcGld() As Byte)
    Dim A As Byte
    
    With outgoingData
        .WriteByte ClientPacketID.GroupChangePorc
        
        For A = 0 To 4
            .WriteByte (PorcExp(A))
            .WriteByte (PorcGld(A))
        Next A
    End With
End Sub
Public Sub HandleGroupPrincipal()
    Dim A As Long
    Dim Bonus(0 To 3) As Boolean
    
    Call incomingData.ReadByte
    
    For A = 1 To MAX_MEMBERS_GROUP
        With Groups
            .User(A).Name = incomingData.ReadASCIIString
            .User(A).PorcExp = incomingData.ReadByte
            .User(A).PorcGld = incomingData.ReadByte
            
            frmParty.lblUser(A - 1) = .User(A).Name
            frmParty.lblExp(A - 1) = .User(A).PorcExp
            frmParty.lblOro(A - 1) = .User(A).PorcGld
        End With
    Next A
    
    Bonus(0) = incomingData.ReadBoolean
    Bonus(1) = incomingData.ReadBoolean
    Bonus(2) = incomingData.ReadBoolean
    Bonus(3) = incomingData.ReadBoolean
    
    With frmParty
        For A = 0 To 3
            If Bonus(A) Then
                .imgBons(A).Picture = LoadPicture(App.path & "\Recursos\Grupos\SI.JPG")
            Else
                .imgBons(A).Picture = LoadPicture(App.path & "\Recursos\Grupos\NO.JPG")
            End If
        Next A
        
        If .Visible = False Then .Show vbModeless, frmMain
    End With
    
    
End Sub

Public Sub HandleGroupRequests()
    Dim A As Long
    
    Call incomingData.ReadByte
    
    frmParty.lstRequest.Clear
    
    For A = 1 To MAX_REQUESTS_GROUP
        With Groups
            .Requests(A) = incomingData.ReadASCIIString

            If .Requests(A) <> vbNullString Then
                frmParty.lstRequest.AddItem .Requests(A)
            End If
        End With
    Next A
    
End Sub

Public Sub HandleGroupReward()
    Call incomingData.ReadByte

    frmParty.lblExpObtenida.Caption = incomingData.ReadLong
    frmParty.lblOroObtenido.Caption = incomingData.ReadLong
End Sub

Private Sub HandleUpdateKey()
    Call incomingData.ReadByte
    
    Dim Packet As Byte
    Dim KeyPacket As Long
    
    Packet = incomingData.ReadByte
    KeyPacket = incomingData.ReadLong
    
    KeyPackets(Packet) = KeyPacket
    
   ' ShowConsoleMsg "ClienteKey n°" & Packet & ", Key: " & KeyPacket
End Sub

Public Sub WriteSendCaptureImage(ByRef Bytes() As Byte)
    Dim A As Long
    
    With outgoingData
        '.WriteByte ClientPacketID.SendCaptureImage
        
        Debug.Print UBound(Bytes)
        .WriteLong UBound(Bytes)
        
        For A = 0 To UBound(Bytes)
            .WriteByte Bytes(A)
        Next A
    
    End With
End Sub

Public Sub WriteNewAccount()
    With outgoingData
        Call .WriteByte(ClientPacketID.PacketAccount)
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString("Create$Account")
        Call .WriteASCIIString(Cuenta.Account)
        Call .WriteASCIIString(Cuenta.Passwd)
        Call .WriteASCIIString(Cuenta.PIN)
        Call .WriteASCIIString(Cuenta.Email)
    End With
End Sub

Public Sub WriteLoginAccount()
    With outgoingData
        Call .WriteByte(ClientPacketID.PacketAccount)
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString("Login$Account")
        Call .WriteASCIIString(Cuenta.Account)
        Call .WriteASCIIString(Cuenta.Passwd)
    End With
End Sub
Public Sub WriteRecoverAccount()
    With outgoingData
        Call .WriteByte(ClientPacketID.PacketAccount)
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString("Recover$Account")
        Call .WriteASCIIString(Cuenta.Account)
        Call .WriteASCIIString(Cuenta.PIN)
        Call .WriteASCIIString(Cuenta.Email)
    End With
End Sub
Public Sub WriteChangePasswdAccount()
    With outgoingData
        Call .WriteByte(ClientPacketID.PacketAccount)
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString("ChangePasswd$Account")
        Call .WriteASCIIString(Cuenta.Account)
        Call .WriteASCIIString(Cuenta.PIN)
        Call .WriteASCIIString(Cuenta.Email)
        Call .WriteASCIIString(Cuenta.Passwd)
        Call .WriteASCIIString(Cuenta.NewPasswd)
    End With
End Sub
Public Sub WriteLoginCharAccount()
    With outgoingData
        Call .WriteByte(ClientPacketID.PacketAccount)
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString("LoginChar$Account")
        Call .WriteASCIIString(Cuenta.Account)
        Call .WriteASCIIString(Cuenta.Passwd)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteCreateCharAccount()
    With outgoingData
        Call .WriteByte(ClientPacketID.PacketAccount)
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString("CreateChar$Account")
        Call .WriteASCIIString(Cuenta.Account)
        Call .WriteASCIIString(Cuenta.Passwd)
        Call .WriteASCIIString(UserName)
        Call .WriteByte(UserClase)
        Call .WriteByte(UserRaza)
        Call .WriteByte(UserSexo)

    End With
End Sub

Public Sub WriteRemoveCharAccount()
    With outgoingData
        Call .WriteByte(ClientPacketID.PacketAccount)
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString("RemoveChar$Account")
        
        Call .WriteASCIIString(Cuenta.Account)
        Call .WriteASCIIString(Cuenta.Passwd)
        Call .WriteByte(SelectedChar)
    End With
End Sub

Public Sub WriteAddTemporal()
    With outgoingData
        Call .WriteByte(ClientPacketID.PacketAccount)
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString("AddTemporal$Account")
        
        Call .WriteASCIIString(Cuenta.Account)
        Call .WriteASCIIString(Cuenta.Passwd)
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(UserPassword)
        Call .WriteASCIIString(UserEmail)
        Call .WriteASCIIString(UserPin)
        
    End With
End Sub

Public Sub handleSendClave()
    Call incomingData.ReadByte
    ClaveActual = incomingData.ReadByte
    
End Sub
Public Sub HandleAccount_Data()

    Dim LoopC As Integer
    
    Call incomingData.ReadByte
    
    FrmCuenta.lstPjs.Clear
    
    For LoopC = 1 To MAX_PJS_ACCOUNT
        With CuentaChars(LoopC)
            .Name = incomingData.ReadASCIIString
            .Ban = incomingData.ReadByte
            .Clase = incomingData.ReadByte
            .Elv = incomingData.ReadByte
            .Raza = incomingData.ReadByte
            
            If .Name <> "0" Then
                FrmCuenta.lstPjs.AddItem .Name
            Else
                FrmCuenta.lstPjs.AddItem "(Vacio)"
            End If
        End With
    Next LoopC
    
    If Not FrmCuenta.Visible Then
        FrmCuenta.Show
    End If
End Sub
Public Sub WriteSearchObj(ByVal BuscoObj As String)
 
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SearchObj)
           
        Call .WriteASCIIString(BuscoObj)
    End With
End Sub
 
Public Sub WriteSearchNpc(ByVal BuscoNpc As String)
 
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SearchNpc)
       
        Call .WriteASCIIString(BuscoNpc)
    End With
End Sub
 
Private Sub HandleListText()
 
    Dim Num As Integer
    Dim Datos As String
    Dim Obj As Boolean
       
    Call incomingData.ReadByte
   
    Num = incomingData.ReadInteger()
    Obj = incomingData.ReadBoolean()
 
    If Not Num = 0 Then
        If Obj = True Then
            frmBuscar.ListCrearObj.AddItem Num
        Else
            frmBuscar.ListCrearNpcs.AddItem Num
        End If
    End If
 
    Datos = incomingData.ReadASCIIString()
 
    frmBuscar.List1.AddItem Datos
 
End Sub
 
Public Sub WriteSearcherShow()
 
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SearcherShow)
 
End Sub

Public Sub WriteEnviarAviso(ByVal Tipo As Byte)
    Call outgoingData.WriteByte(ClientPacketID.EnviarAviso)
    Call outgoingData.WriteByte(Tipo)
End Sub
 Public Sub WriteSeguroClan() '(ByVal Tipo As Byte)
    Call outgoingData.WriteByte(ClientPacketID.seguroclan)
End Sub
 
Private Sub HandleShowSearcher()
 
    Call incomingData.ReadByte
   
    frmBuscar.Show , frmMain
 
End Sub
