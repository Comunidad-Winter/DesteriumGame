Attribute VB_Name = "Mod_Declaraciones"
Option Explicit

'Public CharFichado As Integer

Public SelectedListMAO As Integer


''''''''''''''''''''''''

Public Enum eKeyPackets
    Key_UseItem = 0
    Key_UseSpell = 1
    Key_UseWeapon = 2
End Enum

Public Const MAX_KEY_PACKETS As Byte = 2
Public KeyPackets(MAX_KEY_PACKETS) As Long


' GRUPOS
Public Const MAX_MEMBERS_GROUP As Byte = 5
Public Const MAX_REQUESTS_GROUP As Byte = 10
Private Const MAX_GROUPS As Byte = 100
Private Const SLOT_LEADER As Byte = 1

Private Const EXP_BONUS_MAX_MEMBERS As Single = 1.05 '%
Private Const EXP_BONUS_LEADER_PREMIUM As Single = 1.05 '%

Public Type tUserGroup
    Name As String
    Exp As Long
    Gld As Long
    PorcExp As Byte
    PorcGld As Byte
End Type

Public Type tGroups
    members As Byte
    User(1 To MAX_MEMBERS_GROUP) As tUserGroup
    Requests(1 To MAX_REQUESTS_GROUP) As String
End Type

Public Groups As tGroups




' FIN GRUPOS

Public UserPoints As Long

Public LastCharIndex As Integer

Public CantF8 As Integer
Public CantKey0 As Integer
Public CantKey1 As Integer
Public CantKey2 As Integer

' OPTIMIZACIÓN DE STRINGS EN EL CLIENTE PARA ANTI LAG.
Public Type tHechizos
    Name As String
    Desc As String
    PalabrasMagicas As String
    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String
End Type

Public Hechizos() As tHechizos
Public NumHechizos As Byte

Public Type tObj
    Name As String
End Type

Public ObjName() As tObj
Public NumObjs As Integer

Public Type tNpcs
    Name As String
    Desc As String
End Type

Public Npc() As tNpcs
Public NumNpcs As Integer



'#######
' MENUES

Public UserEvento As Boolean

Public SeguridadCRC() As Byte
Public CRC As Integer

Public Type tMenuAction
    NormalGrh As Integer
    FocusGrh As Integer
    ActionIndex As Byte
End Type

Public Type tMenu
    NumActions As Byte
    Actions() As tMenuAction
End Type

Public MenuInfo() As tMenu

Public Enum eMenues
    ieUser = 1
    ieNpcComercio = 2
    ieNpcNoHostil = 3
    
End Enum

Public Type tHeadsHombre
    Humano(1 To 25) As Integer
    Elfo(1 To 9) As Integer
    ElfoDrow(1 To 5) As Integer
    Gnomo(1 To 4) As Integer
    Enano(1 To 4) As Integer
End Type


Public Type tHeadsMujer
    Humano(1 To 5) As Integer
    Elfo(1 To 7) As Integer
    ElfoDrow(1 To 6) As Integer
    Gnomo(1 To 5) As Integer
    Enano(1 To 2) As Integer
End Type

Public HeadMujer As tHeadsMujer
Public HeadHombre As tHeadsHombre

'**************************************************
Public Const HUMANO_M_PRIMER_CABEZA As Integer = 71
Public Const HUMANO_M_ULTIMA_CABEZA As Integer = 75

Public Const ELFO_M_PRIMER_CABEZA As Integer = 170
Public Const ELFO_M_ULTIMA_CABEZA As Integer = 174

Public Const DROW_M_PRIMER_CABEZA As Integer = 270
Public Const DROW_M_ULTIMA_CABEZA As Integer = 276

Public Const ENANO_M_PRIMER_CABEZA As Integer = 370
Public Const ENANO_M_ULTIMA_CABEZA As Integer = 371

Public Const GNOMO_M_PRIMER_CABEZA As Integer = 471
Public Const GNOMO_M_ULTIMA_CABEZA As Integer = 475

Public Type tRanking
    value(0 To 9) As Long
    Nombre(0 To 9) As String
End Type


Public Ranking As tRanking

Public Enum eRanking
    TopFrags = 1
    TopTorneos = 2
    TopLevel = 3
    TopOro = 4
    TopRetos = 5
    TopClanes = 6
End Enum


#If Testeo = 1 Then
    Public Const CurServerIp As String = "localhost"
    Public Const CurServerPort As Integer = 7666
#Else
    Public Const CurServerIp As String = "localhost"
    Public Const CurServerPort As Integer = 7666
#End If

Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal _
    dwReserved As Long) As Long
 
Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4

Public Declare Function vbDABLalphablend16 Lib "vbDABL" (ByVal iMode As Integer, _
    ByVal bColorKey As Integer, ByRef sPtr As Any, ByRef dPtr As Any, ByVal _
    iAlphaVal As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, ByVal _
    isPitch As Integer, ByVal idPitch As Integer, ByVal iColorKey As Integer) As _
    Integer
Public Declare Function vbDABLcolorblend16555 Lib "vbDABL" (ByRef sPtr As Any, _
    ByRef dPtr As Any, ByVal alpha_val%, ByVal Width%, ByVal Height%, ByVal sPitch%, _
    ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
Public Declare Function vbDABLcolorblend16565 Lib "vbDABL" (ByRef sPtr As Any, _
    ByRef dPtr As Any, ByVal alpha_val%, ByVal Width%, ByVal Height%, ByVal sPitch%, _
    ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
Public Declare Function vbDABLcolorblend16555ck Lib "vbDABL" (ByRef sPtr As Any, _
    ByRef dPtr As Any, ByVal alpha_val%, ByVal Width%, ByVal Height%, ByVal sPitch%, _
    ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
Public Declare Function vbDABLcolorblend16565ck Lib "vbDABL" (ByRef sPtr As Any, _
    ByRef dPtr As Any, ByVal alpha_val%, ByVal Width%, ByVal Height%, ByVal sPitch%, _
    ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long

Public Velocidad As Byte

'----Boton party Style TDS by IRuleDK----
Public PMSGimg As Boolean
Public PMSG As Boolean
'----Boton party Style TDS by IRuleDK----

Public datCM() As Byte ' clave maestra

Public InvTemp    As New clsGrapchicalInventory

Public Enum eMoveType
    Inventory = 1
    Bank
End Enum

Public UserPin As String
Public CANTt As Byte
'Fluidez
Public Movement_Speed As Single

Public Type tMotd
    Texto As String
End Type
 
Public MaxLines As Integer
Public MOTD() As tMotd

'To Put Grafical Cursors
Public Const GLC_HCURSOR = (-12)
Public hSwapCursor As Long
Public Declare Function LoadCursorFromFile Lib "user32" Alias _
    "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal _
    hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Objetos públicos
Public DialogosClanes As New clsGuildDlg

#If Wgl = 0 Then
Public Dialogos As New clsDialogs
#Else
Public Dialogos As New clsDialogsWgl
#End If

Public ConAlfaB As Byte
Public Audio As New clsAudio


Public Const MAX_LIST_ITEMS As Byte = 4

#If Wgl = 0 Then
    Public Inventario As New clsGrapchicalInventory
    Public InvBanco(1) As New clsGrapchicalInventory
    Public InvCanje As New clsGrapchicalInventory
    Public Inventario_InfoPj As New clsGrapchicalInventory
    Public Boveda_InfoPj As New clsGrapchicalInventory
    
    'Inventarios de comercio con usuario
    Public InvComUsu As New clsGrapchicalInventory ' Inventario del usuario visible en el comercio
    Public InvOroComUsu(2) As New clsGrapchicalInventory ' Inventarios de oro (ambos usuarios)
    Public InvOfferComUsu(1) As New clsGrapchicalInventory ' Inventarios de ofertas (ambos usuarios)
    Public InvComNpc As New clsGrapchicalInventory ' Inventario con los items que ofrece el npc
    'Inventarios de herreria
    Public InvLingosHerreria(1 To MAX_LIST_ITEMS) As New clsGrapchicalInventory
    Public InvMaderasCarpinteria(1 To MAX_LIST_ITEMS) As New clsGrapchicalInventory
#Else
    Public Inventario As New clsGrapchicalInventoryWgl
    Public InvBanco(1) As New clsGrapchicalInventoryWgl
    Public InvCanje As New clsGrapchicalInventoryWgl
    Public Inventario_InfoPj As New clsGrapchicalInventoryWgl
    Public Boveda_InfoPj As New clsGrapchicalInventoryWgl
    
    'Inventarios de comercio con usuario
    Public InvComUsu As New clsGrapchicalInventoryWgl ' Inventario del usuario visible en el comercio
    Public InvOroComUsu(2) As New clsGrapchicalInventoryWgl ' Inventarios de oro (ambos usuarios)
    Public InvOfferComUsu(1) As New clsGrapchicalInventoryWgl ' Inventarios de ofertas (ambos usuarios)
    Public InvComNpc As New clsGrapchicalInventoryWgl ' Inventario con los items que ofrece el npc
    'Inventarios de herreria
    Public InvLingosHerreria(1 To MAX_LIST_ITEMS) As New clsGrapchicalInventoryWgl
    Public InvMaderasCarpinteria(1 To MAX_LIST_ITEMS) As New clsGrapchicalInventoryWgl


#End If


                
Public SurfaceDB As clsSurfaceManager   'No va new porque es una interfaz, el new se pone al decidir que clase de objeto es
Public CustomKeys As New clsCustomKeys
Public CustomMessages As New clsCustomMessages

Public incomingData As New clsByteQueue
Public outgoingData As New clsByteQueue

''
'The main timer of the game.
Public MainTimer As New clsTimer

'Sonidos
Public Const SND_CLICK As String = "click.Wav"
Public Const SND_PASOS1 As String = "23.Wav"
Public Const SND_PASOS2 As String = "24.Wav"
Public Const SND_PASOS3 As String = "501.Wav" 'Pie 1 de Arena
Public Const SND_PASOS4 As String = "502.Wav" 'Pie 2 de arena
Public Const SND_PASOS5 As String = "503.Wav" 'Pie 1 de  nieve
Public Const SND_PASOS6 As String = "504.Wav" 'Pie 2 de nieve
Public Const SND_PASOS7 As String = "505.Wav" 'Pie 1 de pasto
Public Const SND_PASOS8 As String = "506.Wav" 'Pie 2 de pasto
Public Const SND_NAVEGANDO As String = "50.wav"
Public Const SND_OVER As String = "click2.Wav"
Public Const SND_DICE As String = "cupdice.Wav"
Public Const SND_LLUVIAINEND As String = "lluviainend.wav"
Public Const SND_LLUVIAOUTEND As String = "lluviaoutend.wav"

' Head index of the casper. Used to know if a char is killed

' Constantes de intervalo
Public Const INT_MACRO_HECHIS As Integer = 2700
Public Const INT_MACRO_TRABAJO As Integer = 900

Public Const INT_ARROWS As Integer = 1000
Public Const INT_WORK As Integer = 750
Public Const INT_USEITEMDCK As Integer = 100
Public Const INT_SENTRPU As Integer = 2000

Public Const INT_ATTACK            As Long = 1250
Public Const INT_CAST_ATTACK      As Long = 950
Public Const INT_ATTACK_CAST      As Long = 800
Public Const INT_CAST_SPELL       As Long = 1000
Public Const INT_USEITEMU         As Long = 435

Public MacroBltIndex As Integer

Public Const CASPER_HEAD As Integer = 500
Public Const FRAGATA_FANTASMAL As Integer = 87

Public Const NUMATRIBUTES As Byte = 5


'Musica
Public Const MP3_Inicio As Byte = 101
Public Const MP3_CREARPJ As Byte = 102

Public RawServersList As String

Public Type tColor
    r As Byte
    g As Byte
    b As Byte
End Type

Public ColoresPJ(0 To 50) As tColor


Public Type tServerInfo
    Ip As String
    Puerto As Integer
    Desc As String
    PassRecPort As Integer
End Type

Public ServersLst() As tServerInfo
Public ServersRecibidos As Boolean

Public CurServer As Integer

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public UserCiego As Boolean
Public UserEstupido As Boolean

Public NoRes As Boolean 'no cambiar la resolucion
Public GraphicsFile As String 'Que graficos.ind usamos

Public RainBufferIndex As Long
Public FogataBufferIndex As Long

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 2000
Public Const tUs = 600

Public Const PrimerBodyBarco = 84
Public Const UltimoBodyBarco = 87

Public NumEscudosAnims As Integer

Public ArmasHerrero(0 To 100) As Integer
Public ArmadurasHerrero(0 To 100) As Integer
Public ObjCarpintero(0 To 100) As Integer
Public Versiones(1 To 7) As Integer



Public CarpinteroMejorar() As tItemsConstruibles
Public HerreroMejorar() As tItemsConstruibles

Public UsaMacro As Boolean
Public CnTd As Byte


Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
Public UserBancoCanjes(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory

Public TradingUserName As String

Public Tips() As String * 255
Public Const LoopAdEternum As Integer = 999

'Direcciones
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'Objetos
Public Const MAX_INVENTORY_OBJS As Integer = 10000
Public Const MAX_INVENTORY_SLOTS As Byte = 30
Public Const MAX_NORMAL_INVENTORY_SLOTS As Byte = 25
Public Const MAX_NPC_INVENTORY_SLOTS As Byte = 50
Public Const MAXHECHI As Byte = 35

Public Const INV_OFFER_SLOTS As Byte = 20
Public Const INV_GOLD_SLOTS As Byte = 1

Public Const MAXSKILLPOINTS As Byte = 100

Public Const MAXSKILLPOINTSL As Byte = 90

Public Const MAXATRIBUTOS As Byte = 38
Public Const MAX_OFFER_SLOTS As Integer = 20
Public Const FLAGORO As Integer = MAX_NORMAL_INVENTORY_SLOTS + 1


Public Const GOLD_OFFER_SLOT As Integer = INV_OFFER_SLOTS + 1
Public Const FOgata As Integer = 1521


Public Enum eClass
    Mage = 1    'Mago
    Cleric      'Clérigo
    Warrior     'Guerrero
    Assasin     'Asesino
    Thief       'Ladrón
    Bard        'Bardo
    Druid       'Druida
    'Bandit       'Bandido
    Paladin     'Paladín
    Hunter      'Cazador
    Worker      'Trabajador
    Pirat       'Pirata
End Enum

Public Enum eCiudad
    cUllathorpe = 1
    cNix
    cBanderbill
    cLindos
    cArghal
End Enum

Enum eRaza
    Humano = 1
    Elfo
    ElfoOscuro
    Gnomo
    Enano
End Enum

Public Enum eSkill
    Magia = 1
    Robar = 2
    Tacticas = 3
    Armas = 4
    Meditar = 5
    Apuñalar = 6
    Ocultarse = 7
    Supervivencia = 8
    Talar = 9
    Comerciar = 10
    Defensa = 11
    Pesca = 12
    Mineria = 13
    Carpinteria = 14
    Herreria = 15
    Liderazgo = 16
    Domar = 17
    Proyectiles = 18
    Wrestling = 19
    Navegacion = 20
    Equitacion = 21
    Resistencia = 22
End Enum

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Enum eGenero
    Hombre = 1
    Mujer
End Enum

Public Enum PlayerType
    User = &H1
    Consejero = &H2
    SemiDios = &H4
    Dios = &H8
    Admin = &H10
    RoleMaster = &H20
    ChaosCouncil = &H40
    RoyalCouncil = &H80
End Enum

Public Enum eOBJType
    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otarboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otBebidas = 13
    otLeña = 14
    otFogata = 15
    otescudo = 16
    otcasco = 17
    otAnillo = 18
    otTeleport = 19
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otGemas = 45
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManchas = 35          'No se usa
    otArbolElfico = 36
    otMochilas = 37
    otMonturas = 38
    otMonturasDraco = 39
    otLunar = 40
    otAzul = 46
    otNaranja = 47
    otCeleste = 48
    otLila = 49
    otroja = 50
    otverde = 51
    otvioleta = 52
    otAnilloNpc = 53
    otCualquiera = 1000
End Enum

 
Public Const FundirMetal As Integer = 88

' Determina el color del nick
Public Enum eNickColor
    ieCriminal = &H1
    ieCiudadano = &H2
    ieAtacable = &H4
    ieTeamUno = &H8
    ieTeamDos = &H10
End Enum

Public Enum eGMCommands
    GMMessage = 1           '/GMSG
    showName = 2              '/SHOWNAME
    OnlineRoyalArmy = 3       '/ONLINEREAL
    OnlineChaosLegion = 4     '/ONLINECAOS
    GoNearby = 5              '/IRCERCA
    SeBusca = 6              '/SEBUSCA
    comment = 7               '/REM
    serverTime = 8            '/HORA
    Where = 9                 '/DONDE
    CreaturesInMap = 10        '/NENE
    WarpMeToTarget = 11        '/TELEPLOC
    WarpChar = 12             '/TELEP
    Silence = 13               '/SILENCIAR
    SOSShowList = 14           '/SHOW SOS
    SOSRemove = 15             'SOSDONE
    GoToChar = 16              '/IRA
    Invisible = 17             '/INVISIBLE
    GMPanel = 18               '/PANELGM
    RequestUserList = 19       'LISTUSU
    Working = 20               '/TRABAJANDO
    Hiding = 21                '/OCULTANDO
    Jail = 22                  '/CARCEL
    KillNPC = 23               '/RMATA
    WarnUser = 24              '/ADVERTENCIA
    RequestCharInfo = 25       '/INFO
    RequestCharStats = 26      '/STAT
    RequestCharGold = 27       '/BAL
    RequestCharInventory = 28  '/INV
    RequestCharBank = 29       '/BOV
    RequestCharSkills = 30     '/SKILLS
    ReviveChar = 31            '/REVIVIR
    OnlineGM = 32              '/ONLINEGM
    OnlineMap = 33             '/ONLINEMAP
    Forgive = 34               '/PERDON
    Kick = 35                  '/ECHAR
    Execute = 36               '/EJECUTAR
    banChar = 37               '/BAN
    UnbanChar = 38             '/UNBAN
    NPCFollow = 39             '/SEGUIR
    SummonChar = 40            '/SUM
    SpawnListRequest = 41      '/CC
    SpawnCreature = 42         'SPA
    ResetNPCInventory = 43     '/RESETINV
    cleanworld = 44            '/LIMPIAR
    ServerMessage = 45         '/RMSG
    RolMensaje = 46            '/ROLEANDO
    nickToIP = 47              '/NICK2IP
    IPToNick = 48              '/IP2NICK
    GuildOnlineMembers = 49    '/ONCLAN
    TeleportCreate = 50        '/CT
    TeleportDestroy = 51       '/DT
    RainToggle = 52            '/LLUVIA
    SetCharDescription = 53    '/SETDESC
    ForceMIDIToMap = 54        '/FORCEMIDIMAP
    ForceWAVEToMap = 55        '/FORCEWAVMAP
    RoyalArmyMessage = 56      '/REALMSG
    ChaosLegionMessage = 57    '/CAOSMSG
    CitizenMessage = 58        '/CIUMSG
    CriminalMessage = 59       '/CRIMSG
    TalkAsNPC = 60             '/TALKAS
    DestroyAllItemsInArea = 61 '/MASSDEST
    AcceptRoyalCouncilMember = 62 '/ACEPTCONSE
    AcceptChaosCouncilMember = 63 '/ACEPTCONSECAOS
    ItemsInTheFloor = 64       '/PISO
    MakeDumb = 65              '/ESTUPIDO
    MakeDumbNoMore = 66        '/NOESTUPIDO
    dumpIPTables = 67          '/DUMPSECURITY
    CouncilKick = 68           '/KICKCONSE
    SetTrigger = 69            '/TRIGGER
    AskTrigger = 70            '/TRIGGER with no args
    BannedIPList = 71          '/BANIPLIST
    BannedIPReload = 72        '/BANIPRELOAD
    GuildMemberList = 73       '/MIEMBROSCLAN
    GuildBan = 74              '/BANCLAN
    BanIP = 75                 '/BANIP
    UnbanIP = 76               '/UNBANIP
    CreateItem = 77            '/CI
    DestroyItems = 78          '/DEST
    ChaosLegionKick = 79       '/NOCAOS
    RoyalArmyKick = 80         '/NOREAL
    ForceMIDIAll = 81          '/FORCEMIDI
    ForceWAVEAll = 82          '/FORCEWAV
    RemovePunishment = 83      '/BORRARPENA
    TileBlockedToggle = 84     '/BLOQ
    KillNPCNoRespawn = 85      '/MATA
    KillAllNearbyNPCs = 86     '/MASSKILL
    lastip = 87                '/LASTIP
    SystemMessage = 88         '/SMSG
    CreateNPC = 89             '/ACC
    CreateNPCWithRespawn = 90  '/RACC
    ImperialArmour = 91        '/AI1 - 4
    ChaosArmour = 92           '/AC1 - 4
    NavigateToggle = 93        '/NAVE
    ServerOpenToUsersToggle = 94 '/HABILITAR
    TurnOffServer = 95         '/APAGAR
    TurnCriminal = 96          '/CONDEN
    ResetFactionCaos = 97        '/RAJAR
    ResetFactionReal = 98         '/RAJAR
    RemoveCharFromGuild = 99   '/RAJARCLAN
    RequestCharMail = 100       '/LASTEMAIL
    AlterPassword = 101         '/APASS
    AlterMail = 102             '/AEMAIL
    AlterName = 103             '/ANAME
    ToggleCentinelActivated = 104 '/CENTINELAACTIVADO
    DoBackUp = 105              '/DOBACKUP
    ShowGuildMessages = 106     '/SHOWCMSG
    SaveMap = 107               '/GUARDAMAPA
    ChangeMapInfoPK = 108       '/MODMAPINFO PK
    ChangeMapInfoBackup = 109   '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted = 110 '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic = 111  '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi = 112   '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu = 113   '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand = 114     '/MODMAPINFO TERRENO
    ChangeMapInfoZone = 115     '/MODMAPINFO ZONA
    ChangeMapInfoStealNpc = 116 '/MODMAPINFO ROBONPCm
    ChangeMapInfoNoOcultar = 117 '/MODMAPINFO OCULTARSINEFECTO
    ChangeMapInfoNoInvocar = 118 '/MODMAPINFO INVOCARSINEFECTO
    SaveChars = 119             '/GRABAR
    CleanSOS = 120              '/BORRAR SOS
    ShowServerForm = 121        '/SHOW INT
    night = 122                 '/NOCHE
    KickAllChars = 123          '/ECHARTODOSPJS
    ReloadNPCs = 124            '/RELOADNPCS
    ReloadServerIni = 125       '/RELOADSINI
    ReloadSpells = 126          '/RELOADHECHIZOS
    ReloadObjects = 127         '/RELOADOBJ
    Restart = 128               '/REINICIAR
    ResetAutoUpdate = 129       '/AUTOUPDATE
    ChatColor = 130             '/CHATCOLOR
    Ignored = 131               '/IGNORADO
    UserOro = 132
    UserPlata = 133
    UserBronce = 134
    CheckSlot = 135             '/SLOT
    SetIniVar = 136             '/SETINIVAR LLAVE CLAVE VALOR
    Seguimiento = 137
           '//Disco.
    CheckHD = 138               '/VERHD USUARIO
    BanHD = 139                 '/BANHD USUARIO
    UnBanHD = 140               '/UNBANHD NROHD
    '///Disco.
    MapMessage = 141            '/MAPMSG
    Impersonate = 142           '/IMPERSONAR
    Imitate = 143               '/MIMETIZAR
    CambioPj = 144              '/CAMBIO
    LarryMataNiños = 145
    ComandoPorDias = 146
    DarPoints = 147
    CreateInvasion = 148
    TerminateInvasion = 149
    SearchNpc = 150             'BUSCAR NPC
    SearchObj = 151             'BUSCAR OBJETO
    SearcherShow = 152          'BUSC
End Enum
'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

Public Const MENSAJE_CRIATURA_FALLA_GOLPE As String = _
    "¡¡¡La criatura falló el golpe!!!"
Public Const MENSAJE_CRIATURA_MATADO As String = _
    "¡¡¡La criatura te ha matado!!!"
Public Const MENSAJE_RECHAZO_ATAQUE_ESCUDO As String = _
    "¡¡¡Has rechazado el ataque con el escudo!!!"
Public Const MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO  As String = _
    "¡¡¡El usuario rechazó el ataque con su escudo!!!"
Public Const MENSAJE_FALLADO_GOLPE As String = "¡¡¡Has fallado el golpe!!!"
Public Const MENSAJE_SEGURO_ACTIVADO As String = "SEGURO ACTIVADO"
Public Const MENSAJE_DRAG_DESACTIVADO As String = "DRAG DESACTIVADO"
Public Const MENSAJE_DRAG_ACTIVADO As String = "DRAG ACTIVADO"
Public Const MENSAJE_SEGURO_DESACTIVADO As String = "SEGURO DESACTIVADO"
Public Const MENSAJE_PIERDE_NOBLEZA As String = _
    "¡¡Has perdido puntaje de nobleza y ganado puntaje de criminalidad!! Si sigues ayudando a criminales te convertirás en uno de ellos y serás perseguido por las tropas de las ciudades."
Public Const MENSAJE_USAR_MEDITANDO As String = _
    "¡Estás meditando! Debes dejar de meditar para usar objetos."

'Public Const MENSAJE_SEGURO_RESU_ON As String = "SEGURO DE RESURRECCION ACTIVADO"
'Public Const MENSAJE_SEGURO_RESU_OFF As String = "SEGURO DE RESURRECCION DESACTIVADO"

Public Const MENSAJE_GOLPE_CABEZA As String = _
    "¡¡La criatura te ha pegado en la cabeza por "
Public Const MENSAJE_GOLPE_BRAZO_IZQ As String = _
    "¡¡La criatura te ha pegado el brazo izquierdo por "
Public Const MENSAJE_GOLPE_BRAZO_DER As String = _
    "¡¡La criatura te ha pegado el brazo derecho por "
Public Const MENSAJE_GOLPE_PIERNA_IZQ As String = _
    "¡¡La criatura te ha pegado la pierna izquierda por "
Public Const MENSAJE_GOLPE_PIERNA_DER As String = _
    "¡¡La criatura te ha pegado la pierna derecha por "
Public Const MENSAJE_GOLPE_TORSO  As String = _
    "¡¡La criatura te ha pegado en el torso por "

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_1 As String = "¡¡"
Public Const MENSAJE_2 As String = "!!"
Public Const MENSAJE_11 As String = "¡"
Public Const MENSAJE_22 As String = "!"

Public Const MENSAJE_GOLPE_CRIATURA_1 As String = _
    "¡¡Le has pegado a la criatura por "

Public Const MENSAJE_ATAQUE_FALLO As String = " te atacó y falló!!"

Public Const MENSAJE_RECIVE_IMPACTO_CABEZA As String = _
    " te ha pegado en la cabeza por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ As String = _
    " te ha pegado el brazo izquierdo por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_DER As String = _
    " te ha pegado el brazo derecho por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ As String = _
    " te ha pegado la pierna izquierda por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_DER As String = _
    " te ha pegado la pierna derecha por "
Public Const MENSAJE_RECIVE_IMPACTO_TORSO As String = _
    " te ha pegado en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_1 As String = "¡¡Le has pegado a "
Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA As String = " en la cabeza por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ As String = _
    " en el brazo izquierdo por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER As String = _
    " en el brazo derecho por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ As String = _
    " en la pierna izquierda por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER As String = _
    " en la pierna derecha por "
Public Const MENSAJE_PRODUCE_IMPACTO_TORSO As String = " en el torso por "

Public Const MENSAJE_TRABAJO_MAGIA As String = "Haz click sobre el objetivo..."
Public Const MENSAJE_TRABAJO_PESCA As String = _
    "Haz click sobre el sitio donde quieres pescar..."
Public Const MENSAJE_TRABAJO_ROBAR As String = "Haz click sobre la víctima..."
Public Const MENSAJE_TRABAJO_TALAR As String = "Haz click sobre el árbol..."
Public Const MENSAJE_TRABAJO_MINERIA As String = _
    "Haz click sobre el yacimiento..."
Public Const MENSAJE_TRABAJO_FUNDIRMETAL As String = _
    "Haz click sobre la fragua..."
Public Const MENSAJE_TRABAJO_PROYECTILES As String = _
    "Haz click sobre la victima..."

Public Const MENSAJE_ENTRAR_PARTY_1 As String = _
    "Si deseas entrar en una party con "
Public Const MENSAJE_ENTRAR_PARTY_2 As String = ", escribe /entrarparty"

Public Const MENSAJE_NENE As String = "Cantidad de NPCs: "

Public Const MENSAJE_FRAGSHOOTER_TE_HA_MATADO As String = "te ha matado!"
Public Const MENSAJE_FRAGSHOOTER_HAS_MATADO As String = "Has matado a"
Public Const MENSAJE_FRAGSHOOTER_HAS_GANADO As String = "Has ganado "
Public Const MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA As String = _
    "puntos de experiencia."
Public Const MENSAJE_HAS_MATADO_A As String = "Has matado a "
Public Const MENSAJE_HAS_GANADO_EXPE_1 As String = "Has ganado "
Public Const MENSAJE_HAS_GANADO_EXPE_2 As String = " puntos de experiencia."
Public Const MENSAJE_TE_HA_MATADO As String = " te ha matado!"

Public Const MENSAJE_HOGAR As String = _
    "Has llegado a tu hogar. El viaje ha finalizado."
Public Const MENSAJE_HOGAR_CANCEL As String = "Tu viaje ha sido cancelado."

Public Enum eMessages
    DontSeeAnything
    NPCSwing
    NPCKillUser
    BlockedWithShieldUser
    BlockedWithShieldOther
    UserSwing
    SafeModeOn
    SafeModeOff
    DragOn
    DragOff
    ResuscitationSafeOff
    ResuscitationSafeOn
    NobilityLost
    CantUseWhileMeditating
    NPCHitUser
    UserHitNPC
    UserAttackedSwing
    UserHittedByUser
    UserHittedUser
    WorkRequestTarget
    HaveKilledUser
    UserKill
    EarnExp
    GoHome
    CancelGoHome
    FinishHome
End Enum

'Inventario
Type Inventory
    ObjIndex As Integer
    Name As String
    GrhIndex As Integer
    '[Alejo]: tipo de datos ahora es Long
    Amount As Long
    '[/Alejo]
    Equipped As Byte
    Valor As Single
    ObjType As Integer
    MaxDef As Integer
    MinDef As Integer 'Budi
    MaxHit As Integer
    MinHit As Integer
End Type

Type NpCinV
    ObjIndex As Integer
    Name As String
    GrhIndex As Integer
    Amount As Integer
    Valor As Single
    Copas As Integer
    Eldhir As Integer
    ObjType As Integer
    MaxDef As Integer
    MinDef As Integer
    MaxHit As Integer
    MinHit As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
End Type

Type tReputacion 'Fama del usuario
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    
    Promedio As Long
End Type

Type tEstadisticasUsu
    CiudadanosMatados As Long
    CriminalesMatados As Long
    UsuariosMatados As Long
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long
End Type

Type tItemsConstruibles
    Name As String
    ObjIndex As Integer
    GrhIndex As Integer
    LinH As Integer
    LinP As Integer
    LinO As Integer
    Madera As Integer
    MaderaElfica As Integer
    Upgrade As Integer
    UpgradeName As String
    UpgradeGrhIndex As Integer
End Type

Public Nombres As Boolean

'User status vars
Global OtroInventario(1 To MAX_INVENTORY_SLOTS) As Inventory

Public UserHechizos(1 To MAXHECHI) As Integer

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV
Public UserMeditar As Boolean
Public UserName As String
Public UserPassword As String
Public UserMaxHP As Integer
Public UserMinHP As Integer
Public UserMaxMAN As Integer
Public UserMinMAN As Integer
Public UserMaxSTA As Integer
Public UserMinSTA As Integer
Public UserMaxAGU As Byte
Public UserMinAGU As Byte
Public UserMaxHAM As Byte
Public UserMinHAM As Byte
Public UserGLD As Long
Public UserLvl As Integer
Public UserPort As Integer
Public UserServerIP As String
Public UserEstado As Byte '0 = Vivo & 1 = Muerto
Public UserPasarNivel As Long
Public UserExp As Long
Public UserOcu As Byte
Public UserReputacion As tReputacion
Public UserEstadisticas As tEstadisticasUsu
Public UserDescansar As Boolean
Public tipf As String
Public PrimeraVez As Boolean
Public bShowTutorial As Boolean
Public FPSFLAG As Boolean
Public pausa As Boolean
Public Iscombate As Boolean
Public UserParalizado As Boolean
Public UserNavegando As Boolean
Public UserHogar As eCiudad
Public UserMontando As Boolean
Public Istrabajando As Boolean

Public UserFuerza As Byte
Public UserAgilidad As Byte

Public UserWeaponEqpSlot As Byte
Public UserArmourEqpSlot As Byte
Public UserHelmEqpSlot As Byte
Public UserShieldEqpSlot As Byte

'<-------------------------NUEVO-------------------------->
Public Canjeando As Boolean
Public Comerciando As Boolean
Public MirandoForo As Boolean
Public MirandoAsignarSkills As Boolean
Public MirandoEstadisticas As Boolean
Public MirandoParty As Boolean
'<-------------------------NUEVO-------------------------->

Public UserClase As eClass
Public UserSexo As eGenero
Public UserRaza As eRaza
Public UserEmail As String

Public Const NUMCIUDADES As Byte = 5
Public Const NUMSKILLS As Byte = 22
Public Const NUMATRIBUTOS As Byte = 5
Public Const NUMCLASES As Byte = 11
Public Const NUMRAZAS As Byte = 5

Public UserSkills(1 To NUMSKILLS) As Byte
Public PorcentajeSkills(1 To NUMSKILLS) As Byte
Public SkillsNames(1 To NUMSKILLS) As String

Public UserAtributos(1 To NUMATRIBUTOS) As Byte
Public AtributosNames(1 To NUMATRIBUTOS) As String

Public Ciudades(1 To NUMCIUDADES) As String

Public ListaRazas(1 To NUMRAZAS) As String
Public ListaClases(1 To NUMCLASES) As String

Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer
Public Oscuridad As Integer

Public UsingSkill As Integer

Public MD5HushYo As String

Public pingTime As Long
Public TimePing As Long

Public EsPartyLeader As Boolean

' Sistema de cuentas
Public Const MAX_PJS_ACCOUNT = 8


Public Type tCuentaUser
    Name As String
    Ban As Byte
    Clase As eClass
    Raza As eRaza
    Elv As Byte
End Type

Private Type tCuenta
    Account As String
    Email As String
    PIN As String
    Passwd As String
    NewPasswd As String
End Type

Public Cuenta As tCuenta
Public CuentaChars(1 To MAX_PJS_ACCOUNT) As tCuentaUser

Public Enum E_MODO
    Normal = 1
    CrearNuevoPj = 2
    Dados = 3
    BorrarPJ = 4
    RecuperarPJ = 5
    
    e_NewAccount = 10
    e_ConnectAccount = 20
    e_LoginCharAccount = 30
    e_CreateCharAccount = 40
    e_RemoveCharAccount = 50
    e_RecoverAccount = 60
    e_ChangePasswdAccount = 70
    e_Temporal = 80
End Enum

Public EstadoLogin As E_MODO
   
Public Enum FxMeditar
    CHICO = 4
    MEDIANO = 5
    GRANDE = 6
    XGRANDE = 16
    XXGRANDE = 34
End Enum

Public Enum eClanType
    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal
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
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param trigger_2 ???
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'
Public Enum eTrigger
    NADA = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public stxtbuffergmsg As String 'Holds temp raw data from server
Public stxtbufferrmsg As String 'Holds temp raw data from server
Public SendNewChar As Boolean 'Used during login
Public Connected As Boolean 'True when connected to server
Public DownloadingMap As Boolean 'Currently downloading a map from server
Public UserMap As Integer
Public MapaActual As String
'Control
Public prgRun As Boolean 'When true the program ends

Public IPdelServidor As String
Public PuertoDelServidor As String

'
'********** FUNCIONES API ***********
'

Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal _
    lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname _
    As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal _
    nSize As Long, ByVal lpFileName As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As _
    Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) _
    As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Para ejecutar el browser y programas externos
Public Const SW_SHOWNORMAL As Long = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal _
    lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As _
    Long

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type

Public EsperandoLevel As Boolean

' Tipos de mensajes
Public Enum eForumMsgType
    ieGeneral
    ieGENERAL_STICKY
    ieREAL
    ieREAL_STICKY
    ieCAOS
    ieCAOS_STICKY
End Enum

' Indica los privilegios para visualizar los diferentes foros
Public Enum eForumVisibility
    ieGENERAL_MEMBER = &H1
    ieREAL_MEMBER = &H2
    ieCAOS_MEMBER = &H4
End Enum

' Indica el tipo de foro
Public Enum eForumType
    ieGeneral
    ieREAL
    ieCAOS
End Enum

' Limite de posts
Public Const MAX_STICKY_POST As Byte = 5
Public Const MAX_GENERAL_POST As Byte = 30
Public Const STICKY_FORUM_OFFSET As Byte = 50

' Estructura contenedora de mensajes
Public Type tForo
    StickyTitle(1 To MAX_STICKY_POST) As String
    StickyPost(1 To MAX_STICKY_POST) As String
    StickyAuthor(1 To MAX_STICKY_POST) As String
    GeneralTitle(1 To MAX_GENERAL_POST) As String
    GeneralPost(1 To MAX_GENERAL_POST) As String
    GeneralAuthor(1 To MAX_GENERAL_POST) As String
End Type

' 1 foro general y 2 faccionarios
Public Foros(0 To 2) As tForo

' Forum info handler
Public clsForos As New clsForum

Public isCapturePending As Boolean
Public Traveling As Boolean

Public GuildNames() As String
Public GuildMembers() As String

Public Const OFFSET_HEAD As Integer = -34

Public Enum eSMType
    sResucitation
    sSafemode
    DragMode
    mSpells
    mWork
End Enum

Public Const SM_CANT As Byte = 4
Public SMStatus(SM_CANT) As Boolean

'Hardcoded grhs and items
Public Const GRH_INI_SM As Integer = 4978

Public Const ORO_INDEX As Integer = 12
Public Const ORO_GRH As Integer = 511

Public Const GRH_HALF_STAR As Integer = 5357
Public Const GRH_FULL_STAR As Integer = 5358
Public Const GRH_GLOW_STAR As Integer = 5359

Public Const LH_GRH As Integer = 724
Public Const LP_GRH As Integer = 725
Public Const LO_GRH As Integer = 723

Public Const MADERA_GRH As Integer = 550
Public Const MADERA_ELFICA_GRH As Integer = 1999

Public picMouseIcon As Picture
