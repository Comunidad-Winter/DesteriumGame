Attribute VB_Name = "Declaraciones"
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

Public ArbT As Long

Public RetosActivos As Boolean
Public EventosActivos As Boolean

Public Declare Sub GetSafeArrayPointer Lib "msvbvm60.dll" Alias "GetMem4" (pArray() As Any, ret As Long)

Public CountDownLimpieza As Long

Public Enum eMenues
    ieUser = 1
    ieNpcComercio = 2
    ieNpcNoHostil = 3
    
End Enum

Public INTERVALO_AUTO_GP As Byte

'Sistema de CuentaRegresiva
Public CuentaRegresivaTimer As Byte

Public Const INTERVALO_GLOBAL As Integer = 10 ' 5 SEGUNDOS, CAMBIALO.

'Sistema Global
Public GlobalActivado As Byte

'********** Constantes de daño.
Public Const DAMAGE_PUÑAL    As Byte = 1
Public Const DAMAGE_NORMAL   As Byte = 2

'********** Constantes de daño.

Public PuedeAbandonar As Boolean
   Public Type Retos
    Retador As Integer
    Oponente As Integer
    End Type
    Public Retoz As Retos
    Public LReto As Byte
Public RetoxItems As Integer

 Public PosMapUsersend As Integer
 Public PosXUsersend As Byte
 Public PosYUsersend As Byte
 Public PosMapUsersend1 As Integer
 Public PosXUsersend1 As Byte
 Public PosYUsersend1 As Byte

#If SeguridadAlkon Then
Public aDos As New ClsAntiDos
#End If

Public aClon As New clsAntiMassClon
Public TrashCollector As New Collection

Public aDos As New ClsAntiDos

Public Const MAXSPAWNATTEMPS = 60
Public Const INFINITE_LOOPS As Integer = -1
Public Const FXSANGRE = 14

''
' The color of chats over head of dead characters.
Public Const CHAT_COLOR_DEAD_CHAR As Long = &HC0C0C0

''
' The color of yells made by any kind of game administrator.
Public Const CHAT_COLOR_GM_YELL As Long = &HF82FF

' Coordinates for normal sounds (not 3D, like rain)
Public Const NO_3D_SOUND As Byte = 0

Public Const iFragataFantasmal = 87
Public Const iFragataReal = 87
Public Const iFragataCaos = 87
Public Const iBarca = 84
Public Const iGalera = 85
Public Const iGaleon = 86
Public Const iBarcaCiuda = 84
Public Const iBarcaPk = 84
Public Const iGaleraCiuda = 85
Public Const iGaleraPk = 85
Public Const iGaleonCiuda = 86
Public Const iGaleonPk = 86

Public Enum iMinerales
    HierroCrudo = 192
    PlataCruda = 193
    OroCrudo = 194
    LingoteDeHierro = 386
    LingoteDePlata = 387
    LingoteDeOro = 388
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

Public Enum eClass
    Mage = 1       'Mago
    Cleric      'Clérigo
    Warrior     'Guerrero
    Assasin     'Asesino
    Thief       'Ladrón
    Bard        'Bardo
    Druid       'Druida
    'Bandit      'Bandido
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
    cArkhein
    cLastCity
End Enum

Public Enum eRaza
    Humano = 1
    Elfo
    Drow
    Gnomo
    Enano
End Enum

Enum eGenero
    Hombre = 1
    Mujer
End Enum

Public Enum eClanType
    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal
End Enum

Public Const LimiteNewbie As Byte = 12

Public Type tCabecera 'Cabecera de los con
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera


Public Const NingunEscudo As Integer = 2
Public Const NingunCasco As Integer = 2
Public Const NingunArma As Integer = 2

Public Const EspadaMataDragonesIndex As Integer = 1037
Public Const LAUDMAGICO As Integer = 696
Public Const FLAUTAMAGICA As Integer = 208

Public Const LAUDSUPERMAGICO As Integer = 921
Public Const FLAUTAANTIGUA As Integer = 922


Public Const LAUDELFICO As Integer = 643
Public Const FLAUTAELFICA As Integer = 648

Public Const LaudBronce As Integer = 1029
Public Const AnilloBronce As Integer = 1030

Public Const LaudPlata As Integer = 1031
Public Const AnilloPlata As Integer = 1032

Public Const APOCALIPSIS_SPELL_INDEX As Integer = 25
Public Const DESCARGA_SPELL_INDEX As Integer = 23

Public Const SLOTS_POR_FILA As Byte = 5

Public Const PROB_ACUCHILLAR As Byte = 20
Public Const DAÑO_ACUCHILLAR As Single = 0.2

Public Const MAXMASCOTASENTRENADOR As Byte = 7

Public Enum FXIDs
    FXWARP = 1
    FXMEDITARCHICO = 4
    FXMEDITARMEDIANO = 5
    FXMEDITARGRANDE = 6
    FXMEDITARXGRANDE = 16
    FXMEDITARXXGRANDE = 16
End Enum

Public Const TIEMPO_CARCEL_PIQUETE As Long = 10

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
    zonaOscura = 7
    CASA = 8
End Enum

''
' constantes para el trigger 6
'
' @see eTrigger
' @param TRIGGER6_PERMITE TRIGGER6_PERMITE
' @param TRIGGER6_PROHIBE TRIGGER6_PROHIBE
' @param TRIGGER6_AUSENTE El trigger no aparece
'
Public Enum eTrigger6
    TRIGGER6_PERMITE = 1
    TRIGGER6_PROHIBE = 2
    TRIGGER6_AUSENTE = 3
End Enum

'TODO : Reemplazar por un enum
Public Const Bosque As String = "BOSQUE"
Public Const Nieve As String = "NIEVE"
Public Const Desierto As String = "DESIERTO"
Public Const Ciudad As String = "CIUDAD"
Public Const Campo As String = "CAMPO"
Public Const Dungeon As String = "DUNGEON"

' <<<<<< Targets >>>>>>
Public Enum TargetType
    uUsuarios = 1
    uNPC = 2
    uUsuariosYnpc = 3
    uTerreno = 4
End Enum

' <<<<<< Acciona sobre >>>>>>
Public Enum TipoHechizo
    uPropiedades = 1
    uEstado = 2
    uMaterializa = 3    'Nose usa
    uInvocacion = 4
End Enum

Public Const MAXUSERHECHIZOS As Byte = 35


' TODO: Y ESTO ? LO CONOCE GD ?
Public Const EsfuerzoTalarGeneral As Byte = 4
Public Const EsfuerzoTalarLeñador As Byte = 2

Public Const EsfuerzoPescarPescador As Byte = 1
Public Const EsfuerzoPescarGeneral As Byte = 3

Public Const EsfuerzoExcavarMinero As Byte = 2
Public Const EsfuerzoExcavarGeneral As Byte = 5

Public Const FX_TELEPORT_INDEX As Integer = 1

Public Const PORCENTAJE_MATERIALES_UPGRADE As Single = 0.85

' La utilidad de esto es casi nula, sólo se revisa si fue a la cabeza...
Public Enum PartesCuerpo
    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6
End Enum

Public Const Guardias As Integer = 6

Public Const MAX_ORO_EDIT As Long = 5000000
Public Const MAX_VIDA_EDIT As Long = 30000


Public Const STANDARD_BOUNTY_HUNTER_MESSAGE As String = "Se te ha otorgado un premio por ayudar al proyecto reportando bugs, el mismo está disponible en tu bóveda."
Public Const TAG_USER_INVISIBLE As String = "[INVISIBLE]"
Public Const TAG_CONSULT_MODE As String = "[CONSULTA]"


Public Const MAXREP As Long = 6000000
Public Const MaxOro As Long = 200000000
Public Const MAXEXP As Long = 1670680720


Public Const MAXUSERMATADOS As Long = 65000

Public Const MAXATRIBUTOS As Byte = 38
Public Const MINATRIBUTOS As Byte = 6

Public Const LingoteHierro As Integer = 386
Public Const LingotePlata As Integer = 387
Public Const LingoteOro As Integer = 388
Public Const Leña As Integer = 58
Public Const LeñaElfica As Integer = 1006

Public Const MAXNPCS As Integer = 10000
Public Const MAXCHARS As Integer = 10000

Public Const HACHA_LEÑADOR As Byte = 127
Public Const HACHA_DORADA As Integer = 630
Public Const PIQUETE_MINERO As Byte = 187
Public Const PIQUETE_ORO As Integer = 685

Public Const DAGA As Integer = 15
Public Const FOGATA_APAG As Integer = 136
Public Const FOGATA As Integer = 63
Public Const ORO_MINA As Integer = 194
Public Const PLATA_MINA As Integer = 193
Public Const HIERRO_MINA As Integer = 192
Public Const MARTILLO_HERRERO As Integer = 389
Public Const SERRUCHO_CARPINTERO As Integer = 198
Public Const ObjArboles As Integer = 4
Public Const RED_PESCA As Integer = 543
Public Const CAÑA_PESCA As Integer = 138
Public Const CAÑA_COFRES As Integer = 1340

Public Enum eNPCType
    Comun = 0
    Revividor = 1
    GuardiaReal = 2
    Entrenador = 3
    Banquero = 4
    Noble = 5
    Dragon = 6
    Timbero = 7
    Guardiascaos = 8
    ResucitadorNewbie = 9
    pretoriano = 10
    Gobernador = 11
    Pirata = 12
    PirataViajes = 13
    NpcCanjes = 14
End Enum

Public Const MIN_APUÑALAR As Byte = 10

'********** CONSTANTANTES ***********

''
' Cantidad de skills
Public Const NUMSKILLS As Byte = 22

''
' Cantidad de Atributos
Public Const NUMATRIBUTOS As Byte = 5

''
' Cantidad de Clases
Public Const NUMCLASES As Byte = 11

''
' Cantidad de Razas
Public Const NUMRAZAS As Byte = 5


''
' Valor maximo de cada skill
Public Const MAXSKILLPOINTS As Byte = 100


''
' Cantidad de Ciudades
Public Const NUMCIUDADES As Byte = 5


''
'Direccion
'
' @param NORTH Norte
' @param EAST Este
' @param SOUTH Sur
' @param WEST Oeste
'
Public Enum eHeading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

''
' Cantidad maxima de mascotas
Public Const MAXMASCOTAS As Byte = 3

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const vlASALTO As Integer = 100
Public Const vlASESINO As Integer = 1000
Public Const vlCAZADOR As Integer = 5
Public Const vlNoble As Integer = 5
Public Const vlLadron As Integer = 25
Public Const vlProleta As Integer = 2

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const iCuerpoMuerto As Integer = 8
Public Const iCabezaMuerto As Integer = 500

'%%%%%%%%%% Cabeza & Cuerpo Caós %%%%%%%%%%%%%%%%
Public Const iCuerpoMuertoCrimi = 145
Public Const iCabezaMuertoCrimi = 501

Public Const iORO As Byte = 12
Public Const Pescado As Byte = 139

Public Enum PECES_POSIBLES
    PESCADO1 = 139
    PESCADO2 = 544
    PESCADO3 = 545
    PESCADO4 = 546
    PESCADO5 = 732
End Enum

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Enum eSkill
    Magia = 1
    Robar = 2
    Tacticas = 3
    Armas = 4
    Meditar = 5
    Apuñalar = 6
    Ocultarse = 7
    Supervivencia = 8
    talar = 9
    Comerciar = 10
    Defensa = 11
    Pesca = 12
    Mineria = 13
    Carpinteria = 14
    herreria = 15
    Liderazgo = 16
    Domar = 17
    Proyectiles = 18
    Wrestling = 19
    Navegacion = 20
    Equitacion = 21
    Resistencia = 22
End Enum

Public Enum eMochilas
    Mediana = 1
    Grande = 2
End Enum

Public Const FundirMetal = 88

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Public Const AdicionalHPGuerrero As Byte = 2 'HP adicionales cuando sube de nivel
Public Const AdicionalHPCazador As Byte = 1 'HP adicionales cuando sube de nivel

Public Const AumentoSTDef As Byte = 15
Public Const AumentoStBandido As Byte = AumentoSTDef + 23
Public Const AumentoSTLadron As Byte = AumentoSTDef + 3
Public Const AumentoSTMago As Byte = AumentoSTDef - 1
Public Const AumentoSTTrabajador As Byte = AumentoSTDef + 25

'Tamaño del mapa
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

'Tamaño del tileset
Public Const TileSizeX As Byte = 32
Public Const TileSizeY As Byte = 32

'Tamaño en Tiles de la pantalla de visualizacion
Public Const XWindow As Byte = 17
Public Const YWindow As Byte = 13

'Sonidos
Public Const SND_SWING As Byte = 2
Public Const SND_TALAR As Byte = 13
Public Const SND_PESCAR As Byte = 14
Public Const SND_MINERO As Byte = 15
Public Const SND_WARP As Byte = 3
Public Const SND_PUERTA As Byte = 5
Public Const SND_NIVEL As Byte = 6

Public Const SND_USERMUERTE As Byte = 11
Public Const SND_IMPACTO As Byte = 10
Public Const SND_IMPACTO2 As Byte = 12
Public Const SND_LEÑADOR As Byte = 13
Public Const SND_FOGATA As Byte = 14
Public Const SND_AVE As Byte = 21
Public Const SND_AVE2 As Byte = 22
Public Const SND_AVE3 As Byte = 34
Public Const SND_GRILLO As Byte = 28
Public Const SND_GRILLO2 As Byte = 29
Public Const SND_SACARARMA As Byte = 25
Public Const SND_ESCUDO As Byte = 37
Public Const MARTILLOHERRERO As Byte = 41
Public Const LABUROCARPINTERO As Byte = 42
Public Const SND_BEBER As Byte = 46
Public Const SND_SACARMONTURA As Byte = 56
Public Const SND_SACARMONTURADRACO As Byte = 30

''
' Cantidad maxima de objetos por slot de inventario
Public Const MAX_INVENTORY_OBJS As Integer = 10000

' Cantidad de "slots" en el inventario con mochila
Public Const MAX_INVENTORY_SLOTS As Byte = 25

''
' Cantidad de "slots" en el inventario sin mochila
Public Const MAX_NORMAL_INVENTORY_SLOTS As Byte = 25

''
' Constante para indicar que se esta usando ORO
Public Const FLAGORO As Integer = MAX_NORMAL_INVENTORY_SLOTS + 1

' CATEGORIAS PRINCIPALES
Public Enum eOBJType
    otUseOnce = 1
    otWeapon = 2
    otarmadura = 3
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
    otCofre = 54
    otGemaTelep = 55
    otCualquiera = 1000
End Enum

'Texto
Public Const FONTTYPE_TALK As String = "~255~255~255~0~0"
Public Const FONTTYPE_FIGHT As String = "~255~0~0~1~0"
Public Const FONTTYPE_WARNING As String = "~32~51~223~1~1"
Public Const FONTTYPE_INFO As String = "~65~190~156~0~0"
Public Const FONTTYPE_INFOBOLD As String = "~65~190~156~1~0"
Public Const FONTTYPE_EJECUCION As String = "~130~130~130~1~0"
Public Const FONTTYPE_PARTY As String = "~255~180~255~0~0"
Public Const FONTTYPE_VENENO As String = "~0~255~0~0~0"
Public Const FONTTYPE_GUILD As String = "~255~255~255~1~0"
Public Const FONTTYPE_SERVER As String = "~0~185~0~0~0"
Public Const FONTTYPE_GUILDMSG As String = "~228~199~27~0~0"
Public Const FONTTYPE_CONSEJO As String = "~130~130~255~1~0"
Public Const FONTTYPE_CONSEJOCAOS As String = "~255~60~00~1~0"
Public Const FONTTYPE_CONSEJOVesA As String = "~0~200~255~1~0"
Public Const FONTTYPE_CONSEJOCAOSVesA As String = "~255~50~0~1~0"
Public Const FONTTYPE_CENTINELA As String = "~0~255~0~1~0"
Public Const FONTTYPE_ADMIN As String = "~255~166~0~1~0"
Public Const FONTTYPE_PREMIUM As String = "~166~255~158~1~0"
Public Const FONTTYPE_RETO As String = "~255~251~206~0~0"
Public Const FONTTYPE_ORO As String = "~255~222~0~1~0"
Public Const FONTTYPE_PLATA As String = "~130~130~130~1~0"
Public Const FONTTYPE_BRONCE As String = "~255~166~0~1~0"

'Estadisticas
Public Const STAT_MAXELV As Byte = 47
Public Const STAT_MAXHP As Integer = 25000
Public Const STAT_MAXSTA As Integer = 25000
Public Const STAT_MAXMAN As Integer = 25000
Public Const STAT_MAXHIT_UNDER36 As Byte = 99
Public Const STAT_MAXHIT_OVER36 As Integer = 999
Public Const STAT_MAXDEF As Byte = 99

Public Const ELU_SKILL_INICIAL As Byte = 250
Public Const EXP_ACIERTO_SKILL As Byte = 50
Public Const EXP_FALLO_SKILL As Byte = 25

' **************************************************************
' **************************************************************
' ************************ TIPOS *******************************
' **************************************************************
' **************************************************************

Public Type tHechizo
    Nombre As String
    desc As String
    PalabrasMagicas As String
    
    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String
    
'    Resis As Byte
    
    tipo As TipoHechizo
    
    WAV As Integer
    FXgrh As Integer
    loops As Byte
    
    SubeHP As Byte
    MinHp As Integer
    MaxHp As Integer
    
    SubeMana As Byte
    MiMana As Integer
    MaMana As Integer
    
    SubeSta As Byte
    MinSta As Integer
    MaxSta As Integer
    
    SubeHam As Byte
    MinHam As Integer
    MaxHam As Integer
    
    SubeSed As Byte
    MinSed As Integer
    MaxSed As Integer
    
    SubeAgilidad As Byte
    MinAgilidad As Integer
    MaxAgilidad As Integer
    
    SubeFuerza As Byte
    MinFuerza As Integer
    MaxFuerza As Integer
    
    SubeCarisma As Byte
    MinCarisma As Integer
    MaxCarisma As Integer
    
    Invisibilidad As Byte
    Paraliza As Byte
    Inmoviliza As Byte
    RemoverParalisis As Byte
    RemoverEstupidez As Byte
    CuraVeneno As Byte
    Envenena As Byte
    Maldicion As Byte
    RemoverMaldicion As Byte
    Bendicion As Byte
    Estupidez As Byte
    Ceguera As Byte
    Revivir As Byte
    Morph As Byte
    Mimetiza As Byte
    RemueveInvisibilidadParcial As Byte
    
    Warp As Byte
    Invoca As Byte
    NumNpc As Integer
    cant As Integer

'    Materializa As Byte
'    ItemIndex As Byte
    
    MinSkill As Integer
    ManaRequerido As Integer

    'Barrin 29/9/03
    StaRequerido As Integer

    Target As TargetType
    
    NeedStaff As Integer
    StaffAffected As Boolean
End Type

Public Type LevelSkill
    LevelValue As Integer
End Type

Public Type UserObj
    ObjIndex As Integer
    Amount As Integer
    Equipped As Byte
End Type

Public Type Inventario
    Object(1 To MAX_INVENTORY_SLOTS) As UserObj
    WeaponEqpObjIndex As Integer
    WeaponEqpSlot As Byte
    ArmourEqpObjIndex As Integer
    ArmourEqpSlot As Byte
    EscudoEqpObjIndex As Integer
    EscudoEqpSlot As Byte
    CascoEqpObjIndex As Integer
    CascoEqpSlot As Byte
    MunicionEqpObjIndex As Integer
    MunicionEqpSlot As Byte
    AnilloEqpObjIndex As Integer
    AnilloEqpSlot As Byte
    BarcoObjIndex As Integer
    BarcoSlot As Byte
    MonturaObjIndex As Integer
    MonturaSlot As Byte
    MochilaEqpObjIndex As Integer
    MochilaEqpSlot As Byte
    AnilloNpcSlot As Byte
    AnilloNpcObjIndex As Integer
    NroItems As Integer
End Type

Public Type Position
    X As Integer
    Y As Integer
End Type

Public Type WorldPos
    map As Integer
    X As Integer
    Y As Integer
End Type

Public Type FXdata
    Nombre As String
    GrhIndex As Integer
    Delay As Integer
End Type

'Datos de user o npc
Public Type Char
    CharIndex As Integer
    Head As Integer
    body As Integer
    
    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    
    FX As Integer
    loops As Integer
    
    Heading As eHeading
End Type

'Tipos de objetos
Public Type ObjData
    NpcTipo As Integer
    Velocidad As Byte 'Velocidad de Monturas
    Name As String 'Nombre del obj
    
    ObjType As eOBJType 'Tipo enum que determina cuales son las caract del obj
    
    GrhIndex As Integer ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    
    'Solo contenedores
    MaxItems As Integer
    Conte As Inventario
    Apuñala As Byte
    Acuchilla As Byte
    
    HechizoIndex As Integer
    
    ForoID As String
    
    MinHp As Integer ' Minimo puntos de vida
    MaxHp As Integer ' Maximo puntos de vida
    
    
        MagiaSkill As Byte
    RMSkill As Byte
    ArmaSkill As Byte
    EscudoSkill As Byte
    ArmaduraSkill As Byte
    ArcoSkill As Byte
    DagaSkill As Byte
    Monturasskill As Byte
    MonturasDracoskill As Byte
    MonturasDinoskill As Byte
    QuitaEnergia As Integer
    
    MineralIndex As Integer
    LingoteInex As Integer
    
    
    proyectil As Integer
    Municion As Integer
    
    Crucial As Byte
    Newbie As Integer
    Quince As Integer
    Treinta As Integer
    HM As Integer
    UM As Integer
    MM As Integer
    
    'Puntos de Stamina que da
    MinSta As Integer ' Minimo puntos de stamina
    
    ' Telep gemas nuevas
    TelepMap As Integer
    TelepX As Byte
    TelepY As Byte
    TelepTime As Integer
    
    
    ''
    
    'Pociones
    TipoPocion As Byte
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long
    MinSkill As Integer
    LingoteIndex As Integer
    
    MinHIT As Integer 'Minimo golpe
    MaxHIT As Integer 'Maximo golpe
    
    MinHam As Integer
    MinSed As Integer
    
    def As Integer
    MinDef As Integer ' Armaduras
    MaxDef As Integer ' Armaduras
    
    Ropaje As Integer 'Indice del grafico del ropaje
    
    WeaponAnim As Integer ' Apunta a una anim de armas
    WeaponRazaEnanaAnim As Integer
    ShieldAnim As Integer ' Apunta a una anim de escudo
    CascoAnim As Integer
    
    valor As Long     ' Precio
    Premium As Byte
    copaS As Integer     ' Precio en copas
    Eldhir As Integer
    
    Cerrada As Integer
    Llave As Byte
    Clave As Long 'si clave=llave la puerta se abre o cierra
    
    Radio As Integer ' Para teleps: El radio para calcular el random de la pos destino
    
    MochilaType As Byte 'Tipo de Mochila (1 la chica, 2 la grande)
    
    Guante As Byte ' Indica si es un guante o no.
    
    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer
    
    RazaEnana As Byte
    RazaDrow As Byte
    RazaElfa As Byte
    RazaGnoma As Byte
    RazaHumana As Byte
    
    Mujer As Byte
    Hombre As Byte
    
    Envenena As Byte
    Paraliza As Byte
    
    Agarrable As Byte
    
    LingH As Integer
    LingO As Integer
    LingP As Integer
    Madera As Integer
    MaderaElfica As Integer
    
    SkHerreria As Integer
    SkCarpinteria As Integer
    
    Texto As String
    
    'Clases que no tienen permitido usar este obj
    ClaseProhibida(1 To NUMCLASES) As eClass
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    Real As Integer
    Caos As Integer
    
    NoSeCae As Integer
        VIP As Byte
        VIPB As Byte
        VIPP As Byte
    DiosTerrenal As Byte
    
    StaffPower As Integer
    StaffDamageBonus As Integer
    DefensaMagicaMax As Integer
    DefensaMagicaMin As Integer
    Refuerzo As Byte
    
    LOG As Byte 'es un objeto que queremos loguear? Pablo (ToxicWaste) 07/09/07
    NoLog As Byte 'es un objeto que esta prohibido loguear?
    
    Upgrade As Integer
End Type

Public Type Obj
    ObjIndex As Integer
    Amount As Integer
End Type


Public MAX_NPCS_CANJE As Byte
Public Canjes() As tCanjes
Public NumCanjes As Byte

Public Type tCanjes
    NumRequired As Byte
    
    ObjRequired() As Obj
    Npcs As Integer
    ObjCanje As Obj
    Points As Long
End Type

Public Type tQuestNpc
    NpcIndex As Integer
    Amount As Integer
End Type
 
Public Type tUserQuest
    NPCsKilled() As Integer
    QuestIndex As Integer
End Type
 
Public Type tQuestStats
    Quests(1 To MAXUSERQUESTS) As tUserQuest
    NumQuestsDone As Integer
    QuestsDone() As Integer
End Type

'[Pablo ToxicWaste]
Public Type ModClase
    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    AtaqueWrestling As Double
    DañoArmas As Double
    DañoProyectiles As Double
    DañoWrestling As Double
    Escudo As Double
End Type

Public Type ModRaza
    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Carisma As Single
    Constitucion As Single
End Type
'[/Pablo ToxicWaste]

'[KEVIN]
'Banco Objs
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
'[/KEVIN]

'[KEVIN]
Public Type BancoInventario
    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserObj
    NroItems As Integer
End Type
'[/KEVIN]

' Determina el color del nick
Public Enum eNickColor
    ieCriminal = &H1
    ieCiudadano = &H2
    ieAtacable = &H4
    ieTeamUno = &H8
    ieTeamDos = &H10
    
End Enum

'*******
'FOROS *
'*******

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
Public Const MAX_STICKY_POST As Byte = 10
Public Const MAX_GENERAL_POST As Byte = 35

' Estructura contenedora de mensajes
Public Type tForo
    StickyTitle(1 To MAX_STICKY_POST) As String
    StickyPost(1 To MAX_STICKY_POST) As String
    GeneralTitle(1 To MAX_GENERAL_POST) As String
    GeneralPost(1 To MAX_GENERAL_POST) As String
End Type

Public Type tQuest
    Nombre As String
    desc As String
    RequiredLevel As Byte
   
    RequiredOBJs As Byte
    RequiredOBJ() As Obj
   
    RequiredNPCs As Byte
    RequiredNPC() As tQuestNpc
   
    RewardGLD As Long
    RewardEXP As Long
   
    RewardOBJs As Byte
    RewardOBJ() As Obj
End Type

'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'******* T I P O S   D E    U S U A R I O S **************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Public Type tReputacion 'Fama del usuario
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    Promedio As Long
End Type

'Estadisticas de los usuarios
Public Type UserStats
    OldHp As Integer
    ViejaHP As Integer
    ViejaMan As Integer
    ViejaminHP As Integer
    ViejaminMan As Integer
    Gld As Long 'Dinero
    Banco As Long
    
    MaxHp As Integer
    MinHp As Integer
    
    MaxSta As Integer
    MinSta As Integer
    MaxMAN As Integer
    MinMAN As Integer
    MaxHIT As Integer
    MinHIT As Integer
    
    MaxHam As Integer
    MinHam As Integer
    
    MaxAGU As Integer
    MinAGU As Integer
        
    def As Integer
    Exp As Double
    ELV As Byte
    ELU As Long
    UserSkills(1 To NUMSKILLS) As Byte
    UserAtributos(1 To NUMATRIBUTOS) As Byte
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Byte
    UserHechizos(1 To MAXUSERHECHIZOS) As Integer
    UsuariosMatados As Long
    CriminalesMatados As Long
    RetosGanados As Integer
    RetosPerdidos As Integer
    OroGanado As Long
    OroPerdido As Long
    TorneosGanados As Integer
    NPCsMuertos As Integer
    SkillPts As Integer
    Points As Long
    
    ExpSkills(1 To NUMSKILLS) As Long
    EluSkills(1 To NUMSKILLS) As Long
    
End Type

'Flags
Public Type UserFlags
    RequestData As String
    TempInt As Integer
    InCVC As Boolean
    IsDios As Boolean
    SlotEvent As Byte
    SlotReto As Byte
    SlotRetoUser As Byte
    SlotUserEvent As Byte
    FightTeam As Byte
    SelectedEvent As Byte
    ElPedidorSeguimiento As Integer
    Siguiendo As Integer
    Infectado As Byte
    Angel As Byte
    Demonio As Byte
    BonosHP As Byte
    Compañero As Integer
    automatico As Boolean
    Plantico As Boolean
     Oro As Byte
     Plata As Byte
     Bronce As Byte
    DiosTerrenal As Byte
    MapU As Integer
    MapX As Integer
    MapY As Integer
    Round As Byte
    EnEvento As Byte
    Peleando As Byte
    BeforeMap As Integer 'Antes mapa
    BeforeX As Integer 'AntesX
    BeforeY As Integer 'AntesY
    Premium As Byte
    EnvioGM As Byte
    Muerto As Byte '¿Esta muerto?
    Escondido As Byte '¿Esta escondido?
    Comerciando As Boolean '¿Esta comerciando?
    UserLogged As Boolean '¿Esta online?
    Meditando As Boolean
    ModoCombate As Boolean
    Descuento As String
    Hambre As Byte
    Sed As Byte
    PuedeMoverse As Byte
    TimerLanzarSpell As Long
    PuedeTrabajar As Byte
    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Estupidez As Byte
    Ceguera As Byte
    invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    Oculto As Byte
    Desnudo As Byte
    Descansar As Boolean
    Hechizo As Integer
    TomoPocion As Boolean
    TipoPocion As Byte
    
    NoPuedeSerAtacado As Boolean
    AtacablePor As Integer
    ShareNpcWith As Integer
    
    Vuela As Byte
    Navegando As Byte
    Montando As Byte
    Seguro As Boolean
    SeguroResu As Boolean
    
    DuracionEfecto As Long
    TargetNPC As Integer ' Npc señalado por el usuario
    TargetNpcTipo As eNPCType ' Tipo del npc señalado
    OwnedNpc As Integer ' Npc que le pertenece (no puede ser atacado)
    NpcInv As Integer
    
    Ban As Byte
    AdministrativeBan As Byte
    
    TargetUser As Integer ' Usuario señalado
    
    TargetObj As Integer ' Obj señalado
    TargetObjMap As Integer
    TargetObjX As Integer
    TargetObjY As Integer
    
    TargetMap As Integer
    TargetX As Integer
    TargetY As Integer
    
    TargetObjInvIndex As Integer
    TargetObjInvSlot As Integer
    
    AtacadoPorNpc As Integer
    AtacadoPorUser As Integer
    NPCAtacado As Integer
    Ignorado As Boolean
    
    EnConsulta As Boolean
    
    StatsChanged As Byte
    Privilegios As PlayerType
    PrivEspecial As Boolean
    
    ValCoDe As Integer
    
    LastCrimMatado As String
    LastCiudMatado As String
    
    OldBody As Integer
    OldHead As Integer
    AdminInvisible As Byte
    AdminPerseguible As Boolean
    
    ChatColor As Long
    
    '[el oso]
    MD5Reportado As String
    '[/el oso]
    
    '[Barrin 30-11-03]
    TimesWalk As Long
    StartWalk As Long
    CountSH As Long
    '[/Barrin 30-11-03]
    
    '[CDT 17-02-04]
    UltimoMensaje As Byte
    '[/CDT]
    
    Silenciado As Byte
    
    Mimetizado As Byte
    
    CentinelaOK As Boolean 'Centinela
    
    lastMap As Integer
    Traveling As Byte 'Travelin Band ¿?
    
        MenuCliente As Byte
    LastSlotClient As Byte

End Type

Public Type UserCounters
    TimeTelep As Long
    TimePotas As Long
    TimePin As Long
    TimeAntiFriz As Long
    TimePotFull As Long
    TimeFight As Long
    TimeCastleMode As Long
    IdleCount As Long
    AttackCounter As Integer
    HPCounter As Integer
    STACounter As Integer
    Frio As Integer
    Lava As Integer
    COMCounter As Integer
    AGUACounter As Integer
    Veneno As Integer
    Paralisis As Integer
    Ceguera As Integer
    Estupidez As Integer
    
    Invisibilidad As Integer
    TiempoOculto As Integer
    Denuncia As Byte
    
    Mimetismo As Integer
    PiqueteC As Long
    Pena As Long
    SendMapCounter As WorldPos
    '[Gonzalo]
    Saliendo As Boolean
    Salir As Integer
    '[/Gonzalo]
    
    TimerLanzarSpell As Long
    TimerPuedeAtacar As Long
    TimerPuedeUsarArco As Long
    TimerPuedeTrabajar As Long
    TimerUsar As Long
    TimerUsarClick As Long
    TimerMagiaGolpe As Long
    TimerGolpeMagia As Long
    TimerGolpeUsar As Long
    TimerPuedeSerAtacado As Long
    TimerPerteneceNpc As Long
    TimerEstadoAtacable As Long
    
    Trabajando As Long  ' Para el centinela
    Ocultando As Long   ' Unico trabajo no revisado por el centinela
    
    failedUsageAttempts As Long
    
    goHome As Long
    AsignedSkills As Byte
    
        Cheat As TimeIntervalos
    Seguimiento As Intervalos
    
End Type

'Cosas faccionarias.
Public Type tFacciones
    ArmadaReal As Byte
    FuerzasCaos As Byte
    CriminalesMatados As Long
    CiudadanosMatados As Long
    RecompensasReal As Long
    RecompensasCaos As Long
    RecibioExpInicialReal As Byte
    RecibioExpInicialCaos As Byte
    RecibioArmaduraReal As Byte
    RecibioArmaduraCaos As Byte
    Reenlistadas As Byte
    NivelIngreso As Integer
    FechaIngreso As String
    MatadosIngreso As Integer 'Para Armadas nada mas
    NextRecompensa As Integer
End Type

Public Type tCrafting
    Cantidad As Long
    PorCiclo As Integer
End Type


Public Type TRetin
RetoIndex As Byte
VolverSeg As Byte
Anteriorposition As WorldPos
MeEnvio As Byte
ApuestaOro As Long
ApuestaInv As Boolean
WaitingReto As String
End Type

Public Type mSecurity
    UseItemInterval    As Byte
    ActualKey          As Integer
    PacketsCounteds    As Byte
    WaitingKey         As Integer
End Type


Public Const MAX_LOGROS As Integer = 23 'SON 24 PERO EMPIEZA DESDE 0
Public Const MAX_GUILDS_DISOLVED As Integer = 5


Public Type tUserRetoTemp
    tipo As eTipoReto
    RequiredGld As Long
    RequiredDsp As Long
    LimiteRojas As Integer
    Users() As String
    Accepts() As Byte
End Type

'Tipo de los Usuarios
' Configuración de Personajes de la cuenta
Public Type tCuentaUser
    Name As String
    Ban As Byte
    clase As eClass
    raza As eRaza
    ELV As Byte
End Type


'Tipo de los Usuarios
Public Type User
    ClaveAC As Byte
    CounterAC As Byte
    KeyPackets(MAX_KEY_PACKETS) As tPackets
    Pointers(1 To MAX_POINTERS) As tPoint
    IsPremium As Boolean
    Account As String
    Account_Bank(1 To MAX_OBJ_PREMIUM) As Obj
    Intervalos(0 To MAX_INTERVAL) As tInterval
    CantMessagePotas As Byte
    KeyPotas As Long
    LastPacket As Byte
    LastPacketComplete As Boolean
    PaquetesBasura As Long
    FailedPot As Byte
    PotFull As Boolean
    RetoTemp As tUserRetoTemp
    
    GuildDisolved(1 To MAX_GUILDS_DISOLVED) As Integer
    PosAnt As WorldPos
    Logros(0 To MAX_LOGROS) As Byte
    Mercado As tUserMercado
    ACT As Boolean
    Security As mSecurity
    Pin As String
    death As Boolean
    hungry As Boolean
    DeathAcuatico As Boolean
    Name As String
    Id As Long
    HD As String '//Disco.
    CPU_ID As String
    StaticHD As String
    Dañoapu As Long
    showName As Boolean 'Permite que los GMs oculten su nick con el comando /SHOWNAME
    
    ultimoGlobal As Long
    Char As Char 'Define la apariencia
    CharMimetizado As Char
    OrigChar As Char
    
    desc As String ' Descripcion
    DescRM As String
    
    clase As eClass
    raza As eRaza
    Genero As eGenero
    Email As String
    Hogar As eCiudad
        
    Invent As Inventario
    
    Pos As WorldPos
    
    ConnIDValida As Boolean
    ConnID As Long 'ID
    
    '[KEVIN]
    BancoInvent As BancoInventario
    '[/KEVIN]
    
    Counters As UserCounters
    
    Construir As tCrafting
    
    MascotasIndex(1 To MAXMASCOTAS) As Integer
    MascotasType(1 To MAXMASCOTAS) As Integer
    NroMascotas As Integer
    
    Stats As UserStats
    flags As UserFlags

    AskTh As Integer
    elpedidor As Integer
    
    Reputacion As tReputacion
    
    Faccion As tFacciones
    
#If SeguridadAlkon Then
    Security As SecurityData
#End If

#If ConUpTime Then
    LogOnTime As Date
    UpTime As Long
#End If

    ip As String
    
    ComUsu As tCOmercioUsuario
    
    SeguroClan As Boolean
    GuildIndex As Integer   'puntero al array global de guilds
    FundandoGuildAlineacion As ALINEACION_GUILD     'esto esta aca hasta que se parchee el cliente y se pongan cadenas de datos distintas para cada alineacion
    EscucheClan As Integer
    
    GroupIndex As Integer   'index a la party q es miembro
    GroupRequired As Integer   'index a la party q solicito
    GroupSlotUser As Byte ' Slot de miembro
    KeyCrypt As Integer
    
    AreasInfo As AreaInfo
    
    'Outgoing and incoming messages
    outgoingData As clsByteQueue
    incomingData As clsByteQueue
    
    CurrentInventorySlots As Byte
    QuestStats As tQuestStats
End Type


'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'**  T I P O S   D E    N P C S **************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Public Type NPCStats
    Alineacion As Integer
    MaxHp As Long
    MinHp As Long
    MaxHIT As Integer
    MinHIT As Integer
    def As Integer
    defM As Integer
End Type

Public Type NpcCounters
    Paralisis As Integer
    TiempoExistencia As Long
End Type

Public Type NPCFlags
    SlotEvent As Byte
    InscribedPrevio As Byte
    Invocacion As Byte
    TeamEvent As Byte
    AfectaParalisis As Byte
    Domable As Integer
    Respawn As Byte
    NPCActive As Boolean '¿Esta vivo?
    Follow As Boolean
    Faccion As Byte
    AtacaDoble As Byte
    LanzaSpells As Byte
    
    ExpCount As Long
    
    OldMovement As TipoAI
    OldHostil As Byte
    
    AguaValida As Byte
    TierraInvalida As Byte
    
    Sound As Integer
    AttackedBy As String
    AttackedFirstBy As String
    BackUp As Byte
    RespawnOrigPos As Byte
    
    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
End Type

Public Type tCriaturasEntrenador
    NpcIndex As Integer
    NpcName As String
    tmpIndex As Integer
End Type

' New type for holding the pathfinding info
Public Type NpcPathFindingInfo
    Path() As tVertice      ' This array holds the path
    Target As Position      ' The location where the NPC has to go
    PathLenght As Integer   ' Number of steps *
    CurPos As Integer       ' Current location of the npc
    TargetUser As Integer   ' UserIndex chased
    NoPath As Boolean       ' If it is true there is no path to the target location
    
    '* By setting PathLenght to 0 we force the recalculation
    '  of the path, this is very useful. For example,
    '  if a NPC or a User moves over the npc's path, blocking
    '  its way, the function NpcLegalPos set PathLenght to 0
    '  forcing the seek of a new path.
    
End Type
' New type for holding the pathfinding info

Public Type tDrops
    ObjIndex As Integer
    Amount As Long
    Probability As Byte
End Type

Public Const MAX_NPC_DROPS As Byte = 5

Public Type Npc
    DropIndex As Byte
    Invasion As Byte
    InvasionNpc As Integer
    Name As String
    Char As Char 'Define como se vera
    desc As String

    NPCtype As eNPCType
    Numero As Integer

    InvReSpawn As Byte

    Comercia As Integer
    Target As Long
    TargetNPC As Long
    TipoItems As Integer

    Veneno As Byte

    Pos As WorldPos 'Posicion
    Orig As WorldPos
    SkillDomar As Integer

    Movement As TipoAI
    Attackable As Byte
    Hostile As Byte
    PoderAtaque As Long
    PoderEvasion As Long

    Owner As Integer

    QuestNumber As Integer
       
    GiveEXP As Long
    GiveGLD As Long
    Drop(1 To MAX_NPC_DROPS) As tDrops
    
    Stats As NPCStats
    flags As NPCFlags
    Contadores As NpcCounters
    
    Invent As Inventario
    CanAttack As Byte
    
    NroExpresiones As Byte
    Expresiones() As String ' le da vida ;)
    
    NroSpells As Byte
    Spells() As Integer  ' le da vida ;)
    
    '<<<<Entrenadores>>>>>
    NroCriaturas As Integer
    Criaturas() As tCriaturasEntrenador
    MaestroUser As Integer
    MaestroNpc As Integer
    Mascotas As Integer
    
    ' New!! Needed for pathfindig
    PFINFO As NpcPathFindingInfo
    AreasInfo As AreaInfo
    
    'Hogar
    Ciudad As Byte
End Type

'**********************************************************
'**********************************************************
'******************** Tipos del mapa **********************
'**********************************************************
'**********************************************************
'Tile
Public Type MapBlock
    Blocked As Byte
    Graphic(1 To 4) As Integer
    userindex As Integer
    NpcIndex As Integer
    ObjInfo As Obj
    TileExit As WorldPos
    trigger As eTrigger
    ObjEvent As Byte
End Type

'Info del mapa
Type MapInfo
    levelmin As Byte
    levelmax As Byte
    SeCaenItems As Boolean
    NumUsers As Integer
    Music As String
    Name As String
    StartPos As WorldPos
    OnDeathGoTo As WorldPos
    
    MapVersion As Integer
    Pk As Boolean
    MagiaSinEfecto As Byte
    NoEncriptarMP As Byte
    
    ' Anti Magias/Habilidades
    InviSinEfecto As Byte
    ResuSinEfecto As Byte
    OcultarSinEfecto As Byte
    InvocarSinEfecto As Byte
    
    RoboNpcsPermitido As Byte
    
    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte
End Type


'********** V A R I A B L E S     P U B L I C A S ***********

Public SERVERONLINE As Boolean
Public ULTIMAVERSION As String
Public BackUp As Boolean ' TODO: Se usa esta variable ?

Public ListaRazas(1 To NUMRAZAS) As String
Public SkillsNames(1 To NUMSKILLS) As String
Public ListaClases(1 To NUMCLASES) As String
Public ListaAtributos(1 To NUMATRIBUTOS) As String


Public recordusuarios As Integer

'
'Directorios
'

''
'Ruta base del server, en donde esta el "server.ini"
Public IniPath As String

''
'Ruta base para guardar los chars
Public CharPath As String

''
'Ruta base para los archivos de mapas
Public MapPath As String

''
'Ruta base para los DATs
Public DatPath As String

''
'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

''
'Numero de usuarios actual
Public NumUsers As Integer
Public LastUser As Integer
Public LastChar As Integer
Public NumChars As Integer
Public LastNPC As Integer
Public NumNPCs As Integer
Public NumFX As Integer
Public NumMaps As Integer
Public NumObjDatas As Integer
Public NumeroHechizos As Integer
Public AllowMultiLogins As Byte
Public IdleLimit As Integer
Public MaxUsers As Integer
Public HideMe As Byte
Public LastBackup As String
Public Minutos As String
Public haciendoBK As Boolean
Public PuedeCrearPersonajes As Integer
Public ServerSoloGMs As Integer

''
'Esta activada la verificacion MD5 ?
Public MD5ClientesActivado As Byte
Public ELMD5 As String

Public EnPausa As Boolean
Public EnTesting As Boolean


'*****************ARRAYS PUBLICOS*************************
Public UserList() As User 'USUARIOS
Public Npclist(1 To MAXNPCS) As Npc 'NPCS
Public MapData() As MapBlock
Public MapInfo() As MapInfo
Public Hechizos() As tHechizo
Public CharList(1 To MAXCHARS) As Integer
Public ObjData() As ObjData
Public FX() As FXdata
Public SpawnList() As tCriaturasEntrenador
Public LevelSkill(1 To 50) As LevelSkill
Public ForbidenNames() As String
Public ArmasHerrero() As Integer
Public ArmadurasHerrero() As Integer
Public ObjCarpintero() As Integer
Public MD5s() As String
Public BanIps As New Collection
Public BanHDs As New Collection '//Disco
Public BanT0s     As New Collection
Public ModClase(1 To NUMCLASES) As ModClase
Public ModRaza(1 To NUMRAZAS) As ModRaza
Public ModVida(1 To NUMCLASES) As Double
Public DistribucionEnteraVida(1 To 5) As Integer
Public DistribucionSemienteraVida(1 To 4) As Integer
Public Ciudades(1 To NUMCIUDADES) As WorldPos
Public distanceToCities() As HomeDistance
Public QuestList() As tQuest
'*********************************************************

Type HomeDistance
    distanceToCity(1 To 5) As Integer
End Type

Public Nix As WorldPos
Public Ullathorpe As WorldPos
Public Banderbill As WorldPos
Public Lindos As WorldPos
Public Arghal As WorldPos
Public Arkhein As WorldPos
Public Nemahuak As WorldPos

Public Prision As WorldPos
Public Libertad As WorldPos

Public Ayuda As New cCola
Public ConsultaPopular As New ConsultasPopulares
Public SonidosMapas As New SoundMapInfo

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef destination As Any, ByVal length As Long)

Public Enum e_ObjetosCriticos
    Manzana = 1
    Manzana2 = 2
    ManzanaNewbie = 467
End Enum

Public Enum eMessages
    DontSeeAnything
    NPCSwing
    NPCKillUser
    BlockedWithShieldUser
    BlockedWithShieldother
    UserSwing
    SafeModeOn
    SafeModeOff
    DragOnn
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
    Home
    CancelHome
    FinishHome
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
    invisible = 17             '/INVISIBLE
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
    SearchObj = 151             'BUSCAR OBJ
    SearcherShow = 152           'BUSC
End Enum
Public Const MATRIX_INITIAL_MAP As Integer = 1

Public Const GOHOME_PENALTY As Integer = 5
Public Const GM_MAP As Integer = 49

Public Const TELEP_OBJ_INDEX As Integer = 1012


' Por ahora la dejo constante.. SI se quisiera extender la propiedad de paralziar, se podria hacer
' una nueva variable en el dat.
Public Const GUANTE_HURTO As Integer = 873

Public Administradores As clsIniManager

Public Type tRange
  Name As String
  Tag As String
  color As String
End Type

Public RangeGm() As tRange
