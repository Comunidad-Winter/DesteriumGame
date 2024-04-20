Attribute VB_Name = "mDeclares"
Option Explicit


Public NumCuerpos As Integer


Public IniPath       As String
Public MiCabecera    As tCabecera


Public Type tCabecera 'Cabecera de los con

    Desc As String * 255
    CRC As Long
    MagicWord As Long

End Type


'Direcciones
Public Enum E_Heading

    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4

End Enum

'Posicion en un mapa
Public Type Position

    X As Long
    Y As Long

End Type


'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData

    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    Speed As Single
    
    'src As wGL_Rectangle
    
    active As Boolean
    MiniMap_color As Long
    Alpha As Boolean

End Type


'Apunta a una estructura grhdata y mantiene la animacion
Public Type Grh

    Grhindex As Long
    FrameCounter As Single
    FrameTimer As Single
    Speed As Single
    Started As Byte
    Loops As Integer

End Type

'Lista de cabezas
Public Type tIndiceCabeza

    Head(1 To 4) As Long

End Type

Public Type tIndiceCuerpo

    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer
    
    BodyOffsetX(1 To 4) As Integer
    BodyOffsetY(1 To 4) As Integer
End Type

Public Type tIndiceFx

    Animacion As Long
    OffsetX As Integer
    OffsetY As Integer

End Type

Public Type BodyData

    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
    BodyOffset As Position

End Type 'Lista de cuerpos

Public Type HeadData

    Head(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type 'Lista de cabezas

Public Type WeaponAnimData

    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type 'Lista de las animaciones de las armas

Public Type ShieldAnimData

    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type 'Lista de las animaciones de los escudos

Public Type AuraData

    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type 'Lista de las animaciones de las auras

Public GrhData()        As GrhData 'Guarda todos los grh

Public BodyData()       As BodyData

Public BodyDataAttack() As BodyData

Public HeadData()       As HeadData

Public FxData()         As tIndiceFx

Public WeaponAnimData() As WeaponAnimData


Public ShieldAnimData() As ShieldAnimData

Public CascoAnimData()  As HeadData

Public AuraAnimData()   As AuraData


' Uso
Public MisCuerpos() As tIndiceCuerpo

Public Function Heading_To_String(ByVal Heading As E_Heading) As String
    Select Case Heading
    
        Case E_Heading.NORTH
            Heading_To_String = "Arriba"
        Case E_Heading.SOUTH
            Heading_To_String = "Abajo"
        Case E_Heading.WEST
            Heading_To_String = "Izquierda"
        Case E_Heading.EAST
            Heading_To_String = "Derecha"
    
    End Select
End Function
