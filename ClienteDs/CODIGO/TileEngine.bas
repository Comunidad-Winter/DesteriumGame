Attribute VB_Name = "Mod_TileEngine"

Option Explicit

Private OffsetCounterX As Single
Private OffsetCounterY As Single
    
'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

Private Const GrhFogata As Integer = 1521

''
'Sets a Grh animation to loop indefinitely.
Public Const INFINITE_LOOPS As Integer = -1


'Encabezado bmp
Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    X As Long
    Y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    map As Integer
    X As Integer
    Y As Integer
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
    
    sX1 As Single
    sX2 As Single
    sY1 As Single
    sY2 As Single
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type


'Apariencia del personaje
Public Type Char
    MinHp As Integer
    MaxHp As Integer
    MinMan As Integer
    MaxMan As Integer
    Movimient As Boolean
    LastDialog     As String
    Active As Byte
    Heading As E_Heading
    Pos As Position
    
        cCont As Integer
   Drawers As Integer
   
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    
    fX As Grh
    FxIndex As Integer
    
    Team As Byte
    
    Criminal As Byte
    Atacable As Boolean
    
    Nombre As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    muerto As Boolean
    Invisible As Boolean
    Infected As Byte
    Angel As Byte
    Demonio As Byte
    
    priv As Byte
    
    AnimTime As Byte
End Type

'Info de un objeto
Public Type Obj
    ObjIndex As Integer
    Amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Damage As DList
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
End Type

'DX7 Objects
Public DirectX As New DirectX7

#If Wgl = 0 Then
Public DirectDraw As DirectDraw7
Private PrimarySurface As DirectDrawSurface7
Private PrimaryClipper As DirectDrawClipper
Private BackBufferSurface As DirectDrawSurface7
#End If

Public IniPath As String
Public MapPath As String


'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public EngineRun As Boolean

Public FPS As Long
Public FramesPerSecCounter As Long
Public fpsLastCheck As Long

'Tamaño del la vista en Tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

Public HalfWindowTileWidth As Integer
Public HalfWindowTileHeight As Integer

'Offset del desde 0,0 del main view
Public MainViewTop As Integer
Public MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

Public TileBufferPixelOffsetX As Integer
Public TileBufferPixelOffsetY As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Public timerElapsedTime As Single
Public timerTicksPerFrame As Single
Public engineBaseSpeed As Single


Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer


Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT

Public MainViewWidth As Integer
Public MainViewHeight As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte




'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bTecho       As Boolean 'hay techo?
Public brstTick     As Long

Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer

Public charlist(1 To 10000) As Char

' Used by GetTextExtentPoint32
Private Type Size
    cx As Long
    cy As Long
End Type

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'#If ConAlfaB Then

Private Declare Function BltAlphaFast Lib "vbabdx" (ByRef lpDDSDest As Any, _
    ByRef lpDDSSource As Any, ByVal iWidth As Long, ByVal iHeight As Long, ByVal _
    pitchSrc As Long, ByVal pitchDst As Long, ByVal dwMode As Long) As Long
Private Declare Function BltEfectoNoche Lib "vbabdx" (ByRef lpDDSDest As Any, _
    ByVal iWidth As Long, ByVal iHeight As Long, ByVal pitchDst As Long, ByVal _
    dwMode As Long) As Long

'#End If
'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency _
    As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" _
    (lpPerformanceCount As Currency) As Long

'Text width computation. Needed to center text.
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias _
    "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal _
    cbString As Long, lpSize As Size) As Long

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As _
    Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As _
    Long, ByVal Y As Long) As Long

Sub CargarCabezas()
          Dim n As Integer
          Dim i As Long
          Dim Numheads As Integer
          Dim Miscabezas() As tIndiceCabeza
          
10        n = FreeFile()
20        Open App.path & "\Recursos\Clases\Clases.ind" For Binary Access Read As #n
          
30        If Not FileExist(App.path & "\INIT\CABEZAS.IND", vbArchive) Then
40    End
50    End If

60    If Not FileExist(App.path & "\Recursos\Clases\Clases.ind", vbArchive) Then
70    End
80    End If
          
          'cabecera
90        Get #n, , MiCabecera
          
          'num de cabezas
100       Get #n, , Numheads
          
          'Resize array
110       ReDim HeadData(0 To Numheads) As HeadData
120       ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
          
130       For i = 1 To Numheads
140           Get #n, , Miscabezas(i)
              
150           If Miscabezas(i).Head(1) Then
160               Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
170               Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
180               Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
190               Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
200           End If
210       Next i
          
220       Close #n
End Sub

Sub CargarCascos()
          Dim n As Integer
          Dim i As Long
          Dim NumCascos As Integer

          Dim Miscabezas() As tIndiceCabeza
          
10        n = FreeFile()
20        Open App.path & "\init\Cascos.ind" For Binary Access Read As #n
          
          'cabecera
30        Get #n, , MiCabecera
          
          'num de cabezas
40        Get #n, , NumCascos
          
          'Resize array
50        ReDim CascoAnimData(0 To NumCascos) As HeadData
60        ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
          
70        For i = 1 To NumCascos
80            Get #n, , Miscabezas(i)
              
90            If Miscabezas(i).Head(1) Then
100               Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
110               Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
120               Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
130               Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
140           End If
150       Next i
          
160       Close #n
End Sub

Sub CargarCuerpos()
          Dim n As Integer
          Dim i As Long
          Dim NumCuerpos As Integer
          Dim MisCuerpos() As tIndiceCuerpo
          
10        n = FreeFile()
20        Open App.path & "\init\Personajes.ind" For Binary Access Read As #n
          
          'cabecera
30        Get #n, , MiCabecera
          
          'num de cabezas
40        Get #n, , NumCuerpos
          
          'Resize array
50        ReDim BodyData(0 To NumCuerpos) As BodyData
60        ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
          
70        For i = 1 To NumCuerpos
80            Get #n, , MisCuerpos(i)
              
90                If MisCuerpos(i).Body(1) Then
100                   InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
110                   InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
120                   InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
130                   InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
                      
140                   BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
150                   BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
160               End If
170       Next i
          
180       Close #n
End Sub

Sub CargarFxs()
          Dim n As Integer
          Dim i As Long
          Dim NumFxs As Integer
          
10        n = FreeFile()
20        Open App.path & "\init\Fxs.ind" For Binary Access Read As #n
          
          'cabecera
30        Get #n, , MiCabecera
          
          'num de cabezas
40        Get #n, , NumFxs
          
          'Resize array
50        ReDim FxData(1 To NumFxs) As tIndiceFx
          
60        For i = 1 To NumFxs
70            Get #n, , FxData(i)
80        Next i
          
90        Close #n
End Sub

Sub CargarTips()
          Dim n As Integer
          Dim i As Long
          Dim NumTips As Integer
          
10        n = FreeFile
20        Open App.path & "\init\Tips.ayu" For Binary Access Read As #n
          
          'cabecera
30        Get #n, , MiCabecera
          
          'num de cabezas
40        Get #n, , NumTips
          
          'Resize array
50        ReDim Tips(1 To NumTips) As String * 255
          
60        For i = 1 To NumTips
70            Get #n, , Tips(i)
80        Next i
          
90        Close #n
End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef _
    tX As Byte, ByRef tY As Byte)
      '******************************************
      'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
      '******************************************
10        tX = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
20        tY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As _
    Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal _
    Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
10    On Error Resume Next
          'Apuntamos al ultimo Char
20        If CharIndex > LastChar Then LastChar = CharIndex
          
30        With charlist(CharIndex)
              'If the char wasn't allready active (we are rewritting it) don't increase char count
40            If .Active = 0 Then NumChars = NumChars + 1
              
50            If Arma = 0 Then Arma = 2
60            If Escudo = 0 Then Escudo = 2
70            If Casco = 0 Then Casco = 2
              
80            .iHead = Head
90            .iBody = Body
100           .Head = HeadData(Head)
110           .Body = BodyData(Body)
120           .Arma = WeaponAnimData(Arma)
              
130           .Escudo = ShieldAnimData(Escudo)
140           .Casco = CascoAnimData(Casco)
              
150           .Heading = Heading
              
              'Reset moving stats
160           .Moving = 0
170           .MoveOffsetX = 0
180           .MoveOffsetY = 0
              
              'Update position
190           .Pos.X = X
200           .Pos.Y = Y
              
              'Make active
210           .Active = 1


             #If Wgl = 1 Then
                ' Culpa del que programo el sistema de area (nos manda duplicado)
                Dim RangeX As Single, RangeY As Single
                Call GetCharacterDimension(CharIndex, RangeX, RangeY)
                
                Call g_Swarm.InsertDynamic(CharIndex, 5, X, Y, RangeX, RangeY)
              #End If
220       End With
          
          'Plot on map
230       MapData(X, Y).CharIndex = CharIndex
End Sub
Public Function GetCharacterDimension(ByVal CharIndex As Integer, ByRef RangeX As Single, ByRef RangeY As Single) ' RTREE

    With charlist(CharIndex)
        If (.iBody <> 0) Then
            RangeX = GrhData(.Body.Walk(.Heading).GrhIndex).TileWidth
            RangeY = GrhData(.Body.Walk(.Heading).GrhIndex).TileHeight
        End If
        
        If (.iHead <> 0) Then
            If (GrhData(.Head.Head(.Heading).GrhIndex).TileWidth > RangeX) Then
                RangeX = GrhData(.Head.Head(.Heading).GrhIndex).TileWidth
            End If
            
            RangeY = RangeY + GrhData(.Head.Head(.Heading).GrhIndex).TileHeight + 2# 'Name + Guild
        End If
    End With
        
End Function


Sub ResetCharInfo(ByVal CharIndex As Integer)

          
10        With charlist(CharIndex)
20            .Active = 0
30            .Team = 0
40            .Criminal = 0
50            .Atacable = False
60            .FxIndex = 0
70            .Invisible = False
80            .Infected = 0
90            .Angel = 0
100           .Demonio = 0
              
110           .Moving = 0
120           .muerto = False
130           .Nombre = ""
140           .pie = False
150           .Pos.X = 0
160           .Pos.Y = 0
170           .UsandoArma = False
180       End With
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
      '*****************************************************************
      'Erases a character from CharList and map
      '*****************************************************************
10    On Error Resume Next
20        charlist(CharIndex).Active = 0
          
          'Update lastchar
30        If CharIndex = LastChar Then
40            Do Until charlist(LastChar).Active = 1
50                LastChar = LastChar - 1
60                If LastChar = 0 Then Exit Do
70            Loop
80        End If
          
          
          If charlist(CharIndex).Pos.X > 0 And charlist(CharIndex).Pos.Y > 0 Then
90          MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y).CharIndex = 0

            #If Wgl = 1 Then
                Call g_Swarm.RemoveDynamic(CharIndex)
            #End If
          End If
          
          'Remove char's dialog
100       Call Dialogos.RemoveDialog(CharIndex)
          
110       Call ResetCharInfo(CharIndex)
          
          'Update NumChars
120       NumChars = NumChars - 1
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal _
    Started As Byte = 2)
      '*****************************************************************
      'Sets up a grh. MUST be done before rendering
      '*****************************************************************
10        Grh.GrhIndex = GrhIndex
          
20        If GrhIndex = 0 Then Exit Sub
          
30        If Started = 2 Then
40            If GrhData(Grh.GrhIndex).NumFrames > 1 Then
50                Grh.Started = 1
60            Else
70                Grh.Started = 0
80            End If
90        Else
              'Make sure the graphic can be started
100           If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
110           Grh.Started = Started
120       End If
          
          
130       If Grh.Started Then
140           Grh.Loops = INFINITE_LOOPS
150       Else
160           Grh.Loops = 0
170       End If
          
180       Grh.FrameCounter = 1
190       Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
      '*****************************************************************
      'Starts the movement of a character in nHeading direction
      '*****************************************************************
          Dim addX As Integer
          Dim addY As Integer
          Dim X As Integer
          Dim Y As Integer
          Dim nX As Integer
          Dim nY As Integer
          
10        With charlist(CharIndex)
20            X = .Pos.X
30            Y = .Pos.Y
              
              'Figure out which way to move
40            Select Case nHeading
                  Case E_Heading.NORTH
50                    addY = -1
              
60                Case E_Heading.EAST
70                    addX = 1
              
80                Case E_Heading.SOUTH
90                    addY = 1
                  
100               Case E_Heading.WEST
110                   addX = -1
120           End Select
              
130           nX = X + addX
140           nY = Y + addY
              'Creditos a la puta de miqueas
150           If nX < 1 Or nX > 100 Or nY < 1 Or nY > 100 Then Exit Sub
160           MapData(nX, nY).CharIndex = CharIndex
170           .Pos.X = nX
180           .Pos.Y = nY

              #If Wgl = 1 Then
              Call g_Swarm.Move(CharIndex, nX, nY)
              #End If
              
190           MapData(X, Y).CharIndex = 0
              
200           .MoveOffsetX = -1 * (TilePixelWidth * addX)
210           .MoveOffsetY = -1 * (TilePixelHeight * addY)
              
220           .Moving = 1
230           .Heading = nHeading
              
240           .scrollDirectionX = addX
250           .scrollDirectionY = addY
260       End With
          
270       If UserEstado = 0 Then Call DoPasosFx(CharIndex)
          
          'areas viejos
          
          
280       If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > _
              MaxLimiteX) Then
290           If CharIndex <> UserCharIndex Then
                  Call EraseChar(CharIndex)
310           End If
320       End If
End Sub

Public Sub DoFogataFx()
          Dim location As Position
          
10        If bFogata Then
20            bFogata = HayFogata(location)
30            If Not bFogata Then
40                Call Audio.StopWave(FogataBufferIndex)
50                FogataBufferIndex = 0
60            End If
70        Else
80            bFogata = HayFogata(location)
90            If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = _
                  Audio.PlayWave("fuego.wav", location.X, location.Y, LoopStyle.Enabled)
100       End If
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
10        With charlist(CharIndex).Pos
20            EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + _
                  MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + _
                  MinYBorder
30        End With
End Function

  Sub DoPasosFx(ByVal CharIndex As Integer)
        Dim Paso As Byte
10      If UserMontando = True Then Exit Sub
20        If Not UserNavegando Then
30            With charlist(CharIndex)
40                If .muerto = False And EstaPCarea(CharIndex) = True And (.priv <> _
                      5) Then
50                    .pie = Not .pie
60                    Paso = Map_GetTerrenoDePaso(GrhData(MapData(.Pos.X, _
                          .Pos.Y).Graphic(1).GrhIndex).FileNum)
               
70                    If Paso = 1 Then
80                        If .pie Then
90                            Call Audio.PlayWave(SND_PASOS7)
100                       Else
110                           Call Audio.PlayWave(SND_PASOS8)
120                       End If
130                   ElseIf Paso = 2 Or Paso = 5 Then
140                       If .pie Then
150                           Call Audio.PlayWave(SND_PASOS1) 'Si no son Pasto ,nieve o arena,se pone este default, 1 y 2!
160                       Else
170                           Call Audio.PlayWave(SND_PASOS2) 'Si no son Pasto ,nieve o arena,se pone este default, 1 y 2!
180                       End If
190                   ElseIf Paso = 3 Then
200                       If .pie Then
210                           Call Audio.PlayWave(SND_PASOS5)
220                       Else
230                           Call Audio.PlayWave(SND_PASOS6)
240                       End If
250                   ElseIf Paso = 4 Then
260                       If .pie Then
270                          Call Audio.PlayWave(SND_PASOS3)
280                       Else
290                          Call Audio.PlayWave(SND_PASOS4)
300                       End If
310                   End If
320               End If
330           End With
         ' ElseIf UserMontando = True Then
           '   Call Audio.PlayWave(23, charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.y)
340       ElseIf UserNavegando = True Then
          '- Saque el sonido del Wather.
350       End If
End Sub
Private Function Map_GetTerrenoDePaso(ByVal TerrainFileNum As Integer) As Byte
10      If (TerrainFileNum >= 6000 And TerrainFileNum <= 6004) Or (TerrainFileNum >= _
            550 And TerrainFileNum <= 552) Or (TerrainFileNum >= 6018 And TerrainFileNum _
            <= 6020) Then
20      Map_GetTerrenoDePaso = 1
30      Exit Function
40      ElseIf (TerrainFileNum >= 7501 And TerrainFileNum <= 7507) Or (TerrainFileNum _
            = 7500 Or TerrainFileNum = 7508 Or TerrainFileNum = 1533 Or TerrainFileNum = _
            2508) Then
50      Map_GetTerrenoDePaso = 2
60      Exit Function
70      ElseIf (TerrainFileNum >= 10139 And TerrainFileNum <= 10143) Then
80      Map_GetTerrenoDePaso = 3
90      Exit Function
100     ElseIf TerrainFileNum = 6021 Then
110     Map_GetTerrenoDePaso = 4
120     Exit Function
130     Else
140     Map_GetTerrenoDePaso = 5
150     End If
End Function

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As _
    Integer)
10    On Error Resume Next
          Dim X As Integer
          Dim Y As Integer
          Dim addX As Integer
          Dim addY As Integer
          Dim nHeading As E_Heading
          
20        With charlist(CharIndex)
30            X = .Pos.X
40            Y = .Pos.Y
              
50            MapData(X, Y).CharIndex = 0
              
60            addX = nX - X
70            addY = nY - Y
              
80            If Sgn(addX) = 1 Then
90                nHeading = E_Heading.EAST
100           ElseIf Sgn(addX) = -1 Then
110               nHeading = E_Heading.WEST
120           ElseIf Sgn(addY) = -1 Then
130               nHeading = E_Heading.NORTH
140           ElseIf Sgn(addY) = 1 Then
150               nHeading = E_Heading.SOUTH
160           End If
              
170           MapData(nX, nY).CharIndex = CharIndex
              
180           .Pos.X = nX
190           .Pos.Y = nY
              
              #If Wgl = 1 Then
                Call g_Swarm.Move(CharIndex, nX, nY)
              #End If
200           .MoveOffsetX = -1 * (TilePixelWidth * addX)
210           .MoveOffsetY = -1 * (TilePixelHeight * addY)
              
220           .Moving = 1
230           .Heading = nHeading
              
240           .scrollDirectionX = Sgn(addX)
250           .scrollDirectionY = Sgn(addY)
              
              'parche para que no medite cuando camina
260           If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or _
                  .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or _
                  .FxIndex = FxMeditar.XXGRANDE Then
270               .FxIndex = 0
280           End If
              
290       End With
          
300       If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
          
310       If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > _
              MaxLimiteX) Then
320           Call EraseChar(CharIndex)
330       End If
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
      '******************************************
      'Starts the screen moving in a direction
      '******************************************
          Dim X As Integer
          Dim Y As Integer
          Dim tX As Integer
          Dim tY As Integer
          
          'Figure out which way to move
10        Select Case nHeading
              Case E_Heading.NORTH
20                Y = -1
              
30            Case E_Heading.EAST
40                X = 1
              
50            Case E_Heading.SOUTH
60                Y = 1
              
70            Case E_Heading.WEST
80                X = -1
90        End Select
          
          'Fill temp pos
100       tX = UserPos.X + X
110       tY = UserPos.Y + Y
          
          'Check to see if its out of bounds
120       If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder _
              Then
130           Exit Sub
140       Else
              'Start moving... MainLoop does the rest
150           AddtoUserPos.X = X
160           UserPos.X = tX
170           AddtoUserPos.Y = Y
180           UserPos.Y = tY
190           UserMoving = 1
              
200           bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                  MapData(UserPos.X, UserPos.Y).Trigger = 2 Or MapData(UserPos.X, _
                  UserPos.Y).Trigger = 4, True, False)
210       End If
End Sub

Private Function HayFogata(ByRef location As Position) As Boolean
          Dim j As Long
          Dim k As Long
          
10        For j = UserPos.X - 8 To UserPos.X + 8
20            For k = UserPos.Y - 6 To UserPos.Y + 6
30                If InMapBounds(j, k) Then
40                    If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
50                        location.X = j
60                        location.Y = k
                          
70                        HayFogata = True
80                        Exit Function
90                    End If
100               End If
110           Next k
120       Next j
End Function

Function NextOpenChar() As Integer
      '*****************************************************************
      'Finds next open char slot in CharList
      '*****************************************************************
          Dim LoopC As Long
          Dim Dale As Boolean
          
10        LoopC = 1
20        Do While charlist(LoopC).Active And Dale
30            LoopC = LoopC + 1
40            Dale = (LoopC <= UBound(charlist))
50        Loop
          
60        NextOpenChar = LoopC
End Function

Private Sub DetectSize(ByVal FileNum As Long, ByRef Width As Integer, ByRef Height As Integer)
    Dim temp As String
    
    temp = GetVar(App.path & "\INIT\MEDIDAS.DAT", "INIT", FileNum) 'Read.GetValue("INIT", FileNum)
    Width = Val(ReadField(1, temp, Asc("-")))
    Height = Val(ReadField(2, temp, Asc("-")))
End Sub

Public Function LoadGrhData() As Boolean
10    On Error GoTo ErrorHandler
          Dim Grh As Long
          Dim Frame As Long
          Dim grhCount As Long
          Dim handle As Integer
          Dim fileVersion As Long


          'Open files
20        handle = FreeFile()
          
30        Open App.path & "\INIT\" & GraphicsFile For Binary Access Read As handle
40        Seek #1, 1
          
          'Get file version
50        Get handle, , fileVersion
          
          'Get number of grhs
60        Get handle, , grhCount
          
          'Resize arrays
70        ReDim GrhData(0 To grhCount) As GrhData
          
80        While Not EOF(handle)
90            Get handle, , Grh
              
100           With GrhData(Grh)
                  'Get number of frames
110               Get handle, , .NumFrames
120               If .NumFrames <= 0 Then GoTo ErrorHandler
                  
130               ReDim .Frames(1 To GrhData(Grh).NumFrames)
                  
140               If .NumFrames > 1 Then
                      'Read a animation GRH set
150                   For Frame = 1 To .NumFrames
160                       Get handle, , .Frames(Frame)
170                       If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
180                           GoTo ErrorHandler
190                       End If
200                   Next Frame
                      
                      Dim sngSpeed As Single
                      
210                   Get handle, , sngSpeed
                      
220                   .Speed = sngSpeed '+ 16.3333

230                   If .Speed <= 0 Then GoTo ErrorHandler
                      
                      'Compute width and height
240                   .pixelHeight = GrhData(.Frames(1)).pixelHeight
250                   If .pixelHeight <= 0 Then GoTo ErrorHandler
                      
260                   .pixelWidth = GrhData(.Frames(1)).pixelWidth
270                   If .pixelWidth <= 0 Then GoTo ErrorHandler
                      
280                   .TileWidth = GrhData(.Frames(1)).TileWidth
290                   If .TileWidth <= 0 Then GoTo ErrorHandler
                      
300                   .TileHeight = GrhData(.Frames(1)).TileHeight
310                   If .TileHeight <= 0 Then GoTo ErrorHandler
320               Else
                      'Read in normal GRH data
330                   Get handle, , .FileNum
340                   If .FileNum <= 0 Then GoTo ErrorHandler
                      
350                   Get handle, , GrhData(Grh).sX
360                   If .sX < 0 Then GoTo ErrorHandler
                      
370                   Get handle, , .sY
380                   If .sY < 0 Then GoTo ErrorHandler
                      
390                   Get handle, , .pixelWidth
400                   If .pixelWidth <= 0 Then GoTo ErrorHandler
                      
410                   Get handle, , .pixelHeight
420                   If .pixelHeight <= 0 Then GoTo ErrorHandler
                      
                      'Compute width and height
430                   .TileWidth = .pixelWidth / TilePixelHeight
440                   .TileHeight = .pixelHeight / TilePixelWidth
                      
450                   .Frames(1) = Grh

                      Dim Width As Integer, Height As Integer
                    '  DetectSize .FileNum, Width, Height

                     ' .sX1 = .sX / Width
                      '.sY1 = .sY / Height
                        
                      '.sX2 = .sX1 + .pixelWidth / Width
                      '.sY2 = .sY1 + .pixelHeight / Height
460               End If
470           End With
480       Wend
          
490       Close handle

500       LoadGrhData = True
510   Exit Function

ErrorHandler:
520       LoadGrhData = False
End Function

Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
      '*****************************************************************
      'Checks to see if a tile position is legal
      '*****************************************************************
          'Limites del mapa
10        If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
20            Exit Function
30        End If
          
          'Tile Bloqueado?
40        If MapData(X, Y).Blocked = 1 Then
50            Exit Function
60        End If
          
          '¿Hay un personaje?
70        If MapData(X, Y).CharIndex > 0 Then
80            Exit Function
90        End If
         
100       If UserNavegando <> HayAgua(X, Y) Then
110           Exit Function
120       End If
          
130       bTecho = (MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, _
              UserPos.Y).Trigger = 2 Or MapData(UserPos.X, UserPos.Y).Trigger = 4)
                             
140               If bTecho And UserMontando = True Then Exit Function
          
150       LegalPos = True
End Function

Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
      '*****************************************************************
      'Author: ZaMa
      'Last Modify Date: 01/08/2009
      'Checks to see if a tile position is legal, including if there is a casper in the tile
      '10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
      '01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
      '*****************************************************************
          Dim CharIndex As Integer
          
          'Limites del mapa
10        If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
20            Exit Function
30        End If
          
          'Tile Bloqueado?
40        If MapData(X, Y).Blocked = 1 Then
50            Exit Function
60        End If
          
70        CharIndex = MapData(X, Y).CharIndex
          '¿Hay un personaje?
80        If CharIndex > 0 Then
          
90            If MapData(UserPos.X, UserPos.Y).Blocked = 1 Then
100               Exit Function
110           End If
              
120           With charlist(CharIndex)
                  ' Si no es casper, no puede pasar
130               If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
140                   Exit Function
150               Else
                      ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
160                   If HayAgua(UserPos.X, UserPos.Y) Then
170                       If Not HayAgua(X, Y) Then Exit Function
180                   Else
                          ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
190                       If HayAgua(X, Y) Then Exit Function
200                   End If
                      
                      ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
210                   If charlist(UserCharIndex).priv > 0 And _
                          charlist(UserCharIndex).priv < 6 Then
220                       If charlist(UserCharIndex).Invisible = True Or UserOcu = 1 _
                              Then Exit Function
230                   End If
240               End If
250           End With
260       End If
         
270       If UserNavegando <> HayAgua(X, Y) Then
280           Exit Function
290       End If
          
300       MoveToLegalPos = True
End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
      '*****************************************************************
      'Checks to see if a tile position is in the maps bounds
      '*****************************************************************
10        If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize _
              Then
20            Exit Function
30        End If
          
40        InMapBounds = True
End Function

#If Wgl = 0 Then
Private Sub DDrawGrhtoSurface(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As _
    Integer, ByVal center As Byte, ByVal Animate As Byte)
          Dim CurrentGrhIndex As Integer
          Dim SourceRect As RECT
10    On Error GoTo error
              
20        If Animate Then
30            If Grh.Started = 1 Then
40                Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * _
                      GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * Movement_Speed
50                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
60                    If Grh.GrhIndex <> 0 Then Grh.FrameCounter = (Grh.FrameCounter _
                          Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                      
70                    If Grh.Loops <> INFINITE_LOOPS Then
80                        If Grh.Loops > 0 Then
90                            Grh.Loops = Grh.Loops - 1
100                       Else
110                           Grh.Started = 0
120                       End If
130                   End If
140               End If
150           End If
160       End If
          
          'If Grh.GrhIndex = 0 Then Exit Sub
          
          'Figure out what frame to draw (always 1 if not animated)
170       If Grh.GrhIndex <> 0 Then CurrentGrhIndex = _
              GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
          
180       With GrhData(CurrentGrhIndex)
              'Center Grh over X,Y pos
190           If center Then
200               If .TileWidth <> 1 Then
210                   X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ _
                          2
220               End If
                  
230               If .TileHeight <> 1 Then
240                   Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
250               End If
260           End If
              
270           SourceRect.Left = .sX
280           SourceRect.Top = .sY
290           SourceRect.Right = SourceRect.Left + .pixelWidth
300           SourceRect.Bottom = SourceRect.Top + .pixelHeight
              
              'Draw
310           Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), _
                  SourceRect, DDBLTFAST_WAIT)
320       End With
330   Exit Sub

error:
340       If Err.number = 9 And Grh.FrameCounter < 1 Then
350           Grh.FrameCounter = 1
360           Resume
370       Else
380           MsgBox _
                  "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." _
                  & vbcrlf & "Descripción del error: " & vbcrlf & Err.Description, _
                  vbExclamation, "[ " & Err.number & " ] Error"
390           End
400       End If
End Sub

Sub DDrawTransGrhIndextoSurface(ByVal GrhIndex As Integer, ByVal X As Integer, _
    ByVal Y As Integer, ByVal center As Byte)
          Dim SourceRect As RECT
          
10        With GrhData(GrhIndex)
              'Center Grh over X,Y pos
20            If center Then
30                If .TileWidth <> 1 Then
40                    X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ _
                          2
50                End If
                  
60                If .TileHeight <> 1 Then
70                    Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
80                End If
90            End If
              
100           SourceRect.Left = .sX
110           SourceRect.Top = .sY
120           SourceRect.Right = SourceRect.Left + .pixelWidth
130           SourceRect.Bottom = SourceRect.Top + .pixelHeight
              
              'Draw
140           Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), _
                  SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
150       End With
End Sub

Sub DDrawTransGrhtoSurface(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As _
    Integer, ByVal center As Byte, ByVal Animate As Byte)
      '*****************************************************************
      'Draws a GRH transparently to a X and Y position
      '*****************************************************************
          Dim CurrentGrhIndex As Integer
          Dim SourceRect As RECT
          
10    On Error GoTo error
          
20       If Animate Then
30            If Grh.Started = 1 Then
40                Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * _
                      GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
                  
50                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
60                    Grh.FrameCounter = (Grh.FrameCounter Mod _
                          GrhData(Grh.GrhIndex).NumFrames) + 1
                      
70                    If Grh.Loops <> INFINITE_LOOPS Then
80                        If Grh.Loops > 0 Then
90                            Grh.Loops = Grh.Loops - 1
100                       Else
110                           Grh.Started = 0
120                       End If
130                   End If
140               End If
150           End If

160       End If
          
          'Figure out what frame to draw (always 1 if not animated)
170       CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
          
180       With GrhData(CurrentGrhIndex)
              'Center Grh over X,Y pos
190           If center Then
200               If .TileWidth <> 1 Then
210                   X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ _
                          2
220               End If
                  
230               If .TileHeight <> 1 Then
240                   Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
250               End If
260           End If
                      
270           SourceRect.Left = .sX
280           SourceRect.Top = .sY
290           SourceRect.Right = SourceRect.Left + .pixelWidth
300           SourceRect.Bottom = SourceRect.Top + .pixelHeight
              
              'Draw
310           Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), _
                  SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
320       End With
330   Exit Sub

error:
340       If Err.number = 9 And Grh.FrameCounter < 1 Then
350           Grh.FrameCounter = 1
360           Resume
370       Else
380           MsgBox _
                  "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." _
                  & vbcrlf & "Descripción del error: " & vbcrlf & Err.Description, _
                  vbExclamation, "[ " & Err.number & " ] Error"
390           End
400       End If
End Sub

'Sub DDrawTransGrhtoSurfaceAlpha(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, _
                        ByVal center As Byte, ByVal Animate As Byte)
                        
Sub DDrawTransGrhtoSurfaceAlpha(Surface As DirectDrawSurface7, Grh As Grh, _
ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, ByVal Alpha As Byte)
'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

  
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / (Grh.Speed))
          
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = 1
              
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If

'Dim CurrentGrh As Grh
Dim iGrhIndex As Integer
'Dim destRect As RECT
Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2
Dim QuitarAnimacion As Boolean

If Grh.GrhIndex = 0 Then Exit Sub


'Figure out what frame to draw (always 1 if not animated)
If Grh.FrameCounter < 1 Then Grh.FrameCounter = 1
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX + IIf(X < 0, Abs(X), 0)
    .Top = GrhData(iGrhIndex).sY + IIf(Y < 0, Abs(Y), 0)
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With

'surface.BltFast X, Y, SurfaceDB.surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

Dim Src As DirectDrawSurface7
Dim rDest As RECT
Dim dArray() As Byte, sArray() As Byte
Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
Dim Modo As Long

Set Src = SurfaceDB.Surface(GrhData(iGrhIndex).FileNum)

Src.GetSurfaceDesc ddsdSrc
Surface.GetSurfaceDesc ddsdDest

With rDest
    .Left = X
    .Top = Y
    .Right = X + GrhData(iGrhIndex).pixelWidth
    .Bottom = Y + GrhData(iGrhIndex).pixelHeight
  
    If .Right > ddsdDest.lWidth Then
        .Right = ddsdDest.lWidth
    End If
    If .Bottom > ddsdDest.lHeight Then
        .Bottom = ddsdDest.lHeight
    End If
End With

' 0 -> 16 bits 555
' 1 -> 16 bits 565
' 2 -> 16 bits raro (Sin implementar)
' 3 -> 24 bits
' 4 -> 32 bits

If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H3E0 Then
    Modo = 555
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 565
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 565
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = 65280 And ddsdSrc.ddpfPixelFormat.lGBitMask = 65280 Then
    Modo = 565
Else
    'Modo = 2 '16 bits raro ?
    Surface.BltFast X, Y, Src, SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Exit Sub
End If

Dim SrcLock As Boolean, DstLock As Boolean
SrcLock = False: DstLock = False

On Local Error GoTo HayErrorAlpha

Src.Lock SourceRect, ddsdSrc, DDLOCK_WAIT, 0
SrcLock = True
Surface.Lock rDest, ddsdDest, DDLOCK_WAIT, 0
DstLock = True

Surface.GetLockedArray dArray()
Src.GetLockedArray sArray()

Call vbDABLalphablend16(Modo, 1, ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), ByVal VarPtr(dArray(X + X, Y)), Alpha, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, 0)

Surface.Unlock rDest
DstLock = False
Src.Unlock SourceRect
SrcLock = False


Exit Sub

HayErrorAlpha:
If SrcLock Then Src.Unlock SourceRect
If DstLock Then Surface.Unlock rDest

End Sub
Sub DDrawTransGrhtoSurfaceAlpha1(ByRef Grh As Grh, ByVal X As Integer, ByVal Y _
    As Integer, ByVal center As Byte, ByVal Animate As Byte)
      '*****************************************************************
      'Draws a GRH transparently to a X and Y position
      '*****************************************************************
          Dim CurrentGrhIndex As Integer
          Dim SourceRect As RECT
          
10        If Animate Then
20         If ConAlfaB = 1 Then
30            If Grh.Started = 1 Then
40                Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * _
                      GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
                  
50                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
60                    Grh.FrameCounter = (Grh.FrameCounter Mod _
                          GrhData(Grh.GrhIndex).NumFrames) + 1
                      
70                    If Grh.Loops <> INFINITE_LOOPS Then
80                        If Grh.Loops > 0 Then
90                            Grh.Loops = Grh.Loops - 1
100                       Else
110                           Grh.Started = 0
120                       End If
130                   End If
140               End If
150           End If
160       End If
170      End If
          
          'Figure out what frame to draw (always 1 if not animated)
180       CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
          
          'Center Grh over X,Y pos
190       If center Then
200           If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
210               X = X - Int(GrhData(CurrentGrhIndex).TileWidth * TilePixelWidth / _
                      2) + TilePixelWidth \ 2
220           End If
230           If GrhData(CurrentGrhIndex).TileHeight <> 1 Then
240               Y = Y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) _
                      + TilePixelHeight
250           End If
260       End If
          
270       With SourceRect
280           .Left = GrhData(CurrentGrhIndex).sX
290           .Top = GrhData(CurrentGrhIndex).sY
300           .Right = .Left + GrhData(CurrentGrhIndex).pixelWidth
310           .Bottom = .Top + GrhData(CurrentGrhIndex).pixelHeight
320       End With
          
          Dim Src As DirectDrawSurface7
          Dim rDest As RECT
          Dim dArray() As Byte, sArray() As Byte
          Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
          Dim Modo As Long
          
330       Set Src = SurfaceDB.Surface(GrhData(CurrentGrhIndex).FileNum)
          
340       Src.GetSurfaceDesc ddsdSrc
350       BackBufferSurface.GetSurfaceDesc ddsdDest
          
360       With rDest
370           .Left = X
380           .Top = Y
390           .Right = X + GrhData(CurrentGrhIndex).pixelWidth
400           .Bottom = Y + GrhData(CurrentGrhIndex).pixelHeight
              
410           If .Right > ddsdDest.lWidth Then
420               .Right = ddsdDest.lWidth
430           End If
440           If .Bottom > ddsdDest.lHeight Then
450               .Bottom = ddsdDest.lHeight
460           End If
470       End With
          
          ' 0 -> 16 bits 555
          ' 1 -> 16 bits 565
          ' 2 -> 16 bits raro (Sin implementar)
          ' 3 -> 24 bits
          ' 4 -> 32 bits
          
480       If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0& And _
              ddsdSrc.ddpfPixelFormat.lGBitMask = &H3E0& Then
490           Modo = 0
500       ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0& And _
              ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0& Then
510           Modo = 1
      'TODO : Revisar las máscaras de 24!! Quizás mirando el campo lRGBBitCount para diferenciar 24 de 32...
520       ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0& And _
              ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0& Then
530           Modo = 3
540       ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &HFF00& And _
              ddsdSrc.ddpfPixelFormat.lGBitMask = &HFF00& Then
550           Modo = 4
560       Else
              'Modo = 2 '16 bits raro ?
570           Call BackBufferSurface.BltFast(X, Y, Src, SourceRect, _
                  DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
580           Exit Sub
590       End If
          
          Dim SrcLock As Boolean
          Dim DstLock As Boolean
          
600       SrcLock = False
610       DstLock = False
          
620   On Local Error GoTo HayErrorAlpha
          
630       Call Src.Lock(SourceRect, ddsdSrc, DDLOCK_WAIT, 0)
640       SrcLock = True
650       Call BackBufferSurface.Lock(rDest, ddsdDest, DDLOCK_WAIT, 0)
660       DstLock = True
          
670       Call BackBufferSurface.GetLockedArray(dArray())
680       Call Src.GetLockedArray(sArray())
          
690       Call BltAlphaFast(ByVal VarPtr(dArray(X + X, Y)), ByVal _
              VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), rDest.Right - _
              rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, Modo)
          
700       BackBufferSurface.Unlock rDest
710       DstLock = False
720       Src.Unlock SourceRect
730       SrcLock = False
740   Exit Sub

HayErrorAlpha:
750       If SrcLock Then Src.Unlock SourceRect
760       If DstLock Then BackBufferSurface.Unlock rDest
End Sub

Function GetBitmapDimensions(ByVal BmpFile As String, ByRef bmWidth As Long, _
    ByRef bmHeight As Long)
      '*****************************************************************
      'Gets the dimensions of a bmp
      '*****************************************************************
          Dim BMHeader As BITMAPFILEHEADER
          Dim BINFOHeader As BITMAPINFOHEADER
          
10        Open BmpFile For Binary Access Read As #1
          
20        Get #1, , BMHeader
30        Get #1, , BINFOHeader
          
40        Close #1
          
50        bmWidth = BINFOHeader.biWidth
60        bmHeight = BINFOHeader.biHeight
End Function

Sub DrawGrhtoHdc(ByVal hdc As Long, ByVal GrhIndex As Integer, ByRef SourceRect _
    As RECT, ByRef destRect As RECT)
      '*****************************************************************
      'Draws a Grh's portion to the given area of any Device Context
      '*****************************************************************
10        Call SurfaceDB.Surface(GrhData(GrhIndex).FileNum).BltToDC(hdc, SourceRect, _
              destRect)
End Sub

Public Sub DrawTransparentGrhtoHdc(ByVal dsthdc As Long, ByVal dstX As Long, _
    ByVal dstY As Long, ByVal GrhIndex As Integer, ByRef SourceRect As RECT, ByVal _
    TransparentColor As Long)
      '**************************************************************
      'Author: Torres Patricio (Pato)
      'Last Modify Date: 12/22/2009
      'This method is SLOW... Don't use in a loop if you care about
      'speed!
      '*************************************************************
          Dim Color As Long
          Dim X As Long
          Dim Y As Long
          Dim srchdc As Long
          Dim Surface As DirectDrawSurface7
          
10        Set Surface = SurfaceDB.Surface(GrhData(GrhIndex).FileNum)
          
20        srchdc = Surface.GetDC
          
30        For X = SourceRect.Left To SourceRect.Right - 1
40            For Y = SourceRect.Top To SourceRect.Bottom - 1
50                Color = GetPixel(srchdc, X, Y)
                  
60                If Color <> TransparentColor Then
70                    Call SetPixel(dsthdc, dstX + (X - SourceRect.Left), dstY + (Y - _
                          SourceRect.Top), Color)
80                End If
90            Next Y
100       Next X
          
110       Call Surface.ReleaseDC(srchdc)
End Sub

Public Sub DrawImageInPicture(ByRef PictureBox As PictureBox, ByRef Picture As _
    StdPicture, ByVal X1 As Single, ByVal Y1 As Single, Optional Width1, Optional _
    Height1, Optional X2, Optional Y2, Optional Width2, Optional Height2)
      '**************************************************************
      'Author: Torres Patricio (Pato)
      'Last Modify Date: 12/28/2009
      'Draw Picture in the PictureBox
      '*************************************************************

10    Call PictureBox.PaintPicture(Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, _
          Height2)
End Sub



Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal _
    PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
      '**************************************************************
      'Author: Aaron Perkins
      'Last Modify Date: 8/14/2007
      'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
      'Renders everything to the viewport
      '**************************************************************
          Dim Y           As Long     'Keeps track of where on map we are
          Dim X           As Long     'Keeps track of where on map we are
          Dim ScreenMinY  As Integer  'Start Y pos on current screen
          Dim ScreenMaxY  As Integer  'End Y pos on current screen
          Dim ScreenMinX  As Integer  'Start X pos on current screen
          Dim ScreenMaxX  As Integer  'End X pos on current screen
          Dim MinY        As Integer  'Start Y pos on current map
          Dim MaxY        As Integer  'End Y pos on current map
          Dim MinX        As Integer  'Start X pos on current map
          Dim MaxX        As Integer  'End X pos on current map
          Dim ScreenX     As Integer  'Keeps track of where to place tile on screen
          Dim ScreenY     As Integer  'Keeps track of where to place tile on screen
          Dim minXOffset  As Integer
          Dim minYOffset  As Integer
          Dim PixelOffsetXTemp As Integer 'For centering grhs
          Dim PixelOffsetYTemp As Integer 'For centering grhs
          
          'Dim Fichedd As Integer
          
          'Figure out Ends and Starts of screen
10        ScreenMinY = tiley - HalfWindowTileHeight
20        ScreenMaxY = tiley + HalfWindowTileHeight
30        ScreenMinX = tilex - HalfWindowTileWidth
40        ScreenMaxX = tilex + HalfWindowTileWidth
          
50        MinY = ScreenMinY - TileBufferSize
60        MaxY = ScreenMaxY + TileBufferSize
70        MinX = ScreenMinX - TileBufferSize
80        MaxX = ScreenMaxX + TileBufferSize
          
          'Make sure mins and maxs are allways in map bounds
90        If MinY < XMinMapSize Then
100           minYOffset = YMinMapSize - MinY
110           MinY = YMinMapSize
120       End If
          
130       If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
          
140       If MinX < XMinMapSize Then
150           minXOffset = XMinMapSize - MinX
160           MinX = XMinMapSize
170       End If
          
180       If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
           
          'If we can, we render around the view area to make it smoother
190       If ScreenMinY > YMinMapSize Then
200           ScreenMinY = ScreenMinY - 1
210       Else
220           ScreenMinY = 1
230           ScreenY = 1
240       End If
          
250       If ScreenMaxY < YMaxMapSize Then ScreenMaxY = ScreenMaxY + 1
          
260       If ScreenMinX > XMinMapSize Then
270           ScreenMinX = ScreenMinX - 1
280       Else
290           ScreenMinX = 1
300           ScreenX = 1
310       End If
          
320       If ScreenMaxX < XMaxMapSize Then ScreenMaxX = ScreenMaxX + 1
          
          'Draw floor layer
330       For Y = ScreenMinY To ScreenMaxY
340           For X = ScreenMinX To ScreenMaxX
                  
                  'Layer 1 **********************************
350               Call DDrawGrhtoSurface(MapData(X, Y).Graphic(1), (ScreenX - 1) * _
                      TilePixelWidth + PixelOffsetX + TileBufferPixelOffsetX, (ScreenY - _
                      1) * TilePixelHeight + PixelOffsetY + TileBufferPixelOffsetY, 0, 1)
                  '******************************************
                  
                  'Layer 2 **********************************
               If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(2), (ScreenX - 1) * _
                      TilePixelWidth + PixelOffsetX + TileBufferPixelOffsetX, (ScreenY - _
                      1) * TilePixelHeight + PixelOffsetY + TileBufferPixelOffsetY, 0, 1)
                  '******************************************
                  
                    'Layer 2 **********************************
450               'If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
460                  ' Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(2), (ScreenX _
                          - 1) * TilePixelWidth, (ScreenY - 1) * _
                          TilePixelHeight, 1, 1)
470              ' End If
                  '******************************************
                  
360               ScreenX = ScreenX + 1
370           Next X
              
              'Reset ScreenX to original value and increment ScreenY
380           ScreenX = ScreenX - X + ScreenMinX
390           ScreenY = ScreenY + 1
400       Next Y
          
          'Draw floor layer 2
410       'ScreenY = minYOffset
420     '  For Y = MinY To MaxY
430        '   ScreenX = minXOffset
440      '     For X = MinX To MaxX
                  

                  
480      '         ScreenX = ScreenX + 1
490     '      Next X
500      '     ScreenY = ScreenY + 1
510     '  Next Y
          
          'Draw Transparent Layers
520       ScreenY = minYOffset
530       For Y = MinY To MaxY
540           ScreenX = minXOffset
550           For X = MinX To MaxX
560               PixelOffsetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffsetX
570               PixelOffsetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffsetY
                  
580               With MapData(X, Y)
                      'Object Layer **********************************
590                   If .ObjGrh.GrhIndex <> 0 Then
600                       Call DDrawTransGrhtoSurface(.ObjGrh, PixelOffsetXTemp, _
                              PixelOffsetYTemp, 1, 1)
610                   End If
                      '***********************************************
                      
                      
                      'Char layer ************************************
620                   If .CharIndex <> 0 Then
630                       Call CharRender(.CharIndex, PixelOffsetXTemp, _
                              PixelOffsetYTemp)
640                   End If
                      '*************************************************
                      
                      'Dibujado del daño en el Render
650                   If .Damage.Activated Then
660                   If TSetup.bGameCombat Then
670                      m_Damages.Draw X, Y, PixelOffsetXTemp + 17, PixelOffsetYTemp _
                             - -2
680                   End If
690                   End If
                      
700                   If ConAlfaB = 0 Then
                      'Layer 3 *****************************************
710                   If .Graphic(3).GrhIndex <> 0 Then
                          'Draw
720                       Call DDrawTransGrhtoSurface(.Graphic(3), PixelOffsetXTemp, _
                              PixelOffsetYTemp, 1, 1)
730                 End If
740               End If
750               If ConAlfaB = 1 Then
                  
760                               If .Graphic(3).GrhIndex = 735 Or _
                                      .Graphic(3).GrhIndex >= 6994 And _
                                      .Graphic(3).GrhIndex <= 7002 Then
770                   If Abs(UserPos.X - X) < 4 And (Abs(UserPos.Y - Y)) < 4 Then
780                   Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, .Graphic(3), PixelOffsetXTemp, _
                          PixelOffsetYTemp, 1, 1, 155)
790                   Else
800                   Call DDrawTransGrhtoSurface(.Graphic(3), PixelOffsetXTemp, _
                          PixelOffsetYTemp, 1, 1)
810                   End If
820                   Else
830                   If .Graphic(3).GrhIndex <> 0 Then
840                   Call DDrawTransGrhtoSurface(.Graphic(3), PixelOffsetXTemp, _
                          PixelOffsetYTemp, 1, 1)
850                   End If
860                   End If
870                   End If
                  
                      '************************************************
880               End With
                  
890               ScreenX = ScreenX + 1
900           Next X
910           ScreenY = ScreenY + 1
920       Next Y
          
          
930               If ConAlfaB = 1 Then
              'Draw blocked tiles and grid
940        ScreenY = minYOffset
950           For Y = MinY To MaxY
960               ScreenX = minXOffset
970               For X = MinX To MaxX
       
                      'Layer 4 **********************************
                   
980                   If MapData(X, Y).Graphic(4).GrhIndex Then
                          'Draw
990                           If alphaT <> 255 Then
1000                              Call _
                                      DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, _
                                      MapData(X, Y).Graphic(4), (ScreenX - 1) * _
                                      TilePixelWidth + PixelOffsetX, (ScreenY - 1) * _
                                      TilePixelHeight + PixelOffsetY, 1, 1, alphaT)
1010                          Else
1020                              Call DDrawTransGrhtoSurface(MapData(X, _
                                      Y).Graphic(4), (ScreenX - 1) * TilePixelWidth + _
                                      PixelOffsetX, (ScreenY - 1) * TilePixelHeight + _
                                      PixelOffsetY, 1, 1)
1030                          End If
1040                  End If
                      '**********************************
       
1050                  ScreenX = ScreenX + 1
1060              Next X
1070              ScreenY = ScreenY + 1
1080          Next Y
1090          End If

1100      If ConAlfaB = 0 Then
1110      If Not bTecho Then
              'Draw blocked tiles and grid
1120          ScreenY = minYOffset
1130          For Y = MinY To MaxY
1140              ScreenX = minXOffset
1150              For X = MinX To MaxX
                      
                      'Layer 4 **********************************
1160                  If MapData(X, Y).Graphic(4).GrhIndex Then
                          'Draw
1170                      Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), _
                              (ScreenX - 1) * TilePixelWidth + PixelOffsetX, (ScreenY - _
                              1) * TilePixelHeight + PixelOffsetY, 1, 1)
1180                  End If
                      '**********************************
                      
1190                  ScreenX = ScreenX + 1
1200              Next X
1210              ScreenY = ScreenY + 1
1220          Next Y
1230      End If
1240      End If
          
1250      If Iscombate = True Then
1260              Call RenderText(260, 260, "Modo Combate", vbRed, frmMain.font)
1270      End If
          
          If SeguroClanes = True And TieneClan = True Then
            Call RenderText(260, 300, "Clan activado", vbRed, frmMain.font)
          End If
1280      If frmMain.macrotrabajo Then
1290          RenderTextCentered 300, 260, "(Trabajando)", RGB(255, 255, 255), _
                  frmMain.font
1300      End If

1310      If frmMain.Check1.value = vbChecked Then
1320          EfectoNoche BackBufferSurface
1330      End If
            
            
            'If CharFichado > 0 Then
                'RenderTextCentered 500, 260, "Marcado: " & charlist(CharFichado).Nombre, RGB(150, 255, 0), frmMain.font
            'End If

          If UserPoints > 0 Then
             RenderText 260, 270, "Canjes: " & UserPoints, vbGreen, frmMain.font
          End If
End Sub

#End If
Public Function RenderSounds()
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 3/30/2008
      'Actualiza todos los sonidos del mapa.
      '**************************************************************
          
10        DoFogataFx
End Function

Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As _
    Integer) As Boolean
10        If GrhIndex > 0 Then
20            HayUserAbajo = charlist(UserCharIndex).Pos.X >= X - _
                  (GrhData(GrhIndex).TileWidth \ 2) And charlist(UserCharIndex).Pos.X <= _
                  X + (GrhData(GrhIndex).TileWidth \ 2) And charlist(UserCharIndex).Pos.Y _
                  >= Y - (GrhData(GrhIndex).TileHeight - 1) And _
                  charlist(UserCharIndex).Pos.Y <= Y
30        End If
End Function

#If Wgl = 0 Then
Sub LoadGraphics()
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero - complete rewrite
      'Last Modify Date: 11/03/2006
      'Initializes the SurfaceDB and sets up the rain rects
      '**************************************************************
          'New surface manager :D
10        Call SurfaceDB.Initialize(DirectDraw, ClientSetup.bUseVideo, DirGraficos, _
              ClientSetup.byMemory)
          
          'Set up te rain rects
20        RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0: _
              RLluvia(3).Top = 0
30        RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256: _
              RLluvia(3).Left = 384
40        RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384: _
              RLluvia(3).Right = 512
50        RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: _
              RLluvia(3).Bottom = 128
          
60        RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128: _
              RLluvia(7).Top = 128
70        RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256: _
              RLluvia(7).Left = 384
80        RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384: _
              RLluvia(7).Right = 512
90        RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: _
              RLluvia(7).Bottom = 256
End Sub

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal _
    setMainViewTop As Integer, ByVal setMainViewLeft As Integer, ByVal _
    setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal _
    setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal _
    setTileBufferSize As Integer, ByVal pixelsToScrollPerFrameX As Integer, _
    pixelsToScrollPerFrameY As Integer, ByVal engineSpeed As Single) As Boolean
      '***************************************************
      'Author: Aaron Perkins
      'Last Modification: 08/14/07
      'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
      'Creates all DX objects and configures the engine to start running.
      '***************************************************
          Dim SurfaceDesc As DDSURFACEDESC2
          Dim ddck As DDCOLORKEY
          
10        IniPath = App.path & "\Init\"
          
20        Movement_Speed = 1
          
          'Fill startup variables
30        MainViewTop = setMainViewTop
40        MainViewLeft = setMainViewLeft
50        TilePixelWidth = setTilePixelWidth
60        TilePixelHeight = setTilePixelHeight
70        WindowTileHeight = setWindowTileHeight
80        WindowTileWidth = setWindowTileWidth
90        TileBufferSize = setTileBufferSize
          
100       HalfWindowTileHeight = setWindowTileHeight \ 2
110       HalfWindowTileWidth = setWindowTileWidth \ 2
          
          'Compute offset in pixels when rendering tile buffer.
          'We diminish by one to get the top-left corner of the tile for rendering.
120       TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
130       TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)
          
140       engineBaseSpeed = engineSpeed
          
          'Set FPS value to 60 for startup
150       FPS = 60
160       FramesPerSecCounter = 60
          
170       MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
180       MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
190       MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
200       MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
          
210       MainViewWidth = TilePixelWidth * WindowTileWidth
220       MainViewHeight = TilePixelHeight * WindowTileHeight
          
          'Resize mapdata array
230       ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As _
              MapBlock
          
          'Set intial user position
240       UserPos.X = MinXBorder
250       UserPos.Y = MinYBorder
          
          'Set scroll pixels per frame
260       ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
270       ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
          
       
          'Set the view rect
280       With MainViewRect
290           .Left = MainViewLeft
300           .Top = MainViewTop
310           .Right = .Left + MainViewWidth
320           .Bottom = .Top + MainViewHeight
330       End With
          
          'Set the dest rect
340       With MainDestRect
350           .Left = TilePixelWidth * TileBufferSize - TilePixelWidth
360           .Top = TilePixelHeight * TileBufferSize - TilePixelHeight
370           .Right = .Left + MainViewWidth
380           .Bottom = .Top + MainViewHeight
390       End With
          
400   On Error Resume Next
410       Set DirectX = New DirectX7
          
420       If Err Then
430           MsgBox _
                  "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
440           Exit Function
450       End If

          
          '****** INIT DirectDraw ******
          ' Create the root DirectDraw object
460       Set DirectDraw = DirectX.DirectDrawCreate("")
          
470       If Err Then
480           MsgBox _
                  "No se puede iniciar DirectDraw. Por favor asegurese de tener la ultima version correctamente instalada."
490           Exit Function
500       End If
          
510   On Error GoTo 0
520       Call DirectDraw.SetCooperativeLevel(setDisplayFormhWnd, DDSCL_NORMAL)
          
          'Primary Surface
          ' Fill the surface description structure
530       With SurfaceDesc
540           .lFlags = DDSD_CAPS
550           .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
560       End With
          ' Create the surface
570       Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)
          
          'Create Primary Clipper
580       Set PrimaryClipper = DirectDraw.CreateClipper(0)
590       Call PrimaryClipper.SetHWnd(frmMain.hWnd)
600       Call PrimarySurface.SetClipper(PrimaryClipper)
          
610       With BackBufferRect
620           .Left = 0
630           .Top = 0
640           .Right = TilePixelWidth * (WindowTileWidth + 2 * TileBufferSize)
650           .Bottom = TilePixelHeight * (WindowTileHeight + 2 * TileBufferSize)
660       End With
          
670       With SurfaceDesc
680           .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
690           If ClientSetup.bUseVideo Then
700               .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
710           Else
720               .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
730           End If
740           .lHeight = BackBufferRect.Bottom
750           .lWidth = BackBufferRect.Right
760       End With
          
          ' Create surface
770       Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)
          
          'Set color key
780       ddck.low = 0
790       ddck.high = 0
800       Call BackBufferSurface.SetColorKey(DDCKEY_SRCBLT, ddck)
          
          'Set font transparency
810       Call BackBufferSurface.SetFontTransparency(D_TRUE)
          
820       'Call LoadGrhData
830       'Call CargarCuerpos
840       'Call CargarCabezas
850       'If MD5File(App.path & "\Recursos\Clases\Clases.ind") <> _
              "b5617f8dc398c34ac499e31b0fffd874" Then
860           'End
870       'End If
880       'Call CargarCascos
890       'Call CargarFxs
          
900       LTLluvia(0) = 224
910       LTLluvia(1) = 352
920       LTLluvia(2) = 480
930       LTLluvia(3) = 608
940       LTLluvia(4) = 736
          
950       Call LoadGraphics
          
          initAntiCheat
960       InitTileEngine = True
End Function

Public Sub DeinitTileEngine()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 08/14/07
      'Destroys all DX objects
      '***************************************************
10    On Error Resume Next
20        Set PrimarySurface = Nothing
30        Set PrimaryClipper = Nothing
40        Set BackBufferSurface = Nothing
          
50        Set DirectDraw = Nothing
          
60        Set DirectX = Nothing
End Sub


Sub ShowNextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As _
    Integer, ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)
      '***************************************************
      'Author: Arron Perkins
      'Last Modification: 08/14/07
      'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
      'Updates the game's model and renders everything.
      '***************************************************

    
          '****** Set main view rectangle ******
10        MainViewRect.Left = (DisplayFormLeft / Screen.TwipsPerPixelX) + MainViewLeft
20        MainViewRect.Top = (DisplayFormTop / Screen.TwipsPerPixelY) + MainViewTop
30        MainViewRect.Right = MainViewRect.Left + MainViewWidth
40        MainViewRect.Bottom = MainViewRect.Top + MainViewHeight
          
50        If EngineRun Then
60            If UserMoving Then
                  '****** Move screen Left and Right if needed ******
70                If AddtoUserPos.X <> 0 Then
80                    If Not UserMontando Then
90                    OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * _
                          AddtoUserPos.X * timerTicksPerFrame
100               Else
110                   OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * _
                          AddtoUserPos.X * timerTicksPerFrame * Velocidad / 10 _
                          'Reemplacen este 10 por el valor por el que quieran dividir la _
                          vel
120               End If
130                   If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) _
                          Then
140                       OffsetCounterX = 0
150                       AddtoUserPos.X = 0
160                       UserMoving = False
170                   End If
180               End If
                  
                  '****** Move screen Up and Down if needed ******
190               If AddtoUserPos.Y <> 0 Then
200                   If Not UserMontando Then
210                   OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * _
                          AddtoUserPos.Y * timerTicksPerFrame
220               Else
230                   OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * _
                          AddtoUserPos.Y * timerTicksPerFrame * Velocidad / 10 _
                          'Reemplacen este 10 por el valor por el que quieran dividir la _
                          vel
240               End If
250                   If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) _
                          Then
260                       OffsetCounterY = 0
270                       AddtoUserPos.Y = 0
280                       UserMoving = False
290                   End If
300               End If
310           End If
              
              'Update mouse position within view area
320           Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
              
              '****** Update screen ******
330           If UserCiego Then
340               Call CleanViewPort
350           Else
360               Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - _
                      AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
370           End If

              'If IScombate Then Call RenderText(170, 165, "Modo Combate", vbRed, frmMain.font)
380           If ClientSetup.bActive Then
              
390               If isCapturePending Then
400                   Call ScreenCapture(True)
410                   isCapturePending = False
420               End If
430           End If


440           Call Dialogos.Render
460           Call DialogosClanes.Draw
              'Display front-buffer!
470           Call PrimarySurface.Blt(MainViewRect, BackBufferSurface, MainDestRect, _
                  DDBLT_WAIT)
              
              'Limit FPS to 100 (an easy number higher than monitor's vertical refresh rates)
             Dim TimeSleep As Long ' Declaramos el tiempo para pausear la PC
480          If TSetup.bFPS Then
490           TimeSleep = (1000 / 60 - 1) - SetElapsedTime(False) ' Cambiar numero 100 para cambiar los FPS
500           Else
510           TimeSleep = (1000 / 100 - 1) - SetElapsedTime(False)
520           End If
530           If TimeSleep > 0 Then 'Si el Tiempo es negativo entonces no limitamos
540               Sleep TimeSleep
550           End If
560           Call SetElapsedTime(True) 'Seteamos el tiempo
              
              'FPS update
570           If fpsLastCheck + 1000 < DirectX.TickCount Then
580               FPS = FramesPerSecCounter
590               FramesPerSecCounter = 1
600               fpsLastCheck = DirectX.TickCount
610           Else
620               FramesPerSecCounter = FramesPerSecCounter + 1
630           End If
              
640            timerElapsedTime = GetElapsedTime()
650           timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
660   End If
End Sub


Public Sub EfectoNoche(ByRef Surface As DirectDrawSurface7)
          Dim dArray() As Byte
          Dim ddsdDest As DDSURFACEDESC2
          Dim Modo As Long
          Dim rRect As RECT
          
10        Surface.GetSurfaceDesc ddsdDest
          
20        With rRect
30            .Left = 0
40            .Top = 0
50            .Right = ddsdDest.lWidth
60            .Bottom = ddsdDest.lHeight
70        End With
          
80        If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
90            Modo = 0
100       ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
110           Modo = 1
120       Else
130           Modo = 2
140       End If
          
          Dim DstLock As Boolean
150       DstLock = False
          
160       On Local Error GoTo HayErrorAlpha
          
170       Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
180       DstLock = True
          
190       Surface.GetLockedArray dArray()
200       Call BltEfectoNoche(ByVal VarPtr(dArray(0, 0)), ddsdDest.lWidth, _
              ddsdDest.lHeight, ddsdDest.lPitch, Modo)
          
HayErrorAlpha:
210       If DstLock = True Then
220           Surface.Unlock rRect
230           DstLock = False
240       End If
End Sub

Public Sub RenderText(ByVal lngXPos As Integer, ByVal lngYPos As Integer, ByRef _
    strText As String, ByVal lngColor As Long, ByRef font As StdFont)
10        If strText <> "" Then
20            Call BackBufferSurface.SetForeColor(vbBlack)
30            Call BackBufferSurface.SetFont(font)
40            Call BackBufferSurface.DrawText(lngXPos - 2, lngYPos - 1, strText, _
                  False)
              
50            Call BackBufferSurface.SetForeColor(lngColor)
60            Call BackBufferSurface.DrawText(lngXPos, lngYPos, strText, False)
70        End If
End Sub

Public Sub RenderTextCentered(ByVal lngXPos As Integer, ByVal lngYPos As _
    Integer, ByRef strText As String, ByVal lngColor As Long, ByRef font As StdFont, _
    Optional ByVal bNegro As Boolean = True)
          Dim hdc As Long
          Dim ret As Size
         
10        If strText <> "" Then
20            Call BackBufferSurface.SetFont(font)
             
              'Get width of text once rendered
30            hdc = BackBufferSurface.GetDC()
40            Call GetTextExtentPoint32(hdc, strText, Len(strText), ret)
50            Call BackBufferSurface.ReleaseDC(hdc)
             
60            lngXPos = lngXPos - ret.cx \ 2
             
70        If bNegro Then
80                Call BackBufferSurface.SetForeColor(vbBlack)
90                Call BackBufferSurface.DrawText(lngXPos - 2, lngYPos - 1, strText, _
                      False)
100           End If
       
110           Call BackBufferSurface.SetForeColor(lngColor)
120           Call BackBufferSurface.DrawText(lngXPos, lngYPos, strText, False)
130       End If
End Sub

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, _
    ByVal PixelOffsetY As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 12/03/04
      'Draw char's to screen without offcentering them
      '***************************************************
          Dim moved As Boolean
          Dim Pos As Integer
          Dim Line As String
          Dim Color As Long
          'Dim fiched As Integer
          Dim lClan As String
          Dim lPos As Integer
          
10        With charlist(CharIndex)
20            If .Moving Then
                  'If needed, move left and right
30                If .scrollDirectionX <> 0 Then
40                    If Not UserMontando Then
50                    .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * _
                          Sgn(.scrollDirectionX) * timerTicksPerFrame
60                Else
70                    .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * _
                          Sgn(.scrollDirectionX) * timerTicksPerFrame * Velocidad / 10 _
                          'Reemplacen este 10 por el valor por el que quieran dividir la _
                          vel
80                End If
                      
                      'Start animations
      'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
90                    If .Body.Walk(.Heading).Speed > 0 Then _
                          .Body.Walk(.Heading).Started = 1
100                   .Arma.WeaponWalk(.Heading).Started = 1
110                   .Escudo.ShieldWalk(.Heading).Started = 1
                      
                      'Char moved
120                   moved = True
                      
                      'Check if we already got there
130                   If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                          (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
140                       .MoveOffsetX = 0
150                       .scrollDirectionX = 0
160                   End If
170               End If
                  
                  'If needed, move up and down
180               If .scrollDirectionY <> 0 Then
190                   If Not UserMontando Then
200                   .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * _
                          Sgn(.scrollDirectionY) * timerTicksPerFrame
210               Else
220                   .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * _
                          Sgn(.scrollDirectionY) * timerTicksPerFrame * Velocidad / 10
230               End If
                      
                      'Start animations
      'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
240                   If .Body.Walk(.Heading).Speed > 0 Then _
                          .Body.Walk(.Heading).Started = 1
250                   .Arma.WeaponWalk(.Heading).Started = 1
260                   .Escudo.ShieldWalk(.Heading).Started = 1
                      
                      'Char moved
270                   moved = True
                      
                      'Check if we already got there
280                   If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                          (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
290                       .MoveOffsetY = 0
300                       .scrollDirectionY = 0
310                   End If
320               End If
330           End If
              
             'If done moving stop animation
340          If Not moved Then
350               If .Heading < 1 Or .Heading > 4 Then .Heading = EAST
                  
                  'Stop animations
360               .Body.Walk(.Heading).Started = 0
370               .Body.Walk(.Heading).FrameCounter = 1
                 
380               If Not .Movimient Then
                 
390               .Arma.WeaponWalk(.Heading).Started = 0
400               .Arma.WeaponWalk(.Heading).FrameCounter = 1
                 
410               .Escudo.ShieldWalk(.Heading).Started = 0
420               .Escudo.ShieldWalk(.Heading).FrameCounter = 1
                 
430               End If
                 
               .Moving = False
           End If
             
              
           PixelOffsetX = PixelOffsetX + .MoveOffsetX
           PixelOffsetY = PixelOffsetY + .MoveOffsetY
       
       
            If .Head.Head(.Heading).GrhIndex Then
                If Not .Invisible Then
                    If ConAlfaB = 0 Then
                      Movement_Speed = 0.5
                                      'Draw Body
                      If .Body.Walk(.Heading).GrhIndex Then
                         Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), _
                                PixelOffsetX, PixelOffsetY, 1, 1)
                      End If
                         
                        'Draw Head
                     If .Head.Head(.Heading).GrhIndex Then
            
                         Call DDrawTransGrhtoSurface(.Head.Head(.Heading), _
                                PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + _
                                .Body.HeadOffset.Y, 1, 0)
                            
                            'Draw Helmet
                          If .Casco.Head(.Heading).GrhIndex Then
                          
                              If .Casco.Head(.Heading).GrhIndex >= 24738 And .Casco.Head(.Heading).GrhIndex <= 24741 Then
                              
                                  Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX _
                                  + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + 13, 1, _
                                  0)
                                  
                              Else
                              
                                  Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX _
                                  + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, _
                                  0)
                              End If
                          End If
                            
                            'Draw Weapon
                         If .Arma.WeaponWalk(.Heading).GrhIndex Then Call _
                                DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), _
                                PixelOffsetX, PixelOffsetY, 1, 1)
                            
                            'Draw Shield
                         If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call _
                                DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), _
                                PixelOffsetX, PixelOffsetY, 1, 1)
                        
                            'Draw name over head
                         If LenB(.Nombre) > 0 Then
                              If Nombres Then
                                   Pos = getTagPosition(.Nombre)
                                      'Pos = InStr(.Nombre, "<")
                                      'If Pos = 0 Then Pos = Len(.Nombre) + 2
                                      
                                   If .priv = 0 Then
                                           If .Criminal Or .Team = 1 Then
                                               Color = RGB(ColoresPJ(50).r, _
                                                      ColoresPJ(50).g, ColoresPJ(50).b)
                                           ElseIf Not .Criminal Or .Team = 2 Then
                                               Color = RGB(ColoresPJ(49).r, _
                                                     ColoresPJ(49).g, ColoresPJ(49).b)
                                           End If
                                   Else
                                       Color = RGB(ColoresPJ(.priv).r, _
                                              ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                                   End If
                                    
                                                                              
                                     
                                        'Nick
                                        Line = Left$(.Nombre, Pos - 2)
                                        If (UCase(Line) = "NEGRA") Then Color = NegraColour()
                                        Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 30, Line, Color, frmMain.font)
                                        
                                   Call RenderTextCentered(PixelOffsetX + _
                                          TilePixelWidth \ 2 + 5, PixelOffsetY + 30, Line, _
                                          Color, frmMain.font)
                                      
                                      'Clan
                                   Line = mid$(.Nombre, Pos)
                                   Call RenderTextCentered(PixelOffsetX + _
                                          TilePixelWidth \ 2 + 5, PixelOffsetY + 45, Line, _
                                          Color, frmMain.font)
                                  
                                    
                              End If
                          End If 'LENB(.NOMBRE) > 0
                      End If 'Head <> 0
                    End If 'ALPHAB = 0
            
            End If 'IF NOT INVISIBLE
        End If 'IF .HEAD.GRHINDEX
        
    If LenB(.Nombre) > 0 Then 'INVI ENTRE MIEMBROS DE CLAN, Y NOMBRE EN BLANCO CUANDO ESTA INVISIBLE TAMBIEN
        If Nombres Then
            
            Pos = getTagPosition(.Nombre)
            Line = mid$(.Nombre, Pos) 'ClanName
            
            lPos = getTagPosition(charlist(UserCharIndex).Nombre)
            lClan = mid$(charlist(UserCharIndex).Nombre, lPos) 'ClanName
            
            If (UCase$(Line) = UCase$(lClan) And LenB(Line) > 3) Or LenB(lClan) = 0 Then 'SOn del mismo clan

                Color = RGB(210, 210, 210)
                Line = Left$(.Nombre, Pos - 2)
                'Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 30, Line, Color, frmMain.font)
                
                 Call RenderTextCentered(PixelOffsetX + _
                        TilePixelWidth \ 2 + 5, PixelOffsetY + 30, Line, _
                        Color, frmMain.font)
                    
                    'Clan
                 Line = mid$(.Nombre, Pos)
                 Call RenderTextCentered(PixelOffsetX + _
                        TilePixelWidth \ 2 + 5, PixelOffsetY + 45, Line, _
                        Color, frmMain.font)
                
            End If
              
        End If
    End If 'LENB(.NOMBRE) > 0
 If .Head.Head(.Heading).GrhIndex Then
          If Not .Invisible Then
               If ConAlfaB = 1 Then
850               Movement_Speed = 0.5
                  
860      If .Body.Walk(.Heading).GrhIndex Then
870      If .iBody <> 8 And .iBody <> 145 Then
880                       Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), _
                              PixelOffsetX, PixelOffsetY, 1, 1)
890                   Else
900                       DDrawTransGrhtoSurfaceAlpha BackBufferSurface, .Body.Walk(.Heading), _
                              PixelOffsetX, PixelOffsetY, 1, 1, 155
910                    End If
920                    End If
                       
                      'Draw Head
930                   If .Head.Head(.Heading).GrhIndex Then
940                   If .iHead <> 500 And .iHead <> 501 Then
950                       Call DDrawTransGrhtoSurface(.Head.Head(.Heading), _
                              PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + _
                              .Body.HeadOffset.Y, 1, 0)
960                      Else
970                      DDrawTransGrhtoSurfaceAlpha BackBufferSurface, .Head.Head(.Heading), _
                             PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + _
                             .Body.HeadOffset.Y, 1, 0, 155
980                      End If

                          'Draw Helmet
990                       If .Casco.Head(.Heading).GrhIndex Then

                            If .Casco.Head(.Heading).GrhIndex >= 24738 And .Casco.Head(.Heading).GrhIndex <= 24741 Then
                                Call _
                              DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX _
                              + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + 13, 1, _
                              0)
                            Else
                                Call _
                              DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX _
                              + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, _
                              0)
                              
                            End If
                          End If
                          
                          'Draw Weapon
1000                      If .Arma.WeaponWalk(.Heading).GrhIndex Then Call _
                              DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), _
                              PixelOffsetX, PixelOffsetY, 1, 1)
                          
                          'Draw Shield
1010                      If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call _
                              DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), _
                              PixelOffsetX, PixelOffsetY, 1, 1)
                      
                          'Draw name over head
1020                      If LenB(.Nombre) > 0 Then
1030                          If Nombres Then
1040                              Pos = getTagPosition(.Nombre)
                                  'Pos = InStr(.Nombre, "<")
                                  'If Pos = 0 Then Pos = Len(.Nombre) + 2
                                  
1050                              If .priv = 0 Then
1060                                  If .Team > 0 Then
1070                                      Select Case .Team
                                              Case 1
1080                                              Color = RGB(57, 255, 0)
1090                                          Case 2
1100                                              Color = RGB(224, 255, 37)
1110                                      End Select
1120                                  Else
1130                                      If .Criminal Then
1140                                          Color = RGB(ColoresPJ(50).r, _
                                                  ColoresPJ(50).g, ColoresPJ(50).b)
1150                                      Else
1160                                          Color = RGB(ColoresPJ(49).r, _
                                                  ColoresPJ(49).g, ColoresPJ(49).b)
1170                                      End If
1180                                  End If
1190                              Else
1200                                  Color = RGB(ColoresPJ(.priv).r, _
                                          ColoresPJ(.priv).g, ColoresPJ(.priv).b)
1210                              End If
                                
                                                                          
                                 
                                  'Nick
1220                             Line = Left$(.Nombre, Pos - 2)
1230                              Call RenderTextCentered(PixelOffsetX + _
                                      TilePixelWidth \ 2 + 5, PixelOffsetY + 30, Line, _
                                      Color, frmMain.font)
                                  
                                  'Clan
1240                              Line = mid$(.Nombre, Pos)
1250                              Call RenderTextCentered(PixelOffsetX + _
                                      TilePixelWidth \ 2 + 5, PixelOffsetY + 45, Line, _
                                      Color, frmMain.font)
                                 
1260                    End If
1270             End If
1280             End If
1290             End If
                      
1300              ElseIf .Invisible Then
1310              If ConAlfaB = 1 Then
1320              If esGM(UserCharIndex) = True Then Exit Sub
1330              If UserOcu = 1 Then Exit Sub
1340               If .cCont > 0 Then .cCont = .cCont - 1
1350             If .cCont = 0 Then
1360              .Drawers = .Drawers + 1
1370              If .Drawers = 156 Then .Drawers = 0: .cCont = 400
                  
1380                     If .Body.Walk(.Heading).GrhIndex Then Call _
                             DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, .Body.Walk(.Heading), _
                             PixelOffsetX, PixelOffsetY, 1, 1, 155)
                  
                      'Draw Head
1390                  If .Head.Head(.Heading).GrhIndex Then
1400                      Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, .Head.Head(.Heading), _
                              PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + _
                              .Body.HeadOffset.Y, 1, 0, 155)
                          
                          
                          'Draw Helmet
1410                      If .Casco.Head(.Heading).GrhIndex Then
                              If .Casco.Head(.Heading).GrhIndex >= 24738 And .Casco.Head(.Heading).GrhIndex <= 24741 Then
                                Call _
                              DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, .Casco.Head(.Heading), _
                              PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + _
                              .Body.HeadOffset.Y + 13, 1, 0, 155)
                              
                              Else
                                Call _
                              DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, .Casco.Head(.Heading), _
                              PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + _
                              .Body.HeadOffset.Y, 1, 0, 155)
                              
                                End If
                          
                          End If
                          'Draw Weapon
1420                      If .Arma.WeaponWalk(.Heading).GrhIndex Then Call _
                              DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, .Arma.WeaponWalk(.Heading), _
                              PixelOffsetX, PixelOffsetY, 1, 1, 155)
                          
                          'Draw Shield
1430                      If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call _
                              DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, .Escudo.ShieldWalk(.Heading), _
                              PixelOffsetX, PixelOffsetY, 1, 1, 155)
                      
1440                              End If
1450                              End If
                              

                         
1460                        ElseIf ConAlfaB = 0 Then
1470                         If .cCont > 0 Then .cCont = .cCont - 1
1480                         If esGM(UserCharIndex) = True Then Exit Sub
1490                         If UserOcu = 1 Then Exit Sub
1500             If .cCont = 0 Then
1510              .Drawers = .Drawers + 1
1520              If .Drawers = 156 Then .Drawers = 0: .cCont = 400
1530                        If .Body.Walk(.Heading).GrhIndex Then Call _
                                DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, _
                                PixelOffsetY, 1, 1)
                  
                      'Draw Head
1540                  If .Head.Head(.Heading).GrhIndex Then
1550                      Call DDrawTransGrhtoSurface(.Head.Head(.Heading), _
                              PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + _
                              .Body.HeadOffset.Y, 1, 0)
                          
                          'Draw Helmet
1560                      If .Casco.Head(.Heading).GrhIndex Then

                            If .Casco.Head(.Heading).GrhIndex >= 24738 And .Casco.Head(.Heading).GrhIndex <= 24741 Then
                                Call _
                              DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX _
                              + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + 13, 1, _
                              0)
                              
                            Else
                                Call _
                              DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX _
                              + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, _
                              0)
                            
                            End If
                          End If
                          
                          'Draw Weapon
1570                      If .Arma.WeaponWalk(.Heading).GrhIndex Then Call _
                              DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), _
                              PixelOffsetX, PixelOffsetY, 1, 1)
                          
                          'Draw Shield
1580                      If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call _
                              DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), _
                              PixelOffsetX, PixelOffsetY, 1, 1)
1590              End If
1600              End If
1610              End If
1620          End If
              
1630                  Else
                  'Draw Body
1640              If .Body.Walk(.Heading).GrhIndex Then Call _
                      DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, _
                      PixelOffsetY, 1, 1)
1650          End If
              
1660          If .Infected = 1 Then
           'Nick infectado
1670       Line = Left$(.Nombre, Pos - 2)
1680       Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, _
               PixelOffsetY + 30, Line, RGB(255, 255, 255), frmMain.font)
                                 
           'Clan infectado
1690       Line = mid$(.Nombre, Pos)
1700       Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, _
               PixelOffsetY + 45, Line, RGB(255, 255, 255), frmMain.font)
             
           'Tagged Infected for user
1710       Line = "      [EVENT BOSS]"
1720       Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, _
               PixelOffsetY + 60, Line, RGB(255, 255, 255), frmMain.font)
1730       End If
              
1740                  If .Angel = 1 Then
           'Nick infectado
1750       Line = Left$(.Nombre, Pos - 2)
1760       Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, _
               PixelOffsetY + 30, Line, RGB(255, 255, 255), frmMain.font)
                                 
           'Clan infectado
1770       Line = mid$(.Nombre, Pos)
1780       Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, _
               PixelOffsetY + 45, Line, RGB(255, 255, 255), frmMain.font)
             
           'Tagged Infected for user
1790       Line = "      [EVENT BOSS]"
1800       Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, _
               PixelOffsetY + 90, Line, RGB(255, 255, 255), frmMain.font)
1810       End If
           
1820               If .Demonio = 1 Then
           'Nick infectado
1830       Line = Left$(.Nombre, Pos - 2)
1840       Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, _
               PixelOffsetY + 30, Line, RGB(255, 255, 255), frmMain.font)
                                 
           'Clan infectado
1850       Line = mid$(.Nombre, Pos)
1860       Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, _
               PixelOffsetY + 45, Line, RGB(255, 255, 255), frmMain.font)
             
           'Tagged Infected for user
1870       Line = "      [EVENT BOSS]"
1880       Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, _
               PixelOffsetY + 60, Line, RGB(255, 255, 255), frmMain.font)
1890       End If
              
              
                'Update dialogs
1900          Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X, _
                  PixelOffsetY + .Body.HeadOffset.Y, CharIndex)  '34 son los pixeles del _
                  grh de la cabeza que quedan superpuestos al cuerpo
1910          Movement_Speed = 1
              
                     'Draw FX
1920          If .FxIndex <> 0 Then
      Dim XDATAFX As Integer, YDATAFX As Integer
              
                  '@Nota de Dunkan: Arreglar desde el INDICE.
1930              If .FxIndex = 1 Then
1940                  XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
1950                  YDATAFX = PixelOffsetY + 25
1960              ElseIf .FxIndex = 18 Then
1970                  XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
1980                  YDATAFX = PixelOffsetY - 15
1990              ElseIf .FxIndex = 17 Then
2000                  XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
2010                  YDATAFX = PixelOffsetY - 15
2020              ElseIf .FxIndex = 19 Then
2030                  XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
2040                  YDATAFX = PixelOffsetY + 25
2050              ElseIf .FxIndex = 7 Then    'TORMENTA DE FUEGO
2060                  XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
2070                  YDATAFX = PixelOffsetY + 30
2080              ElseIf .FxIndex = 8 Then    'PARALIZAR
2090                  XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
2100                  YDATAFX = PixelOffsetY + 35
2110              ElseIf .FxIndex = 9 Then    'CURAR GRAVES
2120                  XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
2130                  YDATAFX = PixelOffsetY + 25
2140              ElseIf .FxIndex = 12 Then   'INMO
2150                  XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
2160                  YDATAFX = PixelOffsetY + 20
2170              Else
2180                  XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
2190                  YDATAFX = PixelOffsetY
2200              End If
2210              If (ConAlfaB = 1) Then
2220              DDrawTransGrhtoSurfaceAlpha BackBufferSurface, .fX, XDATAFX, YDATAFX, 1, 1, 155
2230              Else
2240               Call DDrawTransGrhtoSurface(.fX, XDATAFX, YDATAFX, 1, 1)
2250             End If
                    
                  'Check if animation is over
2260              If .fX.Started = 0 Then .FxIndex = 0
                    

2270          End If
2280      End With
End Sub


Sub DDrawTransGrhtoSurfaceAlphaTecho(Surface As DirectDrawSurface7, Grh As Grh, _
    ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, ByVal _
    Alpha As Byte)
      '[END]'
      '*****************************************************************
      'Draws a GRH transparently to a X and Y position
      '*****************************************************************
      '[CODE]:MatuX
      '
      '  CurrentGrh.GrhIndex = iGrhIndex
      '
      '[END]
       
       
10        If Animate Then
20            If Grh.Started = 1 Then
30            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * _
                  GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * Movement_Speed
               
40                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
50                    Grh.FrameCounter = 1
                   
60                    If Grh.Loops <> INFINITE_LOOPS Then
70                        If Grh.Loops > 0 Then
80                            Grh.Loops = Grh.Loops - 1
90                        Else
100                           Grh.Started = 0
110                       End If
120                   End If
130               End If
140           End If
150       End If
       
      'Dim CurrentGrh As Grh
      Dim iGrhIndex As Integer
      'Dim destRect As RECT
      Dim SourceRect As RECT
      'Dim SurfaceDesc As DDSURFACEDESC2
      Dim QuitarAnimacion As Boolean
       
160   If Grh.GrhIndex = 0 Then Exit Sub
       
       
      'Figure out what frame to draw (always 1 if not animated)
170   If Grh.FrameCounter < 1 Then Grh.FrameCounter = 1
180   iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
       
      'Center Grh over X,Y pos
190   If center Then
200       If GrhData(iGrhIndex).TileWidth <> 1 Then
210           X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
220       End If
230       If GrhData(iGrhIndex).TileHeight <> 1 Then
240           Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
250       End If
260   End If
       
270   With SourceRect
280       .Left = GrhData(iGrhIndex).sX + IIf(X < 0, Abs(X), 0)
290       .Top = GrhData(iGrhIndex).sY + IIf(Y < 0, Abs(Y), 0)
300       .Right = .Left + GrhData(iGrhIndex).pixelWidth
310       .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
320   End With
       
      'surface.BltFast X, Y, SurfaceDB.surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
       
      Dim Src As DirectDrawSurface7
      Dim rDest As RECT
      Dim dArray() As Byte, sArray() As Byte
      Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
      Dim Modo As Long
       
330   Set Src = SurfaceDB.Surface(GrhData(iGrhIndex).FileNum)
       
340   Src.GetSurfaceDesc ddsdSrc
350   Surface.GetSurfaceDesc ddsdDest
       
360   With rDest
370       .Left = X
380       .Top = Y
390       .Right = X + GrhData(iGrhIndex).pixelWidth
400       .Bottom = Y + GrhData(iGrhIndex).pixelHeight
       
410       If .Right > ddsdDest.lWidth Then
420           .Right = ddsdDest.lWidth
430       End If
440       If .Bottom > ddsdDest.lHeight Then
450           .Bottom = ddsdDest.lHeight
460       End If
470   End With
       
      ' 0 -> 16 bits 555
      ' 1 -> 16 bits 565
      ' 2 -> 16 bits raro (Sin implementar)
      ' 3 -> 24 bits
      ' 4 -> 32 bits
       
480   If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 And _
          ddsdSrc.ddpfPixelFormat.lGBitMask = &H3E0 Then
490       Modo = 555
500   ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And _
          ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
510       Modo = 565
520   ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And _
          ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
530       Modo = 565
540   ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = 65280 And _
          ddsdSrc.ddpfPixelFormat.lGBitMask = 65280 Then
550       Modo = 565
560   Else
          'Modo = 2 '16 bits raro ?
570       Surface.BltFast X, Y, Src, SourceRect, DDBLTFAST_SRCCOLORKEY Or _
              DDBLTFAST_WAIT
580       Exit Sub
590   End If
       
      Dim SrcLock As Boolean, DstLock As Boolean
600   SrcLock = False: DstLock = False
       
610   On Local Error GoTo HayErrorAlpha
       
620   Src.Lock SourceRect, ddsdSrc, DDLOCK_WAIT, 0
630   SrcLock = True
640   Surface.Lock rDest, ddsdDest, DDLOCK_WAIT, 0
650   DstLock = True
       
660   Surface.GetLockedArray dArray()
670   Src.GetLockedArray sArray()
       
680   Call vbDABLalphablend16(Modo, 1, ByVal VarPtr(sArray(SourceRect.Left * 2, _
          SourceRect.Top)), ByVal VarPtr(dArray(X + X, Y)), Alpha, rDest.Right - _
          rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, 0)
       
690   Surface.Unlock rDest
700   DstLock = False
710   Src.Unlock SourceRect
720   SrcLock = False
       
       
730   Exit Sub
       
HayErrorAlpha:
740   If SrcLock Then Src.Unlock SourceRect
750   If DstLock Then Surface.Unlock rDest
       
End Sub

#End If

Public Function SetElapsedTime(ByVal Start As Boolean) As Single
      '**************************************************************
      'Author: Aaron Perkins
      'Last Modify Date: 23/05/2011 By MaTeO
      'Gets the time that past since the last call
      '[MaTeO] Agrego cambios a la funcion
      '**************************************************************
          Dim start_time As Currency
          Static end_time As Currency
          Static timer_freq As Currency
          'Get the timer frequency
10        If timer_freq = 0 Then
20            QueryPerformanceFrequency timer_freq
30        End If
         
          'Get current time
40        Call QueryPerformanceCounter(start_time)
         
50        If Not Start Then
              'Calculate elapsed time
60            SetElapsedTime = (start_time - end_time) / timer_freq * 1000
         
              'Get next end time
70        Else
80            Call QueryPerformanceCounter(end_time)
90        End If
End Function

Private Function GetElapsedTime() As Single
      '**************************************************************
      'Author: Aaron Perkins
      'Last Modify Date: 10/07/2002
      'Gets the time that past since the last call
      '**************************************************************
          Dim start_time As Currency
          Static end_time As Currency
          Static timer_freq As Currency

          'Get the timer frequency
10        If timer_freq = 0 Then
20            QueryPerformanceFrequency timer_freq
30        End If
          
          'Get current time
40        Call QueryPerformanceCounter(start_time)
          
          'Calculate elapsed time
50        GetElapsedTime = (start_time - end_time) / timer_freq * 1000
          
          'Get next end time
60        Call QueryPerformanceCounter(end_time)
End Function

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, _
    ByVal Loops As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 12/03/04
      'Sets an FX to the character.
      '***************************************************
10        With charlist(CharIndex)
20            .FxIndex = fX
              
30            If .FxIndex > 0 Then
40                Call InitGrh(.fX, FxData(fX).Animacion)
              
50                .fX.Loops = Loops
60            End If
70        End With
End Sub

#If Wgl = 0 Then
Private Sub CleanViewPort()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 12/03/04
      'Fills the viewport with black.
      '***************************************************
          Dim r As RECT
10        Call BackBufferSurface.BltColorFill(r, vbBlack)
End Sub
#End If

Public Function CharTieneClan() As Boolean

      Dim tPos As Integer

10    tPos = InStr(charlist(UserCharIndex).Nombre, "<")

20    If tPos = 0 Then
30    CharTieneClan = False
40    Exit Function
50    End If

60    CharTieneClan = True

End Function
Public Function CharTieneParty() As Boolean

      Dim tPos As Integer

10    tPos = InStr(charlist(UserCharIndex).Nombre, "<")

20    If tPos = 0 Then
30    CharTieneParty = False
40    Exit Function
50    End If

60    CharTieneParty = True

End Function

 
Private Function NegraColour() As Long

On Error Resume Next

NegraColour = RGB(255, 0, 255)

End Function
