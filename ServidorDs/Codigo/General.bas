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

   Public Expc As Integer
    Public Oroc As Integer

Global LeerNPCs As New clsIniManager


Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, Optional ByVal Mimetizado As Boolean = False)
      '***************************************************
      'Autor: Nacho (Integer)
      'Last Modification: 03/14/07
      'Da cuerpo desnudo a un usuario
      '23/11/2009: ZaMa - Optimizacion de codigo.
      '***************************************************

      Dim CuerpoDesnudo As Integer

10    With UserList(UserIndex)
20        Select Case .Genero
              Case eGenero.Hombre
30                Select Case .raza
                      Case eRaza.Humano
40                        CuerpoDesnudo = 21
50                    Case eRaza.Drow
60                        CuerpoDesnudo = 32
70                    Case eRaza.Elfo
80                        CuerpoDesnudo = 21
90                    Case eRaza.Gnomo
100                       CuerpoDesnudo = 53
110                   Case eRaza.Enano
120                       CuerpoDesnudo = 53
130               End Select
140           Case eGenero.Mujer
150               Select Case .raza
                      Case eRaza.Humano
160                       CuerpoDesnudo = 39
170                   Case eRaza.Drow
180                       CuerpoDesnudo = 40
190                   Case eRaza.Elfo
200                       CuerpoDesnudo = 39
210                   Case eRaza.Gnomo
220                       CuerpoDesnudo = 60
230                   Case eRaza.Enano
240                       CuerpoDesnudo = 60
250               End Select
260       End Select
          
270       If Mimetizado Then
280           .CharMimetizado.body = CuerpoDesnudo
290       Else
300           .Char.body = CuerpoDesnudo
310       End If
          
320       .flags.Desnudo = 1
330   End With

End Sub


Sub Bloquear(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal B As Boolean)
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

10        If toMap Then
20            Call SendData(SendTarget.toMap, sndIndex, PrepareMessageBlockPosition(X, Y, B))
30        Else
40            Call WriteBlockPosition(sndIndex, X, Y, B)
50        End If

End Sub


Function HayAgua(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If map > 0 And map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
20            With MapData(map, X, Y)
30                If ((.Graphic(1) >= 1505 And .Graphic(1) <= 1520) Or _
                  (.Graphic(1) >= 5665 And .Graphic(1) <= 5680) Or _
                  (.Graphic(1) >= 13547 And .Graphic(1) <= 13562)) And _
                     .Graphic(2) = 0 Then
40                        HayAgua = True
50                Else
60                        HayAgua = False
70                End If
80            End With
90        Else
100         HayAgua = False
110       End If

End Function

Private Function HayLava(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
      '***************************************************
      'Autor: Nacho (Integer)
      'Last Modification: 03/12/07
      '***************************************************
10        If map > 0 And map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
20            If MapData(map, X, Y).Graphic(1) >= 5837 And MapData(map, X, Y).Graphic(1) <= 5852 Then
30                HayLava = True
40            Else
50                HayLava = False
60            End If
70        Else
80          HayLava = False
90        End If

End Function
Function HayCura(ByVal UserIndex As Integer) As Boolean
      '******************************
      'Adaptacion a 13.0: Kaneidra
      'Last Modification: 15/05/2012
      '******************************
       
      Dim X As Integer, Y As Integer
       
10    For Y = UserList(UserIndex).Pos.Y - MinYBorder + 1 To UserList(UserIndex).Pos.Y + MinYBorder - 1
20    For X = UserList(UserIndex).Pos.X - MinXBorder + 1 To UserList(UserIndex).Pos.X + MinXBorder - 1
             
30                If MapData(UserList(UserIndex).Pos.map, X, Y).NpcIndex > 0 Then
40                        If Npclist(MapData(UserList(UserIndex).Pos.map, X, Y).NpcIndex).NPCtype = 1 Then
50                            If Distancia(UserList(UserIndex).Pos, Npclist(MapData(UserList(UserIndex).Pos.map, X, Y).NpcIndex).Pos) < 10 Then
60                            HayCura = True
70                            Exit Function
80                        End If
90                    End If
100               End If
                 
110           Next X
120   Next Y
       
130   HayCura = False
       
End Function

Sub EnviarSpawnList(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim k As Long
          Dim npcNames() As String
          
10        ReDim npcNames(1 To UBound(SpawnList)) As String
          
20        For k = 1 To UBound(SpawnList)
30            npcNames(k) = SpawnList(k).NpcName
40        Next k
          
50        Call WriteSpawnList(UserIndex, npcNames())

End Sub

Sub ConfigListeningSocket(ByRef Obj As Object, ByVal Port As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

#If UsarQueSocket = 0 Then

10        Obj.AddressFamily = AF_INET
20        Obj.Protocol = IPPROTO_IP
30        Obj.SocketType = SOCK_STREAM
40        Obj.Binary = False
50        Obj.Blocking = False
60        Obj.BufferSize = 1024
70        Obj.LocalPort = Port
80        Obj.backlog = 5
90        Obj.listen

#End If

End Sub

Sub Main()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error Resume Next
          Dim f As Date
          
20        ChDir App.Path
30        ChDrive App.Path
          
          SetInterval
          
40        Call BanIpCargar
50        Call LoadMAO
60        Call LoadDataDias
70        Call LoadRanking
80        Call BanHDCargar '//Disco
90        Call LoadCanjes
          Call LoadInvasiones
100       Call LoadArenas
          
110       GlobalActivado = 0
          
120       Prision.map = 66
130       Libertad.map = 66
          
140       Prision.X = 75
150       Prision.Y = 47
160       Libertad.X = 75
170       Libertad.Y = 65
          
          
180       LastBackup = Format(Now, "Short Time")
190       Minutos = Format(Now, "Short Time")
          
200       IniPath = App.Path & "\"
210       DatPath = App.Path & "\Dat\"
          
          
          Call LoadCofres
          
220       LevelSkill(1).LevelValue = 3
230       LevelSkill(2).LevelValue = 5
240       LevelSkill(3).LevelValue = 7
250       LevelSkill(4).LevelValue = 10
260       LevelSkill(5).LevelValue = 13
270       LevelSkill(6).LevelValue = 15
280       LevelSkill(7).LevelValue = 17
290       LevelSkill(8).LevelValue = 20
300       LevelSkill(9).LevelValue = 23
310       LevelSkill(10).LevelValue = 25
320       LevelSkill(11).LevelValue = 27
330       LevelSkill(12).LevelValue = 30
340       LevelSkill(13).LevelValue = 33
350       LevelSkill(14).LevelValue = 35
360       LevelSkill(15).LevelValue = 37
370       LevelSkill(16).LevelValue = 40
380       LevelSkill(17).LevelValue = 43
390       LevelSkill(18).LevelValue = 45
400       LevelSkill(19).LevelValue = 47
410       LevelSkill(20).LevelValue = 50
420       LevelSkill(21).LevelValue = 53
430       LevelSkill(22).LevelValue = 55
440       LevelSkill(23).LevelValue = 57
450       LevelSkill(24).LevelValue = 60
460       LevelSkill(25).LevelValue = 63
470       LevelSkill(26).LevelValue = 65
480       LevelSkill(27).LevelValue = 67
490       LevelSkill(28).LevelValue = 70
500       LevelSkill(29).LevelValue = 73
510       LevelSkill(30).LevelValue = 75
520       LevelSkill(31).LevelValue = 77
530       LevelSkill(32).LevelValue = 80
540       LevelSkill(33).LevelValue = 83
550       LevelSkill(34).LevelValue = 85
560       LevelSkill(35).LevelValue = 87
570       LevelSkill(36).LevelValue = 90
580       LevelSkill(37).LevelValue = 93
590       LevelSkill(38).LevelValue = 95
600       LevelSkill(39).LevelValue = 97
610       LevelSkill(40).LevelValue = 100
620       LevelSkill(41).LevelValue = 100
630       LevelSkill(42).LevelValue = 100
640       LevelSkill(43).LevelValue = 100
650       LevelSkill(44).LevelValue = 100
660       LevelSkill(45).LevelValue = 100
670       LevelSkill(46).LevelValue = 100
680       LevelSkill(47).LevelValue = 100
690       LevelSkill(48).LevelValue = 100
700       LevelSkill(49).LevelValue = 100
710       LevelSkill(50).LevelValue = 100
          
          
720       ListaRazas(eRaza.Humano) = "Humano"
730       ListaRazas(eRaza.Elfo) = "Elfo"
740       ListaRazas(eRaza.Drow) = "Drow"
750       ListaRazas(eRaza.Gnomo) = "Gnomo"
760       ListaRazas(eRaza.Enano) = "Enano"
          
770       ListaClases(eClass.Mage) = "Mago"
780       ListaClases(eClass.Cleric) = "Clerigo"
790       ListaClases(eClass.Warrior) = "Guerrero"
800       ListaClases(eClass.Assasin) = "Asesino"
810       ListaClases(eClass.Thief) = "Ladron"
820       ListaClases(eClass.Bard) = "Bardo"
830       ListaClases(eClass.Druid) = "Druida"
          'ListaClases(eClass.Bandit) = "Bandido"
840       ListaClases(eClass.Paladin) = "Paladin"
850       ListaClases(eClass.Hunter) = "Cazador"
860       ListaClases(eClass.Worker) = "Trabajador"
870       ListaClases(eClass.Pirat) = "Pirata"
          
880       SkillsNames(eSkill.Magia) = "Magia"
890       SkillsNames(eSkill.Robar) = "Robar"
900       SkillsNames(eSkill.Tacticas) = "Tactica en combate"
910       SkillsNames(eSkill.Armas) = "Combate con armas"
920       SkillsNames(eSkill.Meditar) = "Meditar"
930       SkillsNames(eSkill.Apuñalar) = "Apuñalar"
940       SkillsNames(eSkill.Ocultarse) = "Ocultarse"
950       SkillsNames(eSkill.Supervivencia) = "Supervivencia"
960       SkillsNames(eSkill.talar) = "Talar"
970       SkillsNames(eSkill.Comerciar) = "Comercio"
980       SkillsNames(eSkill.Defensa) = "Defensa con escudos"
990       SkillsNames(eSkill.Pesca) = "Pesca"
1000      SkillsNames(eSkill.Mineria) = "Mineria"
1010      SkillsNames(eSkill.Carpinteria) = "Carpinteria"
1020      SkillsNames(eSkill.herreria) = "Herreria"
1030      SkillsNames(eSkill.Liderazgo) = "Liderazgo"
1040      SkillsNames(eSkill.Domar) = "Domar animales"
1050      SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
1060      SkillsNames(eSkill.Wrestling) = "Combate sin armas"
1070      SkillsNames(eSkill.Navegacion) = "Navegacion"
1080      SkillsNames(eSkill.Equitacion) = "Equitacion"
1090      SkillsNames(eSkill.Resistencia) = "Resistencia Magica"
          
1100      ListaAtributos(eAtributos.Fuerza) = "Fuerza"
1110      ListaAtributos(eAtributos.Agilidad) = "Agilidad"
1120      ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
1130      ListaAtributos(eAtributos.Carisma) = "Carisma"
1140      ListaAtributos(eAtributos.Constitucion) = "Constitucion"
          
          
1150      frmCargando.Show
          
          'Call PlayWaveAPI(App.Path & "\wav\harp3.wav")
          
1160       Expc = val(GetVar(IniPath & "Server.ini", "INIT", "Expc"))
1170      Oroc = val(GetVar(IniPath & "Server.ini", "INIT", "Oroc"))
          
1180      frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
1190      IniPath = App.Path & "\"
1200      CharPath = App.Path & "\Charfile\"
          
          'Bordes del mapa
1210      MinXBorder = XMinMapSize + (XWindow \ 2)
1220      MaxXBorder = XMaxMapSize - (XWindow \ 2)
1230      MinYBorder = YMinMapSize + (YWindow \ 2)
1240      MaxYBorder = YMaxMapSize - (YWindow \ 2)
1250      DoEvents
          
1260      frmCargando.Label1(2).Caption = "Iniciando Arrays..."
          
1270      Call LoadGuildsDB
          
1280      Call LoadQuests
          
1290      Call CargarSpawnList
1300      Call CargarForbidenWords
          '¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
1310      frmCargando.Label1(2).Caption = "Cargando Server.ini"
          
1320      MaxUsers = 0
1330      Call LoadSini
1340      Call CargaApuestas
          
          '*************************************************
1350      frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
1360      Call CargaNpcsDat
          '*************************************************
          
1370      frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
          'Call LoadOBJData
1380      Call LoadOBJData
              
1390      frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
1400      Call CargarHechizos
              
              
1410      frmCargando.Label1(2).Caption = "Cargando Objetos de Herrería"
1420      Call LoadArmasHerreria
1430      Call LoadArmadurasHerreria
          
1440      frmCargando.Label1(2).Caption = "Cargando Objetos de Carpintería"
1450      Call LoadObjCarpintero
          
1460      frmCargando.Label1(2).Caption = "Cargando Balance.Dat"
1470      Call LoadBalance    '4/01/08 Pablo ToxicWaste
          
1480      frmCargando.Label1(2).Caption = "Cargando ArmadurasFaccionarias.dat"
1490      Call LoadArmadurasFaccion
          
1500      If BootDelBackUp Then
              
1510          frmCargando.Label1(2).Caption = "Cargando BackUp"
1520          Call CargarBackUp
1530      Else
1540          frmCargando.Label1(2).Caption = "Cargando Mapas"
1550          Call LoadMapData
1560      End If
          
          
          ' Cargar invocaciones
1570      LoadInvocaciones
          
1580      Call SonidosMapas.LoadSoundMapInfo

1590      Call generateMatrix(MATRIX_INITIAL_MAP)
          
          'Comentado porque hay worldsave en ese mapa!
          'Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
          '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
          
          Dim LoopC As Integer
          
          'Resetea las conexiones de los usuarios
1600      For LoopC = 1 To MaxUsers
1610          UserList(LoopC).ConnID = -1
1620          UserList(LoopC).ConnIDValida = False
1630          Set UserList(LoopC).incomingData = New clsByteQueue
1640          Set UserList(LoopC).outgoingData = New clsByteQueue
1650      Next LoopC
          
          '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
          
1660      With frmMain
1670          .AutoSave.Enabled = True
1680          .tPiqueteC.Enabled = True
1690          .GameTimer.Enabled = True
1700          .Auditoria.Enabled = True
1710          .KillLog.Enabled = True
1720          .TIMER_AI.Enabled = True
1730          .npcataca.Enabled = True
              '.AutoGP.Enabled = True
1740      End With
          
          '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
          'Configuracion de los sockets
          
1750      Call SecurityIp.InitIpTables(1000)
          
#If UsarQueSocket = 1 Then
          
1760      If LastSockListen >= 0 Then Call apiclosesocket(LastSockListen) 'Cierra el socket de escucha
1770      Call IniciaWsApi(frmMain.hWnd)
1780      SockListen = ListenForConnect(Puerto, hWndMsg, "")
              
1790      If SockListen <> -1 Then
1800          Call WriteVar(IniPath & "Server.ini", "INIT", "LastSockListen", SockListen) ' Guarda el socket escuchando
1810      Else
1820          MsgBox "Ha ocurrido un error al iniciar el socket del Servidor.", vbCritical + vbOKOnly
1830      End If
          
#ElseIf UsarQueSocket = 0 Then
          
1840      frmCargando.Label1(2).Caption = "Configurando Sockets"
          
1850      frmMain.Socket2(0).AddressFamily = AF_INET
1860      frmMain.Socket2(0).Protocol = IPPROTO_IP
1870      frmMain.Socket2(0).SocketType = SOCK_STREAM
1880      frmMain.Socket2(0).Binary = False
1890      frmMain.Socket2(0).Blocking = False
1900      frmMain.Socket2(0).BufferSize = 2048
          
1910      Call ConfigListeningSocket(frmMain.Socket1, Puerto)
          
#ElseIf UsarQueSocket = 2 Then
          
1920      frmMain.Serv.Iniciar Puerto
          
#ElseIf UsarQueSocket = 3 Then
          
1930      frmMain.TCPServ.Encolar True
1940      frmMain.TCPServ.IniciarTabla 1009
1950      frmMain.TCPServ.SetQueueLim 51200
1960      frmMain.TCPServ.Iniciar Puerto
          
#End If
          
1970      If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
          '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
          
1980      Unload frmCargando
          
          'Log
          Dim N As Integer
1990      N = FreeFile
2000      Open App.Path & "\logs\Main.log" For Append Shared As #N
2010      Print #N, Date & " " & time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
2020      Close #N
          
          'Ocultar
2030      If HideMe = 1 Then
2040          Call frmMain.InitMain(1)
2050      Else
2060          Call frmMain.InitMain(0)
2070      End If
          
2080      tInicioServer = GetTickCount() And &H7FFFFFFF
2090      Call InicializaEstadisticas

2100      RetosActivos = True
          EventosActivos = True


End Sub

Function FileExist(ByVal File As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
      '*****************************************************************
      'Se fija si existe el archivo
      '*****************************************************************

10        FileExist = LenB(dir$(File, FileType)) <> 0
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
      '*****************************************************************
      'Gets a field from a string
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 11/15/2004
      'Gets a field from a delimited string
      '*****************************************************************

          Dim i As Long
          Dim lastPos As Long
          Dim CurrentPos As Long
          Dim delimiter As String * 1
          
10        delimiter = Chr$(SepASCII)
          
20        For i = 1 To Pos
30            lastPos = CurrentPos
40            CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
50        Next i
          
60        If CurrentPos = 0 Then
70            ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
80        Else
90            ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
100       End If
End Function

Function MapaValido(ByVal map As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        MapaValido = map >= 1 And map <= NumMaps
End Function

Sub MostrarNumUsers()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        frmMain.CantUsuarios.Caption = "Número de usuarios jugando: " & NumUsers

End Sub
Public Sub LogCheats(desc As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.Path & "\logs\ANTICHEAT.log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & desc
50        Close #nfile
          
60        Exit Sub

Errhandler:

End Sub


Public Sub LogCriticEvent(desc As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & desc
50        Close #nfile
          
60        Exit Sub

Errhandler:

End Sub

Public Sub LogEjercitoReal(desc As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
40        Print #nfile, desc
50        Close #nfile
          
60        Exit Sub

Errhandler:

End Sub

Public Sub LogEjercitoCaos(desc As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
40        Print #nfile, desc
50        Close #nfile

60    Exit Sub

Errhandler:

End Sub


Public Sub LogIndex(ByVal Index As Integer, ByVal desc As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.Path & "\logs\" & Index & ".log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & desc
50        Close #nfile
          
60        Exit Sub

Errhandler:

End Sub
Public Sub LogCanje(desc As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.Path & "\logs\Canjes.log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & desc
50        Close #nfile
          
60        Exit Sub

Errhandler:

End Sub
Public Sub LogChangeNick(desc As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.Path & "\logs\ChangeNick.log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & desc
50        Close #nfile
          
60        Exit Sub

Errhandler:

End Sub
Public Sub ReportError(ByVal Archive As String, desc As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.Path & "\logs\ERROR_" & Archive & ".log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & desc
50        Close #nfile
          
60        Exit Sub

Errhandler:
End Sub
Public Sub LogError(desc As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.Path & "\logs\errores.log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & desc
50        Close #nfile
          
60        Exit Sub

Errhandler:

End Sub

Public Sub LogRetos(desc As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.Path & "\logs\Retos.log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & desc
50        Close #nfile
          
60        Exit Sub

Errhandler:

End Sub
Public Sub LogEventos(desc As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.Path & "\logs\EventosDS.log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & desc
50        Close #nfile
          
60        Exit Sub

Errhandler:

End Sub

Public Sub LogStatic(desc As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & desc
50        Close #nfile

60    Exit Sub

Errhandler:

End Sub

Public Sub LogTarea(desc As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile(1) ' obtenemos un canal
30        Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & desc
50        Close #nfile

60    Exit Sub

Errhandler:


End Sub


Public Sub LogClanes(ByVal Str As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim nfile As Integer
10        nfile = FreeFile ' obtenemos un canal
20        Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
30        Print #nfile, Date & " " & time & " " & Str
40        Close #nfile

End Sub

Public Sub LogIP(ByVal Str As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim nfile As Integer
10        nfile = FreeFile ' obtenemos un canal
20        Open App.Path & "\logs\IP.log" For Append Shared As #nfile
30        Print #nfile, Date & " " & time & " " & Str
40        Close #nfile

End Sub


Public Sub LogDesarrollo(ByVal Str As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim nfile As Integer
10        nfile = FreeFile ' obtenemos un canal
20        Open App.Path & "\logs\desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
30        Print #nfile, Date & " " & time & " " & Str
40        Close #nfile

End Sub

Public Sub LogGM(Nombre As String, texto As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
          'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
30        Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & texto
50        Close #nfile
          
60        Exit Sub

Errhandler:

End Sub
Public Sub LogUserConect(Nombre As String, texto As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
          'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
30        Open App.Path & "\logsuser\" & Nombre & ".log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & texto
50        Close #nfile
          
60        Exit Sub

Errhandler:

End Sub
Public Sub LogT0Error(texto As String)

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
          'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
30        Open App.Path & "\logs\Seguridad\T0\ERROR.log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & texto
50        Close #nfile
          
60        Exit Sub

Errhandler:

End Sub
Public Sub LogT0Ban(Nombre As String, texto As String)

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
          'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
30        Open App.Path & "\logs\Seguridad\T0\Baneos\" & Nombre & ".log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & texto
50        Close #nfile
          
60        Exit Sub

Errhandler:

End Sub
Public Sub LogT0Unban(Nombre As String, texto As String)

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
          'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
30        Open App.Path & "\logs\Seguridad\T0\Unbaneos\" & Nombre & ".log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & texto
50        Close #nfile
          
60        Exit Sub

Errhandler:

End Sub
Public Sub LogFotoDenuncia(Nombre As String, texto As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************ç

10    On Error GoTo Errhandler
          Dim nfile As Integer
          
20        nfile = FreeFile ' obtenemos un canal
          
30        Open App.Path & "\logs\Foto-Denuncias.log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & texto
50        Close #nfile
          
60        Exit Sub

Errhandler:

End Sub

Public Sub LogAsesinato(texto As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler
          Dim nfile As Integer
          
20        nfile = FreeFile ' obtenemos un canal
          
30        Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & texto
50        Close #nfile
          
60        Exit Sub

Errhandler:

End Sub
Public Sub logVentaCasa(ByVal texto As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
          
30        Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
40        Print #nfile, "----------------------------------------------------------"
50        Print #nfile, Date & " " & time & " " & texto
60        Print #nfile, "----------------------------------------------------------"
70        Close #nfile
          
80        Exit Sub

Errhandler:

End Sub
Public Sub LogHackAttemp(texto As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
40        Print #nfile, "----------------------------------------------------------"
50        Print #nfile, Date & " " & time & " " & texto
60        Print #nfile, "----------------------------------------------------------"
70        Close #nfile
          
80        Exit Sub

Errhandler:

End Sub

Public Sub LogCheating(texto As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.Path & "\logs\CH.log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & texto
50        Close #nfile
          
60        Exit Sub

Errhandler:

End Sub


Public Sub LogCriticalHackAttemp(texto As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
40        Print #nfile, "----------------------------------------------------------"
50        Print #nfile, Date & " " & time & " " & texto
60        Print #nfile, "----------------------------------------------------------"
70        Close #nfile
          
80        Exit Sub

Errhandler:

End Sub

Public Sub LogAntiCheat(texto As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler
          
          SendData SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(texto, FontTypeNames.FONTTYPE_BRONCE)
          
          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & texto
50        Print #nfile, ""
60        Close #nfile
          
70        Exit Sub

Errhandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim Arg As String
          Dim i As Integer
          
          
10        For i = 1 To 33
          
20        Arg = ReadField(i, cad, 44)
          
30        If LenB(Arg) = 0 Then Exit Function
          
40        Next i
          
50        ValidInputNP = True

End Function


Sub Restart()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      'Se asegura de que los sockets estan cerrados e ignora cualquier err
10    On Error Resume Next

20        If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."
          
          Dim LoopC As Long
        
#If UsarQueSocket = 0 Then

30        frmMain.Socket1.Cleanup
40        frmMain.Socket1.Startup
            
50        frmMain.Socket2(0).Cleanup
60        frmMain.Socket2(0).Startup

#ElseIf UsarQueSocket = 1 Then

          'Cierra el socket de escucha
70        If SockListen >= 0 Then Call apiclosesocket(SockListen)
          
          'Inicia el socket de escucha
80        SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 2 Then

#End If

90        For LoopC = 1 To MaxUsers
100           Call CloseSocket(LoopC)
110       Next
          
          'Initialize statistics!!
120       Call Statistics.Initialize
          
130       For LoopC = 1 To UBound(UserList())
140           Set UserList(LoopC).incomingData = Nothing
150           Set UserList(LoopC).outgoingData = Nothing
160       Next LoopC
          
170       ReDim UserList(1 To MaxUsers) As User
          
180       For LoopC = 1 To MaxUsers
190           UserList(LoopC).ConnID = -1
200           UserList(LoopC).ConnIDValida = False
210           Set UserList(LoopC).incomingData = New clsByteQueue
220           Set UserList(LoopC).outgoingData = New clsByteQueue
230       Next LoopC
          
240       LastUser = 0
250       NumUsers = 0
          
260       Call FreeNPCs
270       Call FreeCharIndexes
          
280       Call LoadSini
          
290       Call ResetForums
300       Call LoadOBJData
          
310       Call LoadMapData
          
320       Call CargarHechizos

#If UsarQueSocket = 0 Then

          '*****************Setup socket
330       frmMain.Socket1.AddressFamily = AF_INET
340       frmMain.Socket1.Protocol = IPPROTO_IP
350       frmMain.Socket1.SocketType = SOCK_STREAM
360       frmMain.Socket1.Binary = False
370       frmMain.Socket1.Blocking = False
380       frmMain.Socket1.BufferSize = 1024
          
390       frmMain.Socket2(0).AddressFamily = AF_INET
400       frmMain.Socket2(0).Protocol = IPPROTO_IP
410       frmMain.Socket2(0).SocketType = SOCK_STREAM
420       frmMain.Socket2(0).Blocking = False
430       frmMain.Socket2(0).BufferSize = 2048
          
          'Escucha
440       frmMain.Socket1.LocalPort = val(Puerto)
450       frmMain.Socket1.listen

#ElseIf UsarQueSocket = 1 Then

#ElseIf UsarQueSocket = 2 Then

#End If

460       If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
          
          'Log it
          Dim N As Integer
470       N = FreeFile
480       Open App.Path & "\logs\Main.log" For Append Shared As #N
490       Print #N, Date & " " & time & " servidor reiniciado."
500       Close #N
          
          'Ocultar
          
510       If HideMe = 1 Then
520           Call frmMain.InitMain(1)
530       Else
540           Call frmMain.InitMain(0)
550       End If

        
End Sub


Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
      '**************************************************************
      'Author: Unknown
      'Last Modify Date: 15/11/2009
      '15/11/2009: ZaMa - La lluvia no quita stamina en las arenas.
      '23/11/2009: ZaMa - Optimizacion de codigo.
      '**************************************************************

10        With UserList(UserIndex)
20            If MapInfo(.Pos.map).Zona <> "DUNGEON" Then
30                If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger <> 1 And _
                     MapData(.Pos.map, .Pos.X, .Pos.Y).trigger <> 2 And _
                     MapData(.Pos.map, .Pos.X, .Pos.Y).trigger <> 4 Then Intemperie = False
                     
40            Else
50                Intemperie = False
60            End If
70        End With
          
          
          'En las arenas no te afecta la lluvia
          
80        If IsArena(UserIndex) Then Intemperie = False
End Function

Public Sub EfectoLluvia(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

20        If UserList(UserIndex).flags.UserLogged Then
            '  If Intemperie(UserIndex) Then
30                Call FlushBuffer(UserIndex)
            '  End If
40        End If
          
50        Exit Sub
Errhandler:
60        LogError ("Error en EfectoLluvia")
End Sub


Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim i As Integer
10        For i = 1 To MAXMASCOTAS
20            With UserList(UserIndex)
30                If .MascotasIndex(i) > 0 Then
40                    If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
50                       Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia = _
                         Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia - 1
60                       If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(.MascotasIndex(i), 0)
70                    End If
80                End If
90            End With
100       Next i
End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)
      '***************************************************
      'Autor: Unkonwn
      'Last Modification: 23/11/2009
      'If user is naked and it's in a cold map, take health points from him
      '23/11/2009: ZaMa - Optimizacion de codigo.
      '***************************************************
          Dim modifi As Integer
         
10       With UserList(UserIndex)
20            If .Counters.Frio < IntervaloFrio Then
30                .Counters.Frio = .Counters.Frio + 1
40            Else
50                If MapInfo(.Pos.map).Terreno = Dungeon Then
60                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muriendo de frío, abrigate o morirás!!", FontTypeNames.FONTTYPE_INFO)
70                    modifi = Porcentaje(.Stats.MaxHp, 5)
80                    .Stats.MinHp = .Stats.MinHp - modifi
                      
90                    If .Stats.MinHp < 1 Then
100                       Call WriteConsoleMsg(UserIndex, "¡¡Has muerto de frío!!", FontTypeNames.FONTTYPE_INFO)
110                       .Stats.MinHp = 0
120                       Call UserDie(UserIndex)
130                   End If
                      
140                   Call WriteUpdateHP(UserIndex)
150               Else
160                   modifi = Porcentaje(.Stats.MaxSta, 5)
170                   Call QuitarSta(UserIndex, modifi)
180                   Call WriteUpdateSta(UserIndex)
190               End If
                  
200               .Counters.Frio = 0
210           End If
220       End With
End Sub

''
' Maneja  el efecto del estado atacable
'
' @param UserIndex  El index del usuario a ser afectado por el estado atacable
'

Public Sub EfectoEstadoAtacable(ByVal UserIndex As Integer)
      '******************************************************
      'Author: ZaMa
      'Last Update: 18/09/2010 (ZaMa)
      '18/09/2010: ZaMa - Ahora se activa el seguro cuando dejas de ser atacable.
      '******************************************************

          ' Si ya paso el tiempo de penalizacion
10        If Not IntervaloEstadoAtacable(UserIndex) Then
              ' Deja de poder ser atacado
20            UserList(UserIndex).flags.AtacablePor = 0
              
              ' Activo el seguro si deja de estar atacable
30            If Not UserList(UserIndex).flags.Seguro Then
40                Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn)
50            End If
              
              ' Send nick normal
60            Call RefreshCharStatus(UserIndex)
70        End If
          
End Sub

''
' Maneja el tiempo y el efecto del mimetismo
'
' @param UserIndex  El index del usuario a ser afectado por el mimetismo
'

Public Sub EfectoMimetismo(ByVal UserIndex As Integer)
      '******************************************************
      'Author: Unknown
      'Last Update: 12/01/2010 (ZaMa)
      '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
      '******************************************************
          Dim Barco As ObjData
          
10        With UserList(UserIndex)
20            If .Invent.AnilloNpcSlot > 0 Then Exit Sub
              
30            If .Counters.Mimetismo < IntervaloInvisible Then
40                .Counters.Mimetismo = .Counters.Mimetismo + 1
50            Else
                  'restore old char
60                Call WriteConsoleMsg(UserIndex, "Recuperas tu apariencia normal.", FontTypeNames.FONTTYPE_INFO)
                  
70                If .flags.Navegando Then
80                    If .flags.Muerto = 0 Then
90                        If .Faccion.ArmadaReal = 1 Then
100                           .Char.body = iFragataReal
110                       ElseIf .Faccion.FuerzasCaos = 1 Then
120                           .Char.body = iFragataCaos
130                       Else
140                           Barco = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
150                           If criminal(UserIndex) Then
160                               If Barco.Ropaje = iBarca Then .Char.body = iBarcaPk
170                               If Barco.Ropaje = iGalera Then .Char.body = iGaleraPk
180                               If Barco.Ropaje = iGaleon Then .Char.body = iGaleonPk
190                           Else
200                               If Barco.Ropaje = iBarca Then .Char.body = iBarcaCiuda
210                               If Barco.Ropaje = iGalera Then .Char.body = iGaleraCiuda
220                               If Barco.Ropaje = iGaleon Then .Char.body = iGaleonCiuda
230                           End If
240                       End If
250                   Else
260                       .Char.body = iFragataFantasmal
270                   End If
                      
280                   .Char.ShieldAnim = NingunEscudo
290                   .Char.WeaponAnim = NingunArma
300                   .Char.CascoAnim = NingunCasco
310                    If UserList(UserIndex).flags.Montando = 1 Then
320                           .Char.body = ObjData(UserList(UserIndex).Invent.MonturaObjIndex).Ropaje
330                           .Char.Head = .OrigChar.Head
340                           .Char.WeaponAnim = NingunArma
350                           .Char.ShieldAnim = NingunEscudo
360                           .Char.CascoAnim = .Char.CascoAnim
370                   End If
380               Else
390                   .Char.body = .CharMimetizado.body
400                   .Char.Head = .CharMimetizado.Head
410                   .Char.CascoAnim = .CharMimetizado.CascoAnim
420                   .Char.ShieldAnim = .CharMimetizado.ShieldAnim
430                   .Char.WeaponAnim = .CharMimetizado.WeaponAnim
440               End If
                        '     .MimetizmoName = vbNullString
450               With .Char
460                   Call ChangeUserChar(UserIndex, .body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
470               End With
                 
               '   ResetearName userIndex
                 
                  
480               .Counters.Mimetismo = 0
490               .flags.Mimetizado = 0
                  ' Se fue el efecto del mimetismo, puede ser atacado por npcs
500               .flags.Ignorado = False
510           End If
520       End With
End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 16/09/2010 (ZaMa)
      '16/09/2010: ZaMa - Al perder el invi cuando navegas, no se manda el mensaje de sacar invi (ya estas visible).
      '***************************************************

10        With UserList(UserIndex)
20            If .Counters.Invisibilidad < IntervaloInvisible Then
30                .Counters.Invisibilidad = .Counters.Invisibilidad + 1
40            Else
50                .Counters.Invisibilidad = RandomNumber(-100, 100) ' Invi variable :D
60                .flags.invisible = 0
70                If .flags.Oculto = 0 Then
80                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                      
                      ' Si navega ya esta visible..
90                    If Not .flags.Navegando = 1 Then
100                       Call SetInvisible(UserIndex, .Char.CharIndex, False)
110                   End If
                      
120               End If
130           End If
140       End With

End Sub


Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With Npclist(NpcIndex)
20            If .Contadores.Paralisis > 0 Then
30                .Contadores.Paralisis = .Contadores.Paralisis - 1
40            Else
50                .flags.Paralizado = 0
60                .flags.Inmovilizado = 0
70            End If
80        End With

End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With UserList(UserIndex)
20            If .Counters.Ceguera > 0 Then
30                .Counters.Ceguera = .Counters.Ceguera - 1
40            Else
50                If .flags.Ceguera = 1 Then
60                    .flags.Ceguera = 0
70                    Call WriteBlindNoMore(UserIndex)
80                End If
90                If .flags.Estupidez = 1 Then
100                   .flags.Estupidez = 0
110                   Call WriteDumbNoMore(UserIndex)
120               End If
              
130           End If
140       End With

End Sub


Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With UserList(UserIndex)
20            If .Counters.Paralisis > 0 Then
30                .Counters.Paralisis = .Counters.Paralisis - 1
40            Else
50                .flags.Paralizado = 0
60                .flags.Inmovilizado = 0
                  '.Flags.AdministrativeParalisis = 0
70                Call WriteParalizeOK(UserIndex)
80            End If
90        End With

End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With UserList(UserIndex)
20            If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 1 And _
                 MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 2 And _
                 MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 4 Then Exit Sub
              
              
              Dim massta As Integer
30            If .Stats.MinSta < .Stats.MaxSta Then
40                If .Counters.STACounter < Intervalo Then
50                    .Counters.STACounter = .Counters.STACounter + 1
60                Else
70                    EnviarStats = True
80                    .Counters.STACounter = 0
90                    If .flags.Desnudo Then Exit Sub 'Desnudo no sube energía. (ToxicWaste)
                      If .Invent.WeaponEqpObjIndex = Declaraciones.CAÑA_COFRES Then Exit Sub
                      
                      
100                   massta = RandomNumber(1, Porcentaje(.Stats.MaxSta, 5))

110                   .Stats.MinSta = .Stats.MinSta + massta
120                   If .Stats.MinSta > .Stats.MaxSta Then
130                       .Stats.MinSta = .Stats.MaxSta
140                   End If
150               End If
160           End If
170       End With
          
End Sub

Public Sub EfectoVeneno(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim N As Integer
          
10        With UserList(UserIndex)
20            If .Counters.Veneno < IntervaloVeneno Then
30              .Counters.Veneno = .Counters.Veneno + 1
40            Else
50              Call WriteConsoleMsg(UserIndex, "Estás envenenado, si no te curas morirás.", FontTypeNames.FONTTYPE_VENENO)
60              .Counters.Veneno = 0
70              N = RandomNumber(1, 5)
80              .Stats.MinHp = .Stats.MinHp - N
90              If .Stats.MinHp < 1 Then Call UserDie(UserIndex)
100             Call WriteUpdateHP(UserIndex)
110             Call WriteUpdateFollow(UserIndex)
120           End If
130       End With

End Sub

Public Sub DuracionPociones(ByVal UserIndex As Integer)
      '***************************************************
      'Author: ??????
      'Last Modification: 11/27/09 (Budi)
      'Cuando se pierde el efecto de la poción updatea fz y agi (No me gusta que ambos atributos aunque se haya modificado solo uno, pero bueno :p)
      '***************************************************
10        With UserList(UserIndex)
              'Controla la duracion de las pociones
20            If .flags.DuracionEfecto > 0 Then
30               .flags.DuracionEfecto = .flags.DuracionEfecto - 1
40               If .flags.DuracionEfecto = 0 Then
50                    .flags.TomoPocion = False
60                    .flags.TipoPocion = 0
                      'volvemos los atributos al estado normal
                      Dim LoopX As Integer
                      
70                    For LoopX = 1 To NUMATRIBUTOS
80                        .Stats.UserAtributos(LoopX) = .Stats.UserAtributosBackUP(LoopX)
90                    Next LoopX
                      
100                   Call WriteUpdateStrenghtAndDexterity(UserIndex)
110              End If
120           End If
130       End With

End Sub

Public Sub HambreYSed(ByVal UserIndex As Integer, ByRef fenviarAyS As Boolean)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With UserList(UserIndex)
20            If Not .flags.Privilegios And PlayerType.User Then Exit Sub
              
              'Sed
30            If .Stats.MinAGU > 0 Then
40                If .Counters.AGUACounter < IntervaloSed Then
50                    .Counters.AGUACounter = .Counters.AGUACounter + 1
60                Else
70                    .Counters.AGUACounter = 0
80                    .Stats.MinAGU = .Stats.MinAGU - 10
                      
90                    If .Stats.MinAGU <= 0 Then
100                       .Stats.MinAGU = 0
110                       .flags.Sed = 1
120                   End If
                      
130                   fenviarAyS = True
140               End If
150           End If
              
              'hambre
160           If .Stats.MinHam > 0 Then
170              If .Counters.COMCounter < IntervaloHambre Then
180                   .Counters.COMCounter = .Counters.COMCounter + 1
190              Else
200                   .Counters.COMCounter = 0
210                   .Stats.MinHam = .Stats.MinHam - 10
220                   If .Stats.MinHam <= 0 Then
230                          .Stats.MinHam = 0
240                          .flags.Hambre = 1
250                   End If
260                   fenviarAyS = True
270               End If
280           End If
290       End With

End Sub

Public Sub Sanar(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With UserList(UserIndex)
20            If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 1 And _
                 MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 2 And _
                 MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 4 Then Exit Sub
              
              Dim mashit As Integer
              'con el paso del tiempo va sanando....pero muy lentamente ;-)
30            If .Stats.MinHp < .Stats.MaxHp Then
40                If .Counters.HPCounter < Intervalo Then
50                    .Counters.HPCounter = .Counters.HPCounter + 1
60                Else
70                    mashit = RandomNumber(2, Porcentaje(.Stats.MaxSta, 5))
                      
80                    .Counters.HPCounter = 0
90                    .Stats.MinHp = .Stats.MinHp + mashit
100                   If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
110                   Call WriteConsoleMsg(UserIndex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
120                   EnviarStats = True
130               End If
140           End If
150       End With

End Sub

Public Sub CargaNpcsDat()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim npcfile As String
          
10        npcfile = DatPath & "NPCs.dat"
20        Call LeerNPCs.Initialize(npcfile)
End Sub
 
Sub PasarSegundo()

10    On Error GoTo Errhandler
          Dim i As Long
          
20        EventosDS.LoopEvent
          
30        If CountDownLimpieza > 0 Then
40            CountDownLimpieza = CountDownLimpieza - 1
              
50            If CountDownLimpieza <= 0 Then
60                CountDownLimpieza = 0
70                Call LimpiarMundo
80            End If
          
90        End If

100        If CuentaRegresivaTimer > 0 Then
110           CuentaRegresivaTimer = CuentaRegresivaTimer - 1
              
120           If CuentaRegresivaTimer <= 0 Then
130               Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("YA!", FontTypeNames.FONTTYPE_FIGHT))
140           Else
150               Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(CuentaRegresivaTimer, FontTypeNames.FONTTYPE_GUILD))
160           End If
170       End If
          
180       For i = 1 To LastUser
190           With UserList(i)
                  If .Counters.TimeTelep > 0 Then
                        .Counters.TimeTelep = .Counters.TimeTelep - 1
                        
                        If .Counters.TimeTelep <= 0 Then
                            WarpUserChar i, 1, 50, 50, True
                            WriteConsoleMsg i, "El efecto de la teletransportación ha terminado.", FontTypeNames.FONTTYPE_INFO
                            .Counters.TimeTelep = 0
                        End If
                  End If
                  
200               If .Counters.TimePin > 0 Then
210                   .Counters.TimePin = .Counters.TimePin - 1
220               End If
                  
230               .Counters.TimeAntiFriz = .Counters.TimeAntiFriz + 1
                  .Counters.TimePotas = .Counters.TimePotas + 1
                  
                  ' Cada un minuto reiniciamos el anti friz
240               If .Counters.TimeAntiFriz = 30 Then
250                   .PaquetesBasura = 0
260                   .Counters.TimeAntiFriz = 0
270               End If
                  
280               If .Counters.TimeFight > 0 Then
290                   .Counters.TimeFight = .Counters.TimeFight - 1
                          
300                   If .Counters.TimeFight = 0 Then
310                           WriteConsoleMsg i, "Cuenta» ¡YA!", FontTypeNames.FONTTYPE_FIGHT
                              
                              ' En los duelos desparalizamos el cliente
320                           If .flags.SlotEvent > 0 Then
330                               If Events(.flags.SlotEvent).Modality = Enfrentamientos Then
340                                   Call WriteUserInEvent(i)
350                               End If
360                           End If
                              
370                           If .flags.SlotReto > 0 Or .flags.InCVC Then
380                               Call WriteUserInEvent(i)
390                           End If
                              
                              
400                   Else
410                       WriteConsoleMsg i, "Cuenta» " & .Counters.TimeFight, FontTypeNames.FONTTYPE_GUILD
420                   End If
430               End If
                  
440               If .Counters.TimeCastleMode > 0 Then
450                   .Counters.TimeCastleMode = .Counters.TimeCastleMode - 1
                      
460                   WriteConsoleMsg i, "Revivirás en " & .Counters.TimeCastleMode & " ...", FontTypeNames.FONTTYPE_GUILD
                      
470                   If .Counters.TimeCastleMode <= 0 Then
480                       EventosDS.CastleMode_UserRevive i
490                       WriteConsoleMsg i, "Has revivido. ¡Ve a defender a tu Rey!", FontTypeNames.FONTTYPE_GUILD
500                   End If
510               End If
              
520           End With
              
                  'Cerrar usuario
530               If UserList(i).Counters.Saliendo Then
540                   UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
550                   If UserList(i).Counters.Salir <= 0 Then
560                      Call WriteConsoleMsg(i, "Gracias por jugar Desterium AO", FontTypeNames.FONTTYPE_INFO)
570                       Call WriteDisconnect(i)
580                       Call FlushBuffer(i)
                          
590                       Call CloseSocket(i)
600                   End If
610               End If
              
              
620           If UserList(i).Counters.Denuncia > 0 Then
630               UserList(i).Counters.Denuncia = UserList(i).Counters.Denuncia - 1
                    
640               If UserList(i).Counters.Denuncia <= 0 Then
650                   UserList(i).Counters.Denuncia = 0
660               End If
670          End If
              
              
680       Next i
690   Exit Sub

Errhandler:
700       Call LogError("Error en PasarSegundo. Err: " & Err.Description & " - " & Err.Number & " - UserIndex: " & i)
710       Resume Next
End Sub
 
Public Function ReiniciarAutoUpdate() As Double
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          'WorldSave
10        Call ES.DoBackUp

          'commit experiencias
20        Call mGroup.DistributeExpAndGldGroups

          'Guardar Pjs
30        Call GuardarUsuarios
          
40        If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

          'Chauuu
50        Unload frmMain

End Sub
Sub BackupUsers()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************
          'haciendoBK = True
          
         ' Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
          'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> testingG", FontTypeNames.FONTTYPE_SERVER))
          
          Dim i As Integer

          
10        For i = 1 To LastUser
20            If UserList(i).flags.UserLogged Then
30                Call SaveUser(i, CharPath & UCase$(UserList(i).name) & ".chr", False)
40            End If
50        Next i
          
          'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER))
          'Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
          'haciendoBK = False
End Sub
Sub GuardarUsuarios()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************
10        haciendoBK = True
          
20        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
30        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Grabando Personajes", FontTypeNames.FONTTYPE_SERVER))
          
          Dim i As Integer

          
40        For i = 1 To LastUser
50            If UserList(i).flags.UserLogged Then
60                Call SaveUser(i, CharPath & UCase$(UserList(i).name) & ".chr", False)
70            End If
80        Next i
          
90        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER))
100       Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
110       haciendoBK = False
End Sub


Sub InicializaEstadisticas()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim Ta As Long
10        Ta = GetTickCount() And &H7FFFFFFF
          
20        Call EstadisticasWeb.Inicializa(frmMain.hWnd)
30        Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
40        Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
50        Call EstadisticasWeb.Informar(UPTIME_SERVER, (Ta - tInicioServer) / 1000)
60        Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)

End Sub

Public Sub FreeNPCs()
      '***************************************************
      'Autor: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Releases all NPC Indexes
      '***************************************************
          Dim LoopC As Long
          
          ' Free all NPC indexes
10        For LoopC = 1 To MAXNPCS
20            Npclist(LoopC).flags.NPCActive = False
30        Next LoopC
End Sub

Public Sub FreeCharIndexes()
      '***************************************************
      'Autor: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/17/06
      'Releases all char indexes
      '***************************************************
          ' Free all char indexes (set them all to 0)
10        Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))
End Sub
Public Sub elefectofrio(ByVal UserIndex As Integer)
         Dim modifi As Integer
10        With UserList(UserIndex)
20            If .Counters.Frio < IntervaloFrio Then
30                .Counters.Frio = .Counters.Frio + 1
40            Else
50            If MapInfo(.Pos.map).Terreno = Nieve Then
60                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muriendo de frío, abrigate o morirás!!", FontTypeNames.FONTTYPE_INFO)
70                    modifi = Porcentaje(.Stats.MaxHp, 5)
80                    .Stats.MinHp = .Stats.MinHp - modifi
                      
90                    If .Stats.MinHp < 1 Then
100                       Call WriteConsoleMsg(UserIndex, "¡¡Has muerto de frío!!", FontTypeNames.FONTTYPE_INFO)
110                       .Stats.MinHp = 0
120                       Call UserDie(UserIndex)
130                   End If
                      
140                   Call WriteUpdateHP(UserIndex)
150                   Call WriteUpdateFollow(UserIndex)
160               Else
170                   modifi = Porcentaje(.Stats.MaxSta, 5)
180                   Call QuitarSta(UserIndex, modifi)
190                   Call WriteUpdateSta(UserIndex)
                  
200               End If
210               .Counters.Frio = 0
220           End If
230       End With
End Sub
Public Sub LogMD5(LOG As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************
       
10    On Error GoTo Errhandler
       
          Dim mA As Integer
20        mA = FreeFile ' obtenemos un canal
30        Open App.Path & "\logs\MD5-NO OK.log" For Append Shared As #mA
40        Print #mA, Date & " " & time & " " & LOG
50        Close #mA
         
60        Exit Sub
       
Errhandler:
       
End Sub
Public Sub Viajes(ByVal UserIndex As Integer, ByVal Lug As Byte)
10        With UserList(UserIndex)
         
20            Select Case Lug
                  Case 0 'Dungeon 1
30                    If .Stats.Gld < 10000 Then
40                        Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero, necesitas 10.000 monedas de oro para ir hacia el Muelle de Lindos.", FontTypeNames.FONTTYPE_INFO)
50                        Exit Sub
60                    End If
                     
70                    .Stats.Gld = .Stats.Gld - 10000: WriteUpdateGold UserIndex
                     
80                    Call WarpUserChar(UserIndex, 62, 83, 45, False)
90                    Call WriteConsoleMsg(UserIndex, "Has viajado hacia el Muelle de Lindos, se te han descontado 10.000 monedas de oro.", FontTypeNames.FONTTYPE_INFO)
                      
100               Case 1 'Dungeon 2
110                   If .Stats.Gld < 7000 Then
120                       Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero, necesitas 7.000 monedas de oro para viajar hacia el Muelle de Nix.", FontTypeNames.FONTTYPE_INFO)
130                       Exit Sub
140                   End If
                     
150                   .Stats.Gld = .Stats.Gld - 7000: WriteUpdateGold UserIndex
                     
160                   Call WarpUserChar(UserIndex, 34, 27, 79, False)
170      Call WriteConsoleMsg(UserIndex, "Has viajado hacia el Muelle de Nix, se te han descontado 7.000 monedas de oro.", FontTypeNames.FONTTYPE_INFO)
         
180          Case 2 'Dungeon 2
190                   If .Stats.Gld < 15000 Then
200                       Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero, necesitas 15.000 monedas de oro para viajar hacia el Muelle de Arghâl.", FontTypeNames.FONTTYPE_INFO)
210                       Exit Sub
220                   End If
                     
230                   .Stats.Gld = .Stats.Gld - 15000: WriteUpdateGold UserIndex
                     
240                   Call WarpUserChar(UserIndex, 150, 36, 30, False)
                      
250      Call WriteConsoleMsg(UserIndex, "Has viajado hacia el Muelle de Arghâl, se te han descontado 15.000 monedas de oro.", FontTypeNames.FONTTYPE_INFO)
         
260        Case 3 'Dungeon 2
270                   If .Stats.Gld < 15000 Then
280                       Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero, necesitas 15.000 monedas de oro para viajar hacia el Muelle de Banderbill.", FontTypeNames.FONTTYPE_INFO)
290                       Exit Sub
300                   End If
                     
310                   .Stats.Gld = .Stats.Gld - 15000: WriteUpdateGold UserIndex
                     
320                   Call WarpUserChar(UserIndex, 61, 66, 67, False)
330      Call WriteConsoleMsg(UserIndex, "Has viajado hacia el Muelle de Banderbill, se te han descontado 15.000 monedas de oro.", FontTypeNames.FONTTYPE_INFO)
         
340          Case 4 'Dungeon 2
350                   If .Stats.Gld < 350000 Then
360                       Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero, necesitas 350.000 monedas de oro para viajar hacia el Fuerte Pretoriano.", FontTypeNames.FONTTYPE_INFO)
370                       Exit Sub
380                   End If
                     
390                   .Stats.Gld = .Stats.Gld - 350000: WriteUpdateGold UserIndex
                     
400                   Call WarpUserChar(UserIndex, 196, 66, 94, False)
410      Call WriteConsoleMsg(UserIndex, "Has viajado hacia el Fuerte Pretoriano, se te han descontado 350.000 monedas de oro.", FontTypeNames.FONTTYPE_INFO)
         
420           End Select
430       End With
End Sub


Public Sub CheckLogros(ByVal UserIndex As Integer)
10        With UserList(UserIndex)
              
              ' Nuevo usuario
20            If .Logros(0) = 0 Then
30                .Logros(0) = 1
                  
40                WriteConsoleMsg UserIndex, "Has recibido el logro N°1 como agradecimiento por jugar DesteriumAO. Has recibido 10.000 monedas de oro como recompensa.", FontTypeNames.FONTTYPE_GUILD
                  
50                .Stats.Gld = .Stats.Gld + 10000
60                WriteUpdateGold UserIndex
70            End If
              
              ' Nivel 47
80            If .Logros(1) = 0 And .Stats.ELV = STAT_MAXELV Then
90                .Logros(1) = 1
                  
100               WriteConsoleMsg UserIndex, "Has recibido el logro N°2. ¡Felicitaciones has alcanzado el nivel máximo! Recibiste 1.000.000 monedas de oro.", FontTypeNames.FONTTYPE_GUILD
110               .Stats.Gld = .Stats.Gld + 1000000
120               WriteUpdateGold UserIndex
130           End If
              
              ' Fundador de clan
140           If .Logros(2) = 0 And HasFound(.name) Then
150               .Logros(2) = 1
                  
160               WriteConsoleMsg UserIndex, "Has recibido el logro N°3 debido a que has fundado un clan. ", FontTypeNames.FONTTYPE_GUILD

170           End If
          
              ' Usuario Oro
180           If .Logros(3) = 0 And .flags.Oro > 0 Then
190               .Logros(3) = 1
                  
200               WriteConsoleMsg UserIndex, "Has recibido el logro N°4 debido a que eres usuario ORO. ", FontTypeNames.FONTTYPE_GUILD
210           End If
              
              ' Usuario premium
220           If .Logros(4) = 0 And .flags.Premium > 0 Then
230               .Logros(4) = 1
                  
240               WriteConsoleMsg UserIndex, "Has recibido el logro N°5 debido a que eres usuario PREMIUM. ", FontTypeNames.FONTTYPE_GUILD
250           End If
              
              ' Usuarios matados 100
260           If .Logros(5) = 0 And .Stats.UsuariosMatados >= 100 Then
270               .Logros(5) = 1
                  
280               WriteConsoleMsg UserIndex, "Has recibido el logro N°6 debido a que has alcanzado los 100 usuarios matados. ", FontTypeNames.FONTTYPE_GUILD
290               .Stats.Gld = .Stats.Gld + 250000
300               WriteUpdateGold UserIndex
310           End If
              
              ' Usuarios matados 200
320           If .Logros(6) = 0 And .Stats.UsuariosMatados >= 200 Then
330               .Logros(6) = 1
                  
340               WriteConsoleMsg UserIndex, "Has recibido el logro N°7 debido a que has alcanzado los 200 usuarios matados. ", FontTypeNames.FONTTYPE_GUILD
350               .Stats.Gld = .Stats.Gld + 500000
360               WriteUpdateGold UserIndex
370           End If
              
              
              ' Usuarios matados 400
380           If .Logros(7) = 0 And .Stats.UsuariosMatados >= 400 Then
390               .Logros(7) = 1
                  
400               WriteConsoleMsg UserIndex, "Has recibido el logro N°8 debido a que has alcanzado los 400 usuarios matados. ", FontTypeNames.FONTTYPE_GUILD
410               .Stats.Gld = .Stats.Gld + 750000
420               WriteUpdateGold UserIndex
430           End If
              
              ' Usuarios matados 800
440           If .Logros(8) = 0 And .Stats.UsuariosMatados >= 800 Then
450               .Logros(8) = 1
                  
460               WriteConsoleMsg UserIndex, "Has recibido el logro N°9 debido a que has alcanzado los 800 usuarios matados. ", FontTypeNames.FONTTYPE_GUILD
470               .Stats.Gld = .Stats.Gld + 1000000
480               WriteUpdateGold UserIndex
490           End If
              
              ' Usuarios matados 1600
500           If .Logros(9) = 0 And .Stats.UsuariosMatados >= 1600 Then
510               .Logros(9) = 1
                  
520               WriteConsoleMsg UserIndex, "Has recibido el logro N°10 debido a que has alcanzado los 1600 usuarios matados. ", FontTypeNames.FONTTYPE_GUILD
530               .Stats.Gld = .Stats.Gld + 2000000
540               WriteUpdateGold UserIndex
550           End If
              
              ' Usuarios matados 5000
560           If .Logros(10) = 0 And .Stats.UsuariosMatados >= 5000 Then
570               .Logros(10) = 1
                  
580               WriteConsoleMsg UserIndex, "Has recibido el logro N°11 debido a que has alcanzado los 5000 usuarios matados. ", FontTypeNames.FONTTYPE_GUILD
590               .Stats.Gld = .Stats.Gld + 4000000
600               WriteUpdateGold UserIndex
610           End If
              
              ' Retos ganados
620           If .Logros(11) = 0 And .Stats.RetosGanados >= 5 Then
630               .Logros(11) = 1
                  
640               WriteConsoleMsg UserIndex, "Has recibido el logro N°12 debido a que has alcanzado los 5 retos ganados. ", FontTypeNames.FONTTYPE_GUILD
650               .Stats.Gld = .Stats.Gld + 150000
660               WriteUpdateGold UserIndex
670           End If
              
              ' Retos ganados
680           If .Logros(12) = 0 And .Stats.RetosGanados >= 10 Then
690               .Logros(12) = 1
                  
700               WriteConsoleMsg UserIndex, "Has recibido el logro N°13 debido a que has alcanzado los 10 retos ganados. ", FontTypeNames.FONTTYPE_GUILD
710               .Stats.Gld = .Stats.Gld + 300000
720               WriteUpdateGold UserIndex
730           End If
              
              ' Retos ganados
740           If .Logros(13) = 0 And .Stats.RetosGanados >= 50 Then
750               .Logros(13) = 1
                  
760               WriteConsoleMsg UserIndex, "Has recibido el logro N°14 debido a que has alcanzado los 50 retos ganados. ", FontTypeNames.FONTTYPE_GUILD
770               .Stats.Gld = .Stats.Gld + 500000
780               WriteUpdateGold UserIndex
790           End If
              
              ' Retos ganados
800           If .Logros(14) = 0 And .Stats.RetosGanados >= 100 Then
810               .Logros(14) = 1
                  
820               WriteConsoleMsg UserIndex, "Has recibido el logro N°15 debido a que has alcanzado los 100 retos ganados. ", FontTypeNames.FONTTYPE_GUILD
830               .Stats.Gld = .Stats.Gld + 600000
840               WriteUpdateGold UserIndex
850           End If
              
              ' Retos ganados
860           If .Logros(15) = 0 And .Stats.RetosGanados >= 250 Then
870               .Logros(15) = 1
                  
880               WriteConsoleMsg UserIndex, "Has recibido el logro N°16 debido a que has alcanzado los 250 retos ganados. ", FontTypeNames.FONTTYPE_GUILD
890               .Stats.Gld = .Stats.Gld + 1000000
900               WriteUpdateGold UserIndex
910           End If
              
              ' Retos ganados
920           If .Logros(16) = 0 And .Stats.RetosGanados >= 1000 Then
930               .Logros(16) = 1
                  
940               WriteConsoleMsg UserIndex, "Has recibido el logro N°17 debido a que has alcanzado los 1000 retos ganados. ", FontTypeNames.FONTTYPE_GUILD
950               .Stats.Gld = .Stats.Gld + 3500000
960               WriteUpdateGold UserIndex
970           End If
              
980           If .Logros(17) = 0 And .Stats.TorneosGanados >= 5 Then
990               .Logros(17) = 1
                  
1000              WriteConsoleMsg UserIndex, "Has recibido el logro N°18 debido a que has alcanzado los 5 retos ganados. ", FontTypeNames.FONTTYPE_GUILD
1010              .Stats.Gld = .Stats.Gld + 300000
1020              WriteUpdateGold UserIndex
1030          End If
              
1040          If .Logros(18) = 0 And .Stats.TorneosGanados >= 10 Then
1050              .Logros(18) = 1
                  
1060              WriteConsoleMsg UserIndex, "Has recibido el logro N°19 debido a que has alcanzado los 10 retos ganados. ", FontTypeNames.FONTTYPE_GUILD
1070              .Stats.Gld = .Stats.Gld + 500000
1080              WriteUpdateGold UserIndex
1090          End If
              
1100          If .Logros(19) = 0 And .Stats.TorneosGanados >= 20 Then
1110              .Logros(19) = 1
                  
1120              WriteConsoleMsg UserIndex, "Has recibido el logro N°20 debido a que has alcanzado los 20 retos ganados. ", FontTypeNames.FONTTYPE_GUILD
1130              .Stats.Gld = .Stats.Gld + 600000
1140              WriteUpdateGold UserIndex
1150          End If
              
1160          If .Logros(20) = 0 And .Stats.TorneosGanados >= 30 Then
1170              .Logros(20) = 1
                  
1180              WriteConsoleMsg UserIndex, "Has recibido el logro N°21 debido a que has alcanzado los 30 retos ganados. ", FontTypeNames.FONTTYPE_GUILD
1190              .Stats.Gld = .Stats.Gld + 700000
1200              WriteUpdateGold UserIndex
1210          End If
              
1220          If .Logros(21) = 0 And .Stats.TorneosGanados >= 40 Then
1230              .Logros(21) = 1
                  
1240              WriteConsoleMsg UserIndex, "Has recibido el logro N°22 debido a que has alcanzado los 40 retos ganados. ", FontTypeNames.FONTTYPE_GUILD
1250              .Stats.Gld = .Stats.Gld + 800000
1260              WriteUpdateGold UserIndex
1270          End If
              
1280          If .Logros(22) = 0 And .Stats.TorneosGanados >= 50 Then
1290              .Logros(22) = 1
                  
1300              WriteConsoleMsg UserIndex, "Has recibido el logro N°23 debido a que has alcanzado los 50 retos ganados. ", FontTypeNames.FONTTYPE_GUILD
1310              .Stats.Gld = .Stats.Gld + 900000
1320              WriteUpdateGold UserIndex
1330          End If
              
1340          If .Logros(23) = 0 And .Stats.TorneosGanados >= 100 Then
1350              .Logros(23) = 1
                  
1360              WriteConsoleMsg UserIndex, "Has recibido el logro N°24 debido a que has alcanzado los 100 retos ganados. ", FontTypeNames.FONTTYPE_GUILD
1370              .Stats.Gld = .Stats.Gld + 1000000
1380              WriteUpdateGold UserIndex
1390          End If
1400      End With
          
End Sub

Public Function UserPoderoso(ByVal UserIndex As Integer) As Integer
   On Error GoTo UserPoderoso_Error

10        With UserList(UserIndex)
20            If .Invent.AnilloNpcObjIndex > 0 Then
30                UserPoderoso = ObjData(.Invent.AnilloNpcObjIndex).NpcTipo
40            End If
50        End With

   On Error GoTo 0
   Exit Function

UserPoderoso_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure UserPoderoso of Módulo General in line " & Erl
End Function

'Transformusernpc Userindex, ObjData(.Invent.AnilloNpcObjIndex).NpcTipo, True
Public Sub TransformUserNpc(ByVal UserIndex As Integer, ByVal NpcTipo As Integer, ByVal Transformar As Boolean)
10        With UserList(UserIndex)
20            If Transformar Then
30                .CharMimetizado.body = .Char.body
40                .CharMimetizado.Head = .Char.Head
50                .CharMimetizado.CascoAnim = .Char.CascoAnim
60                .CharMimetizado.ShieldAnim = .Char.ShieldAnim
70                .CharMimetizado.WeaponAnim = .Char.WeaponAnim
                      
80                .flags.Mimetizado = 1
                      
90                .Char.body = NpcTipo
100               .Char.Head = 0
110               .Char.CascoAnim = NingunCasco
120               .Char.ShieldAnim = NingunEscudo
130               .Char.WeaponAnim = NingunArma
140           Else
150               .Char.body = .CharMimetizado.body
160               .Char.Head = .CharMimetizado.Head
170               .Char.CascoAnim = .CharMimetizado.CascoAnim
180               .Char.ShieldAnim = .CharMimetizado.ShieldAnim
190               .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                  
200               .CharMimetizado.body = 0
210               .CharMimetizado.Head = 0
220               .CharMimetizado.CascoAnim = NingunCasco
230               .CharMimetizado.ShieldAnim = NingunEscudo
240               .CharMimetizado.WeaponAnim = NingunArma
                      
250               .flags.Mimetizado = 0
260           End If
              
270           Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
          
280       End With
End Sub


Public Sub ChangeNick(ByVal UserIndex As Integer, ByVal UserName As String)

          Dim OldUserName As String
          Dim cantPenas As Byte
          
10        With UserList(UserIndex)
20            OldUserName = UCase$(.name)
              
30            If .flags.Muerto Then Exit Sub
              
40            If .GuildIndex > 0 Then
50                If modGuilds.GuildLeader(.GuildIndex) = UCase$(.name) Then
60                    WriteConsoleMsg UserIndex, "No puedes cambiar tu NICK si estás en un clan y eres el lider. Sal de él", FontTypeNames.FONTTYPE_INFO
70                    Exit Sub
80                End If
                  
90                WriteConsoleMsg UserIndex, "No puedes cambiar el nick estando en un clan.", FontTypeNames.FONTTYPE_INFO
100               Exit Sub
110           End If
              
120           If .flags.IsDios Then
130               WriteConsoleMsg UserIndex, "No puedes cambiar tu nick estánd en MODO DIOS.", FontTypeNames.FONTTYPE_INFO
140               Exit Sub
150           End If
              
160           If Not AsciiValidos(UserName) Or LenB(UserName) = 0 Then
170               Call WriteConsoleMsg(UserIndex, "Escoge un nombre válido", FontTypeNames.FONTTYPE_WARNING)
180               Exit Sub
190           End If
              
200           If PersonajeExiste(UCase$(UserName)) Then
210               WriteConsoleMsg UserIndex, "El nombre que has escogido ya existe. Prueba otro.", FontTypeNames.FONTTYPE_WARNING
220               Exit Sub
230           End If
          
240           If .Stats.Gld < 25000000 Then
250               WriteConsoleMsg UserIndex, "No tienes 25.000.000 monedas de oro en tu billetera.", FontTypeNames.FONTTYPE_INFO
260               Exit Sub
270           End If
              

280           .Stats.Gld = .Stats.Gld - 25000000
290           .flags.Ban = 1
300           WriteUpdateGold UserIndex
              
310           CloseSocket UserIndex
               
              
320           Call FileCopy(CharPath & OldUserName & ".chr", CharPath & UCase$(UserName) & ".chr")
330           Kill CharPath & OldUserName & ".chr"
              
                      
340           Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "0")

350           cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))

360           Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", CStr(cantPenas + 1))

370           Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & CStr(cantPenas + 1), LCase$(.name) & ": Cambio de NICK " & UCase$(UserName) & " (Ex " & OldUserName & ")" & Date & " " & time)
                             
              
380           If PersonajeExiste(OldUserName) Then
390               Call LogChangeNick("El personaje " & OldUserName & " no fue borrado. ¡CORROBORAR!")
400               Exit Sub
410           End If
              
420           Call LogChangeNick("El nick del personaje " & OldUserName & " ha pasado a ser " & UserName)
              
              Call UpdateAccountUserName(.Account, OldUserName, UserName)
              
430       End With
End Sub

Public Sub LoadCanjes()
          Dim FilePath As String
          Dim LoopC As Integer, LoopY As Integer
          Dim strTemp As String
          
10        FilePath = App.Path & "\DAT\Canjes.DAT"
20        NumCanjes = val(GetVar(FilePath, "INIT", "NumCanjes"))
          
30        ReDim Canjes(1 To NumCanjes) As tCanjes
          
40        For LoopC = 1 To NumCanjes
50            With Canjes(LoopC)
60                .NumRequired = val(GetVar(FilePath, "CANJE" & LoopC, "NumRequired"))
                  
                  
70                If .NumRequired <> 0 Then
80                ReDim .ObjRequired(1 To .NumRequired) As Obj
                  
90                For LoopY = 1 To .NumRequired
100                   strTemp = GetVar(FilePath, "CANJE" & LoopC, "ObjRequired" & LoopY)
                      
110                   .ObjRequired(LoopY).ObjIndex = val(ReadField(1, strTemp, Asc("-")))
120                   .ObjRequired(LoopY).Amount = val(ReadField(2, strTemp, Asc("-")))
                  
130               Next LoopY
                  
140               End If
150               strTemp = GetVar(FilePath, "CANJE" & LoopC, "ObjCanje")
160               .ObjCanje.ObjIndex = val(ReadField(1, strTemp, Asc("-")))
170               .ObjCanje.Amount = val(ReadField(2, strTemp, Asc("-")))
180               .Points = val(GetVar(FilePath, "CANJE" & LoopC, "Points"))
                  
                  
                  'NPCS
                  
190               .Npcs = val(GetVar(FilePath, "CANJE" & LoopC, "Npc"))
                  
200           End With
210       Next LoopC
          
          
End Sub

Public Sub CanjearObjeto(ByVal UserIndex As Integer, ByVal CanjeIndex As Byte)

          Dim LoopC As Integer
          Dim TempObj As Obj
          
10        With UserList(UserIndex)
          
20            If Canjes(CanjeIndex).Points > .Stats.Points Then
30                WriteConsoleMsg UserIndex, "Tus puntos no son lo suficientes para conseguir este canje.", FontTypeNames.FONTTYPE_WARNING
40                Exit Sub
50            End If
              
60            For LoopC = 1 To Canjes(CanjeIndex).NumRequired
70                If Not TieneObjetos(Canjes(CanjeIndex).ObjRequired(LoopC).ObjIndex, Canjes(CanjeIndex).ObjRequired(LoopC).Amount, UserIndex) Then
80                    WriteConsoleMsg UserIndex, "Para realizar este canje necesitas tener en tu inventario " & Canjes(CanjeIndex).ObjRequired(LoopC).Amount & " del item " & ObjData(Canjes(CanjeIndex).ObjRequired(LoopC).ObjIndex).name, FontTypeNames.FONTTYPE_WARNING
90                    Exit Sub
100               End If
110           Next LoopC

120           For LoopC = 1 To Canjes(CanjeIndex).NumRequired
130               Call QuitarObjetos(Canjes(CanjeIndex).ObjRequired(LoopC).ObjIndex, Canjes(CanjeIndex).ObjRequired(LoopC).Amount, UserIndex)
140           Next LoopC
              
              
150           TempObj = Canjes(CanjeIndex).ObjCanje
                  
160           If Not MeterItemEnInventario(UserIndex, TempObj) Then
170               TirarItemAlPiso .Pos, TempObj
180           End If
              
190           If Canjes(CanjeIndex).Points > 0 Then
                  .Stats.Points = .Stats.Points - Canjes(CanjeIndex).Points
                  WriteUpdatePoints UserIndex
200               '.Stats.TorneosGanados = .Stats.TorneosGanados - Canjes(CanjeIndex).Points
210           End If
              
              
220           LogCanje "El personaje " & .name & " ha canjeado el canje número " & CanjeIndex
230       End With
End Sub

Public Sub WarpPosAnt(ByVal UserIndex As Integer)
          ' • Warpeo del personaje a su posición anterior.
          
          Dim Pos As WorldPos
          
   On Error GoTo WarpPosAnt_Error

10        With UserList(UserIndex)
20            Pos.map = .PosAnt.map
30            Pos.X = .PosAnt.X
40            Pos.Y = .PosAnt.Y
                          
50            Call FindLegalPos(UserIndex, Pos.map, Pos.X, Pos.Y)
60            Call WarpUserChar(UserIndex, Pos.map, Pos.X, Pos.Y, False)
              
70            .PosAnt.map = 0
80            .PosAnt.X = 0
90            .PosAnt.Y = 0
          
100       End With

   On Error GoTo 0
   Exit Sub

WarpPosAnt_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WarpPosAnt of Módulo General in line " & Erl
End Sub

Public Sub Check_AutoRed(ByVal UserIndex As Integer)

          ' Chequeamos: Cuando un usuario llega al máximo de vida, se considera que no necesita más pociones azules.
10        With UserList(UserIndex)
          
20            If .Stats.MinHp = .Stats.MaxHp Then
30                .PotFull = True
40                .Counters.TimePotFull = 250
50                Exit Sub
60            End If
              
70        End With
          
End Sub


Public Function Tilde(data As String) As String
     
Tilde = Replace(Replace(Replace(Replace(Replace(UCase$(data), "Á", "A"), "É", "E"), "Í", "I"), "Ó", "O"), "Ú", "U")
     
End Function
