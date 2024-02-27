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

Sub DarCuerpoDesnudo(ByVal Userindex As Integer, Optional ByVal Mimetizado As Boolean = False)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/14/07
'Da cuerpo desnudo a un usuario
'23/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************

Dim CuerpoDesnudo As Integer

With UserList(Userindex)
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
        .CharMimetizado.body = CuerpoDesnudo
    Else
        .Char.body = CuerpoDesnudo
    End If
    
    .flags.Desnudo = 1
End With

End Sub


Sub Bloquear(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal b As Boolean)
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

    If toMap Then
        Call SendData(SendTarget.toMap, sndIndex, PrepareMessageBlockPosition(X, Y, b))
    Else
        Call WriteBlockPosition(sndIndex, X, Y, b)
    End If

End Sub


Function HayAgua(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If map > 0 And map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
        With MapData(map, X, Y)
            If ((.Graphic(1) >= 1505 And .Graphic(1) <= 1520) Or _
            (.Graphic(1) >= 5665 And .Graphic(1) <= 5680) Or _
            (.Graphic(1) >= 13547 And .Graphic(1) <= 13562)) And _
               .Graphic(2) = 0 Then
                    HayAgua = True
            Else
                    HayAgua = False
            End If
        End With
    Else
      HayAgua = False
    End If

End Function

Private Function HayLava(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/12/07
'***************************************************
    If map > 0 And map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
        If MapData(map, X, Y).Graphic(1) >= 5837 And MapData(map, X, Y).Graphic(1) <= 5852 Then
            HayLava = True
        Else
            HayLava = False
        End If
    Else
      HayLava = False
    End If

End Function
Function HayCura(ByVal Userindex As Integer) As Boolean
'******************************
'Adaptacion a 13.0: Kaneidra
'Last Modification: 15/05/2012
'******************************
 
Dim X As Integer, Y As Integer
 
For Y = UserList(Userindex).Pos.Y - MinYBorder + 1 To UserList(Userindex).Pos.Y + MinYBorder - 1
For X = UserList(Userindex).Pos.X - MinXBorder + 1 To UserList(Userindex).Pos.X + MinXBorder - 1
       
            If MapData(UserList(Userindex).Pos.map, X, Y).NpcIndex > 0 Then
                    If Npclist(MapData(UserList(Userindex).Pos.map, X, Y).NpcIndex).NPCtype = 1 Then
                        If Distancia(UserList(Userindex).Pos, Npclist(MapData(UserList(Userindex).Pos.map, X, Y).NpcIndex).Pos) < 10 Then
                        HayCura = True
                        Exit Function
                    End If
                End If
            End If
           
        Next X
Next Y
 
HayCura = False
 
End Function


Sub LimpiarMundo()
'***************************************************
'Author: Unknow
'Last Modification: 04/15/2008
'01/14/2008: Marcos Martinez (ByVal) - La funcion FOR estaba mal. En ves de i habia un 1.
'04/15/2008: (NicoNZ) - La funcion FOR estaba mal, de la forma que se hacia tiraba error.
'***************************************************
On Error GoTo errhandler

    Dim i As Integer
    Dim d As New cGarbage
    
    For i = TrashCollector.Count To 1 Step -1
        Set d = TrashCollector(i)
        Call EraseObj(1, d.map, d.X, d.Y)
        Call TrashCollector.Remove(i)
        Set d = Nothing
    Next i
    
    Call SecurityIp.IpSecurityMantenimientoLista
    
    Exit Sub

errhandler:
    Call LogError("Error producido en el sub LimpiarMundo: " & Err.description)
End Sub

Sub EnviarSpawnList(ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim k As Long
    Dim npcNames() As String
    
    ReDim npcNames(1 To UBound(SpawnList)) As String
    
    For k = 1 To UBound(SpawnList)
        npcNames(k) = SpawnList(k).NpcName
    Next k
    
    Call WriteSpawnList(Userindex, npcNames())

End Sub

Sub ConfigListeningSocket(ByRef Obj As Object, ByVal Port As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

#If UsarQueSocket = 0 Then

    Obj.AddressFamily = AF_INET
    Obj.Protocol = IPPROTO_IP
    Obj.SocketType = SOCK_STREAM
    Obj.Binary = False
    Obj.Blocking = False
    Obj.BufferSize = 1024
    Obj.LocalPort = Port
    Obj.backlog = 5
    Obj.listen

#End If

End Sub

Sub Main()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next
    Dim f As Date
    
    ChDir App.Path
    ChDrive App.Path
    
    Call BanIpCargar
    LoadMercadoArgentum
    Call LoadRanking
    Call eventLoad
    Call BanHDCargar '//Disco
    GlobalActivado = 0
    
    Prision.map = 66
    Libertad.map = 66
    
    Prision.X = 75
    Prision.Y = 47
    Libertad.X = 75
    Libertad.Y = 65
    
    
    LastBackup = Format(Now, "Short Time")
    Minutos = Format(Now, "Short Time")
    
    IniPath = App.Path & "\"
    DatPath = App.Path & "\Dat\"
    
    
    
    LevelSkill(1).LevelValue = 3
    LevelSkill(2).LevelValue = 5
    LevelSkill(3).LevelValue = 7
    LevelSkill(4).LevelValue = 10
    LevelSkill(5).LevelValue = 13
    LevelSkill(6).LevelValue = 15
    LevelSkill(7).LevelValue = 17
    LevelSkill(8).LevelValue = 20
    LevelSkill(9).LevelValue = 23
    LevelSkill(10).LevelValue = 25
    LevelSkill(11).LevelValue = 27
    LevelSkill(12).LevelValue = 30
    LevelSkill(13).LevelValue = 33
    LevelSkill(14).LevelValue = 35
    LevelSkill(15).LevelValue = 37
    LevelSkill(16).LevelValue = 40
    LevelSkill(17).LevelValue = 43
    LevelSkill(18).LevelValue = 45
    LevelSkill(19).LevelValue = 47
    LevelSkill(20).LevelValue = 50
    LevelSkill(21).LevelValue = 53
    LevelSkill(22).LevelValue = 55
    LevelSkill(23).LevelValue = 57
    LevelSkill(24).LevelValue = 60
    LevelSkill(25).LevelValue = 63
    LevelSkill(26).LevelValue = 65
    LevelSkill(27).LevelValue = 67
    LevelSkill(28).LevelValue = 70
    LevelSkill(29).LevelValue = 73
    LevelSkill(30).LevelValue = 75
    LevelSkill(31).LevelValue = 77
    LevelSkill(32).LevelValue = 80
    LevelSkill(33).LevelValue = 83
    LevelSkill(34).LevelValue = 85
    LevelSkill(35).LevelValue = 87
    LevelSkill(36).LevelValue = 90
    LevelSkill(37).LevelValue = 93
    LevelSkill(38).LevelValue = 95
    LevelSkill(39).LevelValue = 97
    LevelSkill(40).LevelValue = 100
    LevelSkill(41).LevelValue = 100
    LevelSkill(42).LevelValue = 100
    LevelSkill(43).LevelValue = 100
    LevelSkill(44).LevelValue = 100
    LevelSkill(45).LevelValue = 100
    LevelSkill(46).LevelValue = 100
    LevelSkill(47).LevelValue = 100
    LevelSkill(48).LevelValue = 100
    LevelSkill(49).LevelValue = 100
    LevelSkill(50).LevelValue = 100
    
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.Drow) = "Drow"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    
    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Worker) = "Trabajador"
    ListaClases(eClass.Pirat) = "Pirata"
    
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Tactica en combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.talar) = "Talar"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"
    SkillsNames(eSkill.Equitacion) = "Equitacion"
    SkillsNames(eSkill.Resistencia) = "Resistencia Magica"
    
    ListaAtributos(eAtributos.Fuerza) = "Fuerza"
    ListaAtributos(eAtributos.Agilidad) = "Agilidad"
    ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
    ListaAtributos(eAtributos.Carisma) = "Carisma"
    ListaAtributos(eAtributos.Constitucion) = "Constitucion"
    
    
    frmCargando.Show
    
    'Call PlayWaveAPI(App.Path & "\wav\harp3.wav")
    
     Expc = val(GetVar(IniPath & "Server.ini", "INIT", "Expc"))
    Oroc = val(GetVar(IniPath & "Server.ini", "INIT", "Oroc"))
    
    frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    IniPath = App.Path & "\"
    CharPath = App.Path & "\Charfile\"
    
    'Bordes del mapa
    MinXBorder = XMinMapSize + (XWindow \ 2)
    MaxXBorder = XMaxMapSize - (XWindow \ 2)
    MinYBorder = YMinMapSize + (YWindow \ 2)
    MaxYBorder = YMaxMapSize - (YWindow \ 2)
    DoEvents
    
    frmCargando.Label1(2).Caption = "Iniciando Arrays..."
    
    Call LoadGuildsDB
    
    Call LoadQuests
    
    Call CargarSpawnList
    Call CargarForbidenWords
    '¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
    frmCargando.Label1(2).Caption = "Cargando Server.ini"
    
    MaxUsers = 0
    Call LoadSini
    Call CargaApuestas
    
    '*************************************************
    frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
    Call CargaNpcsDat
    '*************************************************
    
    frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
    'Call LoadOBJData
    Call LoadOBJData
        
    frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
    Call CargarHechizos
        
        
    frmCargando.Label1(2).Caption = "Cargando Objetos de Herrería"
    Call LoadArmasHerreria
    Call LoadArmadurasHerreria
    
    frmCargando.Label1(2).Caption = "Cargando Objetos de Carpintería"
    Call LoadObjCarpintero
    
    frmCargando.Label1(2).Caption = "Cargando Balance.Dat"
    Call LoadBalance    '4/01/08 Pablo ToxicWaste
    
    frmCargando.Label1(2).Caption = "Cargando ArmadurasFaccionarias.dat"
    Call LoadArmadurasFaccion
    
    If BootDelBackUp Then
        
        frmCargando.Label1(2).Caption = "Cargando BackUp"
        Call CargarBackUp
    Else
        frmCargando.Label1(2).Caption = "Cargando Mapas"
        Call LoadMapData
    End If
    
    
    Call SonidosMapas.LoadSoundMapInfo

    Call generateMatrix(MATRIX_INITIAL_MAP)
    
    'Comentado porque hay worldsave en ese mapa!
    'Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    Dim LoopC As Integer
    
    'Resetea las conexiones de los usuarios
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
        Set UserList(LoopC).incomingData = New clsByteQueue
        Set UserList(LoopC).outgoingData = New clsByteQueue
    Next LoopC
    
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    With frmMain
        .AutoSave.Enabled = True
        .tLluvia.Enabled = True
        .tPiqueteC.Enabled = True
        .GameTimer.Enabled = True
        .tLluviaEvent.Enabled = True
        .FX.Enabled = True
        .Auditoria.Enabled = True
        .KillLog.Enabled = True
        .TIMER_AI.Enabled = True
        .npcataca.Enabled = True
        '.AutoGP.Enabled = True
    End With
    
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    'Configuracion de los sockets
    
    Call SecurityIp.InitIpTables(1000)
    
#If UsarQueSocket = 1 Then
    
    If LastSockListen >= 0 Then Call apiclosesocket(LastSockListen) 'Cierra el socket de escucha
    Call IniciaWsApi(frmMain.hWnd)
    SockListen = ListenForConnect(Puerto, hWndMsg, "")
        
    If SockListen <> -1 Then
        Call WriteVar(IniPath & "Server.ini", "INIT", "LastSockListen", SockListen) ' Guarda el socket escuchando
    Else
        MsgBox "Ha ocurrido un error al iniciar el socket del Servidor.", vbCritical + vbOKOnly
    End If
    
#ElseIf UsarQueSocket = 0 Then
    
    frmCargando.Label1(2).Caption = "Configurando Sockets"
    
    frmMain.Socket2(0).AddressFamily = AF_INET
    frmMain.Socket2(0).Protocol = IPPROTO_IP
    frmMain.Socket2(0).SocketType = SOCK_STREAM
    frmMain.Socket2(0).Binary = False
    frmMain.Socket2(0).Blocking = False
    frmMain.Socket2(0).BufferSize = 2048
    
    Call ConfigListeningSocket(frmMain.Socket1, Puerto)
    
#ElseIf UsarQueSocket = 2 Then
    
    frmMain.Serv.Iniciar Puerto
    
#ElseIf UsarQueSocket = 3 Then
    
    frmMain.TCPServ.Encolar True
    frmMain.TCPServ.IniciarTabla 1009
    frmMain.TCPServ.SetQueueLim 51200
    frmMain.TCPServ.Iniciar Puerto
    
#End If
    
    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    Unload frmCargando
    
    'Log
    Dim N As Integer
    N = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #N
    Print #N, Date & " " & time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
    Close #N
    
    'Ocultar
    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)
    End If
    
    tInicioServer = GetTickCount() And &H7FFFFFFF
    Call InicializaEstadisticas


End Sub

Function FileExist(ByVal file As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************

    FileExist = LenB(dir$(file, FileType)) <> 0
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
End Function

Function MapaValido(ByVal map As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    MapaValido = map >= 1 And map <= NumMaps
End Function

Sub MostrarNumUsers()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    frmMain.CantUsuarios.Caption = "Número de usuarios jugando: " & NumUsers

End Sub
Public Sub LogCheats(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\ANTICHEAT.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & desc
    Close #nfile
    
    Exit Sub

errhandler:

End Sub


Public Sub LogCriticEvent(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & desc
    Close #nfile
    
    Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoReal(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
    Print #nfile, desc
    Close #nfile
    
    Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoCaos(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
    Print #nfile, desc
    Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogIndex(ByVal index As Integer, ByVal desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\" & index & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & desc
    Close #nfile
    
    Exit Sub

errhandler:

End Sub


Public Sub LogError(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & desc
    Close #nfile
    
    Exit Sub

errhandler:

End Sub

Public Sub LogStatic(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & desc
    Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogTarea(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile(1) ' obtenemos un canal
    Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & desc
    Close #nfile

Exit Sub

errhandler:


End Sub


Public Sub LogClanes(ByVal Str As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Str
    Close #nfile

End Sub

Public Sub LogIP(ByVal Str As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\IP.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Str
    Close #nfile

End Sub


Public Sub LogDesarrollo(ByVal Str As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Str
    Close #nfile

End Sub

Public Sub LogGM(Nombre As String, texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
    Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
    
    Exit Sub

errhandler:

End Sub

Public Sub LogDropItemImportante(texto As String)

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
    Open App.Path & "\logs\ITEMS-IMPORTANTES.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
    
    Exit Sub

errhandler:

End Sub


Public Sub LogUserConect(Nombre As String, texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
    Open App.Path & "\logsuser\" & Nombre & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
    
    Exit Sub

errhandler:

End Sub
Public Sub LogT0Error(texto As String)

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
    Open App.Path & "\logs\Seguridad\T0\ERROR.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
    
    Exit Sub

errhandler:

End Sub
Public Sub LogT0Ban(Nombre As String, texto As String)

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
    Open App.Path & "\logs\Seguridad\T0\Baneos\" & Nombre & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
    
    Exit Sub

errhandler:

End Sub
Public Sub LogT0Unban(Nombre As String, texto As String)

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
    Open App.Path & "\logs\Seguridad\T0\Unbaneos\" & Nombre & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
    
    Exit Sub

errhandler:

End Sub
Public Sub LogFotoDenuncia(Nombre As String, texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************ç

On Error GoTo errhandler
    Dim nfile As Integer
    
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\Foto-Denuncias.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
    
    Exit Sub

errhandler:

End Sub

Public Sub LogAsesinato(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler
    Dim nfile As Integer
    
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
    
    Exit Sub

errhandler:

End Sub
Public Sub logVentaCasa(ByVal texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

errhandler:

End Sub
Public Sub LogHackAttemp(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

errhandler:

End Sub

Public Sub LogCheating(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\CH.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
    
    Exit Sub

errhandler:

End Sub


Public Sub LogCriticalHackAttemp(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

errhandler:

End Sub

Public Sub LogAntiCheat(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, ""
    Close #nfile
    
    Exit Sub

errhandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim Arg As String
    Dim i As Integer
    
    
    For i = 1 To 33
    
    Arg = ReadField(i, cad, 44)
    
    If LenB(Arg) = 0 Then Exit Function
    
    Next i
    
    ValidInputNP = True

End Function


Sub Restart()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

    If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."
    
    Dim LoopC As Long
  
#If UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    frmMain.Socket1.Startup
      
    frmMain.Socket2(0).Cleanup
    frmMain.Socket2(0).Startup

#ElseIf UsarQueSocket = 1 Then

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 2 Then

#End If

    For LoopC = 1 To MaxUsers
        Call CloseSocket(LoopC)
    Next
    
    'Initialize statistics!!
    Call Statistics.Initialize
    
    For LoopC = 1 To UBound(UserList())
        Set UserList(LoopC).incomingData = Nothing
        Set UserList(LoopC).outgoingData = Nothing
    Next LoopC
    
    ReDim UserList(1 To MaxUsers) As User
    
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
        Set UserList(LoopC).incomingData = New clsByteQueue
        Set UserList(LoopC).outgoingData = New clsByteQueue
    Next LoopC
    
    LastUser = 0
    NumUsers = 0
    
    Call FreeNPCs
    Call FreeCharIndexes
    
    Call LoadSini
    
    Call ResetForums
    Call LoadOBJData
    
    Call LoadMapData
    
    Call CargarHechizos

#If UsarQueSocket = 0 Then

    '*****************Setup socket
    frmMain.Socket1.AddressFamily = AF_INET
    frmMain.Socket1.Protocol = IPPROTO_IP
    frmMain.Socket1.SocketType = SOCK_STREAM
    frmMain.Socket1.Binary = False
    frmMain.Socket1.Blocking = False
    frmMain.Socket1.BufferSize = 1024
    
    frmMain.Socket2(0).AddressFamily = AF_INET
    frmMain.Socket2(0).Protocol = IPPROTO_IP
    frmMain.Socket2(0).SocketType = SOCK_STREAM
    frmMain.Socket2(0).Blocking = False
    frmMain.Socket2(0).BufferSize = 2048
    
    'Escucha
    frmMain.Socket1.LocalPort = val(Puerto)
    frmMain.Socket1.listen

#ElseIf UsarQueSocket = 1 Then

#ElseIf UsarQueSocket = 2 Then

#End If

    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    
    'Log it
    Dim N As Integer
    N = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #N
    Print #N, Date & " " & time & " servidor reiniciado."
    Close #N
    
    'Ocultar
    
    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)
    End If

  
End Sub


Public Function Intemperie(ByVal Userindex As Integer) As Boolean
'**************************************************************
'Author: Unknown
'Last Modify Date: 15/11/2009
'15/11/2009: ZaMa - La lluvia no quita stamina en las arenas.
'23/11/2009: ZaMa - Optimizacion de codigo.
'**************************************************************

    With UserList(Userindex)
        If MapInfo(.Pos.map).Zona <> "DUNGEON" Then
            If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger <> 1 And _
               MapData(.Pos.map, .Pos.X, .Pos.Y).trigger <> 2 And _
               MapData(.Pos.map, .Pos.X, .Pos.Y).trigger <> 4 Then Intemperie = False
               
        Else
            Intemperie = False
        End If
    End With
    
    
    'En las arenas no te afecta la lluvia
    
    If IsArena(Userindex) Then Intemperie = False
End Function

Public Sub EfectoLluvia(ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    If UserList(Userindex).flags.UserLogged Then
      '  If Intemperie(UserIndex) Then
            Call FlushBuffer(Userindex)
      '  End If
    End If
    
    Exit Sub
errhandler:
    LogError ("Error en EfectoLluvia")
End Sub


Public Sub TiempoInvocacion(ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        With UserList(Userindex)
            If .MascotasIndex(i) > 0 Then
                If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                   Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia = _
                   Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia - 1
                   If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(.MascotasIndex(i), 0)
                End If
            End If
        End With
    Next i
End Sub

Public Sub EfectoFrio(ByVal Userindex As Integer)
'***************************************************
'Autor: Unkonwn
'Last Modification: 23/11/2009
'If user is naked and it's in a cold map, take health points from him
'23/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************
    Dim modifi As Integer
   
   With UserList(Userindex)
        If .Counters.Frio < IntervaloFrio Then
            .Counters.Frio = .Counters.Frio + 1
        Else
            If MapInfo(.Pos.map).Terreno = Dungeon Then
                Call WriteConsoleMsg(Userindex, "¡¡Estás muriendo de frío, abrigate o morirás!!", FontTypeNames.FONTTYPE_INFO)
                modifi = Porcentaje(.Stats.MaxHp, 5)
                .Stats.MinHp = .Stats.MinHp - modifi
                
                If .Stats.MinHp < 1 Then
                    Call WriteConsoleMsg(Userindex, "¡¡Has muerto de frío!!", FontTypeNames.FONTTYPE_INFO)
                    .Stats.MinHp = 0
                    Call UserDie(Userindex)
                End If
                
                Call WriteUpdateHP(Userindex)
            Else
                modifi = Porcentaje(.Stats.MaxSta, 5)
                Call QuitarSta(Userindex, modifi)
                Call WriteUpdateSta(Userindex)
            End If
            
            .Counters.Frio = 0
        End If
    End With
End Sub

''
' Maneja  el efecto del estado atacable
'
' @param UserIndex  El index del usuario a ser afectado por el estado atacable
'

Public Sub EfectoEstadoAtacable(ByVal Userindex As Integer)
'******************************************************
'Author: ZaMa
'Last Update: 18/09/2010 (ZaMa)
'18/09/2010: ZaMa - Ahora se activa el seguro cuando dejas de ser atacable.
'******************************************************

    ' Si ya paso el tiempo de penalizacion
    If Not IntervaloEstadoAtacable(Userindex) Then
        ' Deja de poder ser atacado
        UserList(Userindex).flags.AtacablePor = 0
        
        ' Activo el seguro si deja de estar atacable
        If Not UserList(Userindex).flags.Seguro Then
            Call WriteMultiMessage(Userindex, eMessages.SafeModeOn)
        End If
        
        ' Send nick normal
        Call RefreshCharStatus(Userindex)
    End If
    
End Sub

''
' Maneja el tiempo y el efecto del mimetismo
'
' @param UserIndex  El index del usuario a ser afectado por el mimetismo
'

Public Sub EfectoMimetismo(ByVal Userindex As Integer)
'******************************************************
'Author: Unknown
'Last Update: 12/01/2010 (ZaMa)
'12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
'******************************************************
    Dim Barco As ObjData
    
    With UserList(Userindex)
        If .Counters.Mimetismo < IntervaloInvisible Then
            .Counters.Mimetismo = .Counters.Mimetismo + 1
        Else
            'restore old char
            Call WriteConsoleMsg(Userindex, "Recuperas tu apariencia normal.", FontTypeNames.FONTTYPE_INFO)
            
            If .flags.Navegando Then
                If .flags.Muerto = 0 Then
                    If .Faccion.ArmadaReal = 1 Then
                        .Char.body = iFragataReal
                    ElseIf .Faccion.FuerzasCaos = 1 Then
                        .Char.body = iFragataCaos
                    Else
                        Barco = ObjData(UserList(Userindex).Invent.BarcoObjIndex)
                        If criminal(Userindex) Then
                            If Barco.Ropaje = iBarca Then .Char.body = iBarcaPk
                            If Barco.Ropaje = iGalera Then .Char.body = iGaleraPk
                            If Barco.Ropaje = iGaleon Then .Char.body = iGaleonPk
                        Else
                            If Barco.Ropaje = iBarca Then .Char.body = iBarcaCiuda
                            If Barco.Ropaje = iGalera Then .Char.body = iGaleraCiuda
                            If Barco.Ropaje = iGaleon Then .Char.body = iGaleonCiuda
                        End If
                    End If
                Else
                    .Char.body = iFragataFantasmal
                End If
                
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
                 If UserList(Userindex).flags.Montando = 1 Then
                        .Char.body = ObjData(UserList(Userindex).Invent.MonturaObjIndex).Ropaje
                        .Char.Head = .OrigChar.Head
                        .Char.WeaponAnim = NingunArma
                        .Char.ShieldAnim = NingunEscudo
                        .Char.CascoAnim = .Char.CascoAnim
                End If
            Else
                .Char.body = .CharMimetizado.body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            End If
                  '     .MimetizmoName = vbNullString
            With .Char
                Call ChangeUserChar(Userindex, .body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
            End With
           
         '   ResetearName userIndex
           
            
            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
            ' Se fue el efecto del mimetismo, puede ser atacado por npcs
            .flags.Ignorado = False
        End If
    End With
End Sub

Public Sub EfectoInvisibilidad(ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/09/2010 (ZaMa)
'16/09/2010: ZaMa - Al perder el invi cuando navegas, no se manda el mensaje de sacar invi (ya estas visible).
'***************************************************

    With UserList(Userindex)
        If .Counters.Invisibilidad < IntervaloInvisible Then
            .Counters.Invisibilidad = .Counters.Invisibilidad + 1
        Else
            .Counters.Invisibilidad = RandomNumber(-100, 100) ' Invi variable :D
            .flags.invisible = 0
            If .flags.Oculto = 0 Then
                Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                
                ' Si navega ya esta visible..
                If Not .flags.Navegando = 1 Then
                    Call SetInvisible(Userindex, .Char.CharIndex, False)
                End If
                
            End If
        End If
    End With

End Sub


Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With Npclist(NpcIndex)
        If .Contadores.Paralisis > 0 Then
            .Contadores.Paralisis = .Contadores.Paralisis - 1
        Else
            .flags.Paralizado = 0
            .flags.Inmovilizado = 0
        End If
    End With

End Sub

Public Sub EfectoCegueEstu(ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(Userindex)
        If .Counters.Ceguera > 0 Then
            .Counters.Ceguera = .Counters.Ceguera - 1
        Else
            If .flags.Ceguera = 1 Then
                .flags.Ceguera = 0
                Call WriteBlindNoMore(Userindex)
            End If
            If .flags.Estupidez = 1 Then
                .flags.Estupidez = 0
                Call WriteDumbNoMore(Userindex)
            End If
        
        End If
    End With

End Sub


Public Sub EfectoParalisisUser(ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(Userindex)
        If .Counters.Paralisis > 0 Then
            .Counters.Paralisis = .Counters.Paralisis - 1
        Else
            .flags.Paralizado = 0
            .flags.Inmovilizado = 0
            '.Flags.AdministrativeParalisis = 0
            Call WriteParalizeOK(Userindex)
        End If
    End With

End Sub

Public Sub RecStamina(ByVal Userindex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(Userindex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 1 And _
           MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 2 And _
           MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 4 Then Exit Sub
        
        
        Dim massta As Integer
        If .Stats.MinSta < .Stats.MaxSta Then
            If .Counters.STACounter < Intervalo Then
                .Counters.STACounter = .Counters.STACounter + 1
            Else
                EnviarStats = True
                .Counters.STACounter = 0
                If .flags.Desnudo Then Exit Sub 'Desnudo no sube energía. (ToxicWaste)
               
                massta = RandomNumber(1, Porcentaje(.Stats.MaxSta, 5))
                .Stats.MinSta = .Stats.MinSta + massta
                If .Stats.MinSta > .Stats.MaxSta Then
                    .Stats.MinSta = .Stats.MaxSta
                End If
            End If
        End If
    End With
    
End Sub

Public Sub EfectoVeneno(ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim N As Integer
    
    With UserList(Userindex)
        If .Counters.Veneno < IntervaloVeneno Then
          .Counters.Veneno = .Counters.Veneno + 1
        Else
          Call WriteConsoleMsg(Userindex, "Estás envenenado, si no te curas morirás.", FontTypeNames.FONTTYPE_VENENO)
          .Counters.Veneno = 0
          N = RandomNumber(1, 5)
          .Stats.MinHp = .Stats.MinHp - N
          If .Stats.MinHp < 1 Then Call UserDie(Userindex)
          Call WriteUpdateHP(Userindex)
          Call WriteUpdateFollow(Userindex)
        End If
    End With

End Sub

Public Sub DuracionPociones(ByVal Userindex As Integer)
'***************************************************
'Author: ??????
'Last Modification: 11/27/09 (Budi)
'Cuando se pierde el efecto de la poción updatea fz y agi (No me gusta que ambos atributos aunque se haya modificado solo uno, pero bueno :p)
'***************************************************
    With UserList(Userindex)
        'Controla la duracion de las pociones
        If .flags.DuracionEfecto > 0 Then
           .flags.DuracionEfecto = .flags.DuracionEfecto - 1
           If .flags.DuracionEfecto = 0 Then
                .flags.TomoPocion = False
                .flags.TipoPocion = 0
                'volvemos los atributos al estado normal
                Dim loopx As Integer
                
                For loopx = 1 To NUMATRIBUTOS
                    .Stats.UserAtributos(loopx) = .Stats.UserAtributosBackUP(loopx)
                Next loopx
                
                Call WriteUpdateStrenghtAndDexterity(Userindex)
           End If
        End If
    End With

End Sub

Public Sub HambreYSed(ByVal Userindex As Integer, ByRef fenviarAyS As Boolean)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(Userindex)
        If Not .flags.Privilegios And PlayerType.User Then Exit Sub
        
        'Sed
        If .Stats.MinAGU > 0 Then
            If .Counters.AGUACounter < IntervaloSed Then
                .Counters.AGUACounter = .Counters.AGUACounter + 1
            Else
                .Counters.AGUACounter = 0
                .Stats.MinAGU = .Stats.MinAGU - 10
                
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
                .Stats.MinHam = .Stats.MinHam - 10
                If .Stats.MinHam <= 0 Then
                       .Stats.MinHam = 0
                       .flags.Hambre = 1
                End If
                fenviarAyS = True
            End If
        End If
    End With

End Sub

Public Sub Sanar(ByVal Userindex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(Userindex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 1 And _
           MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 2 And _
           MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 4 Then Exit Sub
        
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
                Call WriteConsoleMsg(Userindex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
                EnviarStats = True
            End If
        End If
    End With

End Sub

Public Sub CargaNpcsDat()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim npcfile As String
    
    npcfile = DatPath & "NPCs.dat"
    Call LeerNPCs.Initialize(npcfile)
End Sub
Function eventAttack(ByVal attackerIndex As Integer, _
                     ByVal victimIndex As Integer) As Boolean
 
        '
        ' @ amishar
   
        If UserList(attackerIndex).reto2Data.reto_used = True Then
                If Mod_2vs2.can_Attack(attackerIndex, victimIndex) = False Then
                        Call WriteConsoleMsg(attackerIndex, "No puedes atacar a tu compañero.", FontTypeNames.FONTTYPE_INFO)
                        eventAttack = False
 
                        Exit Function
 
                End If
        End If
   
        eventAttack = True
 
End Function
 
Sub eventLoad()
 
        '
        ' @ amishar
   
        Call Mod_2vs2.initRetoData2
 
End Sub
 
Sub eventDie(ByVal Userindex As Integer)
 
        '
        ' @ amishar
   
        With UserList(Userindex)
         
                If reto_List(.reto2Data.reto_Index).used_ring Then
                        Call Mod_2vs2.user_die_reto(Userindex)
                End If
         
        End With
 
End Sub
 
Sub eventClose(ByVal Userindex As Integer)
 
        '
        ' @ amishar
 
        With UserList(Userindex)
         
                If reto_List(.reto2Data.reto_Index).used_ring Then
                        Call Mod_2vs2.disconnect_Reto(Userindex)
                End If
         
        End With
 
End Sub
Sub PasarSegundo()

On Error GoTo errhandler
    Dim i As Long

     If CuentaRegresivaTimer > 0 Then
        If CuentaRegresivaTimer > 1 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("" & CuentaRegresivaTimer - 1, FontTypeNames.FONTTYPE_GUILD))
        Else
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("YA!", FontTypeNames.FONTTYPE_FIGHT))
        End If
        CuentaRegresivaTimer = CuentaRegresivaTimer - 1
    End If
    
    'Retos 1vs1
    Retos1vs1.PasaSegundo
    
    'CUENTA REGRESIVA DEATHAMTCH
    If DeathMatch.Activo Then
       Call ModDeath.Cuenta
    End If
    
           'Hay death?
    If JDH.Activo Then
       Call Mod_Jdh.Cuenta
    End If

    Call Mod_2vs2.loop_reto
    
    For i = 1 To LastUser

    Call Mod_2vs2.user_retoLoop(CInt(i))

    'Cerrar usuario
            If UserList(i).Counters.Saliendo Then
                UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
                If UserList(i).Counters.Salir <= 0 Then
                   Call WriteConsoleMsg(i, "Gracias por jugar Desterium AO", FontTypeNames.FONTTYPE_INFO)
                    Call WriteDisconnect(i)
                    Call FlushBuffer(i)
                    
                    Call CloseSocket(i)
                End If
            End If
        
        
          If UserList(i).Counters.Denuncia > 0 Then
       UserList(i).Counters.Denuncia = UserList(i).Counters.Denuncia - 1
       If UserList(i).Counters.Denuncia <= 0 Then
        UserList(i).Counters.Denuncia = 0
     End If
       End If
        
        
    Next i
Exit Sub

errhandler:
    Call LogError("Error en PasarSegundo. Err: " & Err.description & " - " & Err.Number & " - UserIndex: " & i)
    Resume Next
End Sub
 
Public Function ReiniciarAutoUpdate() As Double
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    'WorldSave
    Call ES.DoBackUp

    'commit experiencias
    Call mdParty.ActualizaExperiencias

    'Guardar Pjs
    Call GuardarUsuarios
    
    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    Unload frmMain

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

    
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i, CharPath & UCase$(UserList(i).name) & ".chr", False)
        End If
    Next i
    
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
    haciendoBK = True
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Grabando Personajes", FontTypeNames.FONTTYPE_SERVER))
    
    Dim i As Integer

    
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i, CharPath & UCase$(UserList(i).name) & ".chr", False)
        End If
    Next i
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    haciendoBK = False
End Sub


Sub InicializaEstadisticas()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim Ta As Long
    Ta = GetTickCount() And &H7FFFFFFF
    
    Call EstadisticasWeb.Inicializa(frmMain.hWnd)
    Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
    Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    Call EstadisticasWeb.Informar(UPTIME_SERVER, (Ta - tInicioServer) / 1000)
    Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)

End Sub

Public Sub FreeNPCs()
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
End Sub

Public Sub FreeCharIndexes()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all char indexes
'***************************************************
    ' Free all char indexes (set them all to 0)
    Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))
End Sub
Public Sub elefectofrio(ByVal Userindex As Integer)
   Dim modifi As Integer
    With UserList(Userindex)
        If .Counters.Frio < IntervaloFrio Then
            .Counters.Frio = .Counters.Frio + 1
        Else
        If MapInfo(.Pos.map).Terreno = Nieve Then
                Call WriteConsoleMsg(Userindex, "¡¡Estás muriendo de frío, abrigate o morirás!!", FontTypeNames.FONTTYPE_INFO)
                modifi = Porcentaje(.Stats.MaxHp, 5)
                .Stats.MinHp = .Stats.MinHp - modifi
                
                If .Stats.MinHp < 1 Then
                    Call WriteConsoleMsg(Userindex, "¡¡Has muerto de frío!!", FontTypeNames.FONTTYPE_INFO)
                    .Stats.MinHp = 0
                    Call UserDie(Userindex)
                End If
                
                Call WriteUpdateHP(Userindex)
                Call WriteUpdateFollow(Userindex)
            Else
                modifi = Porcentaje(.Stats.MaxSta, 5)
                Call QuitarSta(Userindex, modifi)
                Call WriteUpdateSta(Userindex)
            
            End If
            .Counters.Frio = 0
        End If
    End With
End Sub
Public Sub LogMD5(LOG As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
 
On Error GoTo errhandler
 
    Dim mA As Integer
    mA = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\MD5-NO OK.log" For Append Shared As #mA
    Print #mA, Date & " " & time & " " & LOG
    Close #mA
   
    Exit Sub
 
errhandler:
 
End Sub
Public Sub Viajes(ByVal Userindex As Integer, ByVal Lug As Byte)
    With UserList(Userindex)
   
        Select Case Lug
            Case 0 'Dungeon 1
                If .Stats.GLD < 10000 Then
                    Call WriteConsoleMsg(Userindex, "No tienes suficiente dinero, necesitas 10.000 monedas de oro para ir hacia el Muelle de Lindos.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
               
                .Stats.GLD = .Stats.GLD - 10000: WriteUpdateGold Userindex
               
                Call WarpUserChar(Userindex, 62, 83, 45, False)
                Call WriteConsoleMsg(Userindex, "Has viajado hacia el Muelle de Lindos, se te han descontado 10.000 monedas de oro.", FontTypeNames.FONTTYPE_INFO)
                
            Case 1 'Dungeon 2
                If .Stats.GLD < 7000 Then
                    Call WriteConsoleMsg(Userindex, "No tienes suficiente dinero, necesitas 7.000 monedas de oro para viajar hacia el Muelle de Nix.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
               
                .Stats.GLD = .Stats.GLD - 7000: WriteUpdateGold Userindex
               
                Call WarpUserChar(Userindex, 34, 27, 79, False)
   Call WriteConsoleMsg(Userindex, "Has viajado hacia el Muelle de Nix, se te han descontado 7.000 monedas de oro.", FontTypeNames.FONTTYPE_INFO)
   
       Case 2 'Dungeon 2
                If .Stats.GLD < 15000 Then
                    Call WriteConsoleMsg(Userindex, "No tienes suficiente dinero, necesitas 15.000 monedas de oro para viajar hacia el Muelle de Arghâl.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
               
                .Stats.GLD = .Stats.GLD - 15000: WriteUpdateGold Userindex
               
                Call WarpUserChar(Userindex, 150, 36, 30, False)
                
   Call WriteConsoleMsg(Userindex, "Has viajado hacia el Muelle de Arghâl, se te han descontado 15.000 monedas de oro.", FontTypeNames.FONTTYPE_INFO)
   
     Case 3 'Dungeon 2
                If .Stats.GLD < 15000 Then
                    Call WriteConsoleMsg(Userindex, "No tienes suficiente dinero, necesitas 15.000 monedas de oro para viajar hacia el Muelle de Banderbill.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
               
                .Stats.GLD = .Stats.GLD - 15000: WriteUpdateGold Userindex
               
                Call WarpUserChar(Userindex, 61, 66, 67, False)
   Call WriteConsoleMsg(Userindex, "Has viajado hacia el Muelle de Banderbill, se te han descontado 15.000 monedas de oro.", FontTypeNames.FONTTYPE_INFO)
   
       Case 4 'Dungeon 2
                If .Stats.GLD < 350000 Then
                    Call WriteConsoleMsg(Userindex, "No tienes suficiente dinero, necesitas 350.000 monedas de oro para viajar hacia el Fuerte Pretoriano.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
               
                .Stats.GLD = .Stats.GLD - 350000: WriteUpdateGold Userindex
               
                Call WarpUserChar(Userindex, 196, 66, 94, False)
   Call WriteConsoleMsg(Userindex, "Has viajado hacia el Fuerte Pretoriano, se te han descontado 350.000 monedas de oro.", FontTypeNames.FONTTYPE_INFO)
   
        End Select
    End With
End Sub
