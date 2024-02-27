Attribute VB_Name = "ES"
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

Public Administradores As clsIniManager

Public Sub CargarSpawnList()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim N As Integer, LoopC As Integer
10        N = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
20        ReDim SpawnList(N) As tCriaturasEntrenador
30        For LoopC = 1 To N
40            SpawnList(LoopC).NpcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & LoopC))
50            SpawnList(LoopC).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & LoopC)
60        Next LoopC
          
End Sub

Function EsAdmin(ByRef Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 27/03/2011
'27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
'***************************************************
    EsAdmin = (val(Administradores.GetValue("Admin", Name)) = 1)
End Function

Function EsDios(ByRef Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 27/03/2011
'27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
'***************************************************
    EsDios = (val(Administradores.GetValue("Dios", Name)) = 1)
End Function

Function EsSemiDios(ByRef Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 27/03/2011
'27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
'***************************************************
    EsSemiDios = (val(Administradores.GetValue("SemiDios", Name)) = 1)
End Function

Function EsGmEspecial(ByRef Name As String) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 27/03/2011
'27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
'***************************************************
    EsGmEspecial = (val(Administradores.GetValue("Especial", Name)) = 1)
End Function

Function EsConsejero(ByRef Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 27/03/2011
'27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
'***************************************************
    EsConsejero = (val(Administradores.GetValue("Consejero", Name)) = 1)
End Function

Function EsRolesMaster(ByRef Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 27/03/2011
'27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
'***************************************************
    EsRolesMaster = (val(Administradores.GetValue("RM", Name)) = 1)
End Function

Public Function EsGmChar(ByRef Name As String) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 27/03/2011
'Returns true if char is administrative user.
'***************************************************
    
    Dim EsGm As Boolean
    
    ' Admin?
    EsGm = EsAdmin(Name)
    ' Dios?
    If Not EsGm Then EsGm = EsDios(Name)
    ' Semidios?
    If Not EsGm Then EsGm = EsSemiDios(Name)
    ' Consejero?
    If Not EsGm Then EsGm = EsConsejero(Name)

    EsGmChar = EsGm

End Function


Public Sub loadAdministrativeUsers()
'Admines     => Admin
'Dioses      => Dios
'SemiDioses  => SemiDios
'Especiales  => Especial
'Consejeros  => Consejero
'RoleMasters => RM

    'Si esta mierda tuviese array asociativos el código sería tan lindo.
    Dim buf As Integer
    Dim i As Long
    Dim Name As String
    Dim Temp As String
    
    ' Public container
    Set Administradores = New clsIniManager
    
    ' Server ini info file
    Dim ServerIni As clsIniManager
    Set ServerIni = New clsIniManager
    
    Call ServerIni.Initialize(IniPath & "Server.ini")
    
       
    ' Admines
    buf = val(ServerIni.GetValue("INIT", "Admines"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Admines", "Admin" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Admin", Name, "1")

    Next i
    
    ' Dioses
    buf = val(ServerIni.GetValue("INIT", "Dioses"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Dioses", "Dios" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Dios", Name, "1")
        
    Next i
    
    ' Especiales
    buf = val(ServerIni.GetValue("INIT", "Especiales"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Especiales", "Especial" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Especial", Name, "1")
        
    Next i
    
    ' SemiDioses
    buf = val(ServerIni.GetValue("INIT", "SemiDioses"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("SemiDioses", "SemiDios" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("SemiDios", Name, "1")
        
    Next i
    
    ' Consejeros
    buf = val(ServerIni.GetValue("INIT", "Consejeros"))
        
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Consejeros", "Consejero" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Consejero", Name, "1")
        
    Next i
    
    ' RolesMasters
    buf = val(ServerIni.GetValue("INIT", "RolesMasters"))
        
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("RolesMasters", "RM" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("RM", Name, "1")
    Next i
    
    ' Rangos ficticios para colores y visuales
    buf = val(ServerIni.GetValue("RANGOS", "Ultimo"))
    
    ReDim RangeGm(0 To buf) As tRange
    
    For i = 1 To buf
        Temp = ServerIni.GetValue("RANGOS", i)
        
        RangeGm(i).Name = ReadField(1, Temp, Asc("-"))
        RangeGm(i).Tag = ReadField(2, Temp, Asc("-"))
        'RangeGm(i).Color
    Next i
    
    
    Set ServerIni = Nothing
    
End Sub
Public Function GetRangeData(ByVal UserName As String) As String
    Dim A As Long
    
    For A = LBound(RangeGm) To UBound(RangeGm)
        If RangeGm(A).Name = UserName Then
            GetRangeData = RangeGm(A).Tag
            Exit For
        End If
    Next A
End Function
Public Function GetCharPrivs(ByRef UserName As String) As PlayerType
      '****************************************************
      'Author: ZaMa
      'Last Modification: 18/11/2010
      'Reads the user's charfile and retrieves its privs.
      '***************************************************

          Dim Privs As PlayerType

10        If EsAdmin(UserName) Then
20            Privs = PlayerType.Admin
              
30        ElseIf EsDios(UserName) Then
40            Privs = PlayerType.Dios

50        ElseIf EsSemiDios(UserName) Then
60            Privs = PlayerType.SemiDios
              
70        ElseIf EsConsejero(UserName) Then
80            Privs = PlayerType.Consejero
          
90        Else
100           Privs = PlayerType.User
110       End If

120       GetCharPrivs = Privs

End Function


Public Function TxtDimension(ByVal Name As String) As Long
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim N As Integer, cad As String, Tam As Long
10        N = FreeFile(1)
20        Open Name For Input As #N
30        Tam = 0
40        Do While Not EOF(N)
50            Tam = Tam + 1
60            Line Input #N, cad
70        Loop
80        Close N
90        TxtDimension = Tam
End Function

Public Sub CargarForbidenWords()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
          Dim N As Integer, i As Integer
20        N = FreeFile(1)
30        Open DatPath & "NombresInvalidos.txt" For Input As #N
          
40        For i = 1 To UBound(ForbidenNames)
50            Line Input #N, ForbidenNames(i)
60        Next i
          
70        Close N

End Sub

Public Sub CargarHechizos()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      '###################################################
      '#               ATENCION PELIGRO                  #
      '###################################################
      '
      '  ¡¡¡¡ NO USAR GetVar PARA LEER Hechizos.dat !!!!
      '
      'El que ose desafiar esta LEY, se las tendrá que ver
      'con migo. Para leer Hechizos.dat se deberá usar
      'la nueva clase clsLeerInis.
      '
      'Alejo
      '
      '###################################################

10    On Error GoTo Errhandler

20        If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."
          
          Dim Hechizo As Integer
          Dim Leer As New clsIniManager
          
30        Call Leer.Initialize(DatPath & "Hechizos.dat")
          
          'obtiene el numero de hechizos
40        NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))
          
50        ReDim Hechizos(1 To NumeroHechizos) As tHechizo
          
60        frmCargando.cargar.min = 0
70        frmCargando.cargar.max = NumeroHechizos
80        frmCargando.cargar.Value = 0
          
          'Llena la lista
90        For Hechizo = 1 To NumeroHechizos
100           With Hechizos(Hechizo)
110               .Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
120               .desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
130               .PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
                  
140               .HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
150               .TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
160               .PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
                  
170               .Tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
180               .WAV = val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
190               .FXgrh = val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
                  
200               .loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
                  
              '    .Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
                  
210               .SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
220               .MinHp = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
230               .MaxHp = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
                  
240               .SubeMana = val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
250               .MiMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
260               .MaMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
                  
270               .SubeSta = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
280               .MinSta = val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
290               .MaxSta = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
                  
300               .SubeHam = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
310               .MinHam = val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
320               .MaxHam = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
                  
330               .SubeSed = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
340               .MinSed = val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
350               .MaxSed = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
                  
360               .SubeAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
370               .MinAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
380               .MaxAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
                  
390               .SubeFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
400               .MinFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
410               .MaxFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
                  
420               .SubeCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
430               .MinCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
440               .MaxCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
                  
                  
450               .Invisibilidad = val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
460               .Paraliza = val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
470               .Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
480               .RemoverParalisis = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
490               .RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
500               .RemueveInvisibilidadParcial = val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
                  
                  
510               .CuraVeneno = val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
520               .Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
530               .Maldicion = val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
540               .RemoverMaldicion = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
550               .Bendicion = val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
560               .Revivir = val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
                  
570               .Ceguera = val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
580               .Estupidez = val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
                  
590               .Warp = val(Leer.GetValue("Hechizo" & Hechizo, "Warp"))
                  
600               .Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
610               .NumNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
620               .cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
630               .Mimetiza = val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
                  
              '    .Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
              '    .ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
                  
640               .MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
650               .ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
                  
                  'Barrin 30/9/03
660               .StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
                  
670               .Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
680               frmCargando.cargar.Value = frmCargando.cargar.Value + 1
                  
690               .NeedStaff = val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
700               .StaffAffected = CBool(val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
710           End With
720       Next Hechizo
          
730       Set Leer = Nothing
          
740       Exit Sub

Errhandler:
750       MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.Description
       
End Sub
Public Sub DoBackUp()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        haciendoBK = True
          Dim i As Integer
          
          
          
          ' Lo saco porque elimina elementales y mascotas - Maraxus
          ''''''''''''''lo pongo aca x sugernecia del yind
          'For i = 1 To LastNPC
          '    If Npclist(i).flags.NPCActive Then
          '        If Npclist(i).Contadores.TiempoExistencia > 0 Then
          '            Call MuereNpc(i, 0)
          '        End If
          '    End If
          'Next i
          '''''''''''/'lo pongo aca x sugernecia del yind
          
          
          
20        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
          
          
30        Call WorldSave
40        Call modGuilds.v_RutinaElecciones
50        Call ResetCentinelaInfo     'Reseteamos al centinela
          
          
60        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
          
          'Call EstadisticasWeb.Informar(EVENTO_NUEVO_CLAN, 0)
          
70        haciendoBK = False
          
          'Log
80        On Error Resume Next
          Dim nfile As Integer
90        nfile = FreeFile ' obtenemos un canal
100       Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
110       Print #nfile, Date & " " & time
120       Close #nfile
End Sub

Public Sub GrabarMapa(ByVal map As Long, ByRef MAPFILE As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: 12/01/2011
      '10/08/2010 - Pato: Implemento el clsByteBuffer para el grabado de mapas
      '28/10/2010:ZaMa - Ahora no se hace backup de los pretorianos.
      '12/01/2011 - Amraphen: Ahora no se hace backup de NPCs prohibidos (Pretorianos, Mascotas, Invocados y Centinela)
      '***************************************************

10    On Error Resume Next
          Dim FreeFileMap As Long
          Dim FreeFileInf As Long
          Dim Y As Long
          Dim X As Long
          Dim ByFlags As Byte
          Dim LoopC As Long
          Dim MapWriter As clsByteBuffer
          Dim InfWriter As clsByteBuffer
          Dim IniManager As clsIniManager
          Dim NpcInvalido As Boolean
          
20        Set MapWriter = New clsByteBuffer
30        Set InfWriter = New clsByteBuffer
40        Set IniManager = New clsIniManager
          
50        If FileExist(MAPFILE & ".map", vbNormal) Then
60            Kill MAPFILE & ".map"
70        End If
          
80        If FileExist(MAPFILE & ".inf", vbNormal) Then
90            Kill MAPFILE & ".inf"
100       End If
          
          'Open .map file
110       FreeFileMap = FreeFile
120       Open MAPFILE & ".Map" For Binary As FreeFileMap
          
130       Call MapWriter.initializeWriter(FreeFileMap)
          
          'Open .inf file
140       FreeFileInf = FreeFile
150       Open MAPFILE & ".Inf" For Binary As FreeFileInf
          
160       Call InfWriter.initializeWriter(FreeFileInf)
          
          'map Header
170       Call MapWriter.putInteger(MapInfo(map).MapVersion)
              
180       Call MapWriter.putString(MiCabecera.desc, False)
190       Call MapWriter.putLong(MiCabecera.CRC)
200       Call MapWriter.putLong(MiCabecera.MagicWord)
          
210       Call MapWriter.putDouble(0)
          
          'inf Header
220       Call InfWriter.putDouble(0)
230       Call InfWriter.putInteger(0)
          
          'Write .map file
240       For Y = YMinMapSize To YMaxMapSize
250           For X = XMinMapSize To XMaxMapSize
260               With MapData(map, X, Y)
270                   ByFlags = 0
                      
280                   If .Blocked Then ByFlags = ByFlags Or 1
290                   If .Graphic(2) Then ByFlags = ByFlags Or 2
300                   If .Graphic(3) Then ByFlags = ByFlags Or 4
310                   If .Graphic(4) Then ByFlags = ByFlags Or 8
320                   If .trigger Then ByFlags = ByFlags Or 16
                      
330                   Call MapWriter.putByte(ByFlags)
                      
340                   Call MapWriter.putInteger(.Graphic(1))
                      
350                   For LoopC = 2 To 4
360                       If .Graphic(LoopC) Then _
                              Call MapWriter.putInteger(.Graphic(LoopC))
370                   Next LoopC
                      
380                   If .trigger Then _
                          Call MapWriter.putInteger(CInt(.trigger))
                      
                      '.inf file
390                   ByFlags = 0
                      
400                   If .ObjInfo.ObjIndex > 0 Then
410                      If ObjData(.ObjInfo.ObjIndex).ObjType = eOBJType.otFogata Then
420                           .ObjInfo.ObjIndex = 0
430                           .ObjInfo.Amount = 0
440                       End If
450                   End If
          
460                   If .TileExit.map Then ByFlags = ByFlags Or 1
                      
                      ' No hacer backup de los NPCs inválidos (Pretorianos, Mascotas, Invocados y Centinela)
470                   If .NpcIndex Then
480                       NpcInvalido = (Npclist(.NpcIndex).NPCtype = eNPCType.pretoriano) Or (Npclist(.NpcIndex).MaestroUser > 0)
                          
490                       If Not NpcInvalido Then ByFlags = ByFlags Or 2
500                   End If
                      
510                   If .ObjInfo.ObjIndex Then ByFlags = ByFlags Or 4
                      
520                   Call InfWriter.putByte(ByFlags)
                      
530                   If .TileExit.map Then
540                       Call InfWriter.putInteger(.TileExit.map)
550                       Call InfWriter.putInteger(.TileExit.X)
560                       Call InfWriter.putInteger(.TileExit.Y)
570                   End If
                      
580                   If .NpcIndex And Not NpcInvalido Then _
                          Call InfWriter.putInteger(Npclist(.NpcIndex).Numero)
                      
590                   If .ObjInfo.ObjIndex Then
600                       Call InfWriter.putInteger(.ObjInfo.ObjIndex)
610                       Call InfWriter.putInteger(.ObjInfo.Amount)
620                   End If
                      
630                   NpcInvalido = False
640               End With
650           Next X
660       Next Y
          
670       Call MapWriter.saveBuffer
680       Call InfWriter.saveBuffer
          
          'Close .map file
690       Close FreeFileMap

          'Close .inf file
700       Close FreeFileInf
          
710       Set MapWriter = Nothing
720       Set InfWriter = Nothing

730       With MapInfo(map)
              'write .dat file
740           Call IniManager.ChangeValue("Mapa" & map, "Name", .Name)
750           Call IniManager.ChangeValue("Mapa" & map, "MusicNum", .Music)
760           Call IniManager.ChangeValue("Mapa" & map, "MagiaSinefecto", .MagiaSinEfecto)
770           Call IniManager.ChangeValue("Mapa" & map, "InviSinEfecto", .InviSinEfecto)
780           Call IniManager.ChangeValue("Mapa" & map, "ResuSinEfecto", .ResuSinEfecto)
790           Call IniManager.ChangeValue("Mapa" & map, "StartPos", .StartPos.map & "-" & .StartPos.X & "-" & .StartPos.Y)
800           Call IniManager.ChangeValue("Mapa" & map, "OnDeathGoTo", .OnDeathGoTo.map & "-" & .OnDeathGoTo.X & "-" & .OnDeathGoTo.Y)

          
810           Call IniManager.ChangeValue("Mapa" & map, "Terreno", .Terreno)
820           Call IniManager.ChangeValue("Mapa" & map, "Zona", .Zona)
830           Call IniManager.ChangeValue("Mapa" & map, "Restringir", .Restringir)
840           Call IniManager.ChangeValue("Mapa" & map, "BackUp", Str(.BackUp))
          
850           If .Pk Then
860               Call IniManager.ChangeValue("Mapa" & map, "Pk", "0")
870           Else
880               Call IniManager.ChangeValue("Mapa" & map, "Pk", "1")
890           End If
              
900           Call IniManager.ChangeValue("Mapa" & map, "OcultarSinEfecto", .OcultarSinEfecto)
910           Call IniManager.ChangeValue("Mapa" & map, "InvocarSinEfecto", .InvocarSinEfecto)
920           Call IniManager.ChangeValue("Mapa" & map, "NoEncriptarMP", .NoEncriptarMP)
930           Call IniManager.ChangeValue("Mapa" & map, "RoboNpcsPermitido", .RoboNpcsPermitido)
          
940           Call IniManager.DumpFile(MAPFILE & ".dat")
950       End With
          
960       Set IniManager = Nothing
End Sub
Sub LoadArmasHerreria()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim N As Integer, lc As Integer
          
10        N = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))
          
20        ReDim Preserve ArmasHerrero(1 To N) As Integer
          
30        For lc = 1 To N
40            ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
50        Next lc

End Sub

Sub LoadArmadurasHerreria()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim N As Integer, lc As Integer
          
10        N = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))
          
20        ReDim Preserve ArmadurasHerrero(1 To N) As Integer
          
30        For lc = 1 To N
40            ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
50        Next lc

End Sub

Sub LoadBalance()
      '***************************************************
      'Author: Unknown
      'Last Modification: 15/04/2010
      '15/04/2010: ZaMa - Agrego recompensas faccionarias.
      '***************************************************

          Dim i As Long
          
          'Modificadores de Clase
10        For i = 1 To NUMCLASES
20            With ModClase(i)
30                .Evasion = val(GetVar(DatPath & "Balance.dat", "MODEVASION", ListaClases(i)))
40                .AtaqueArmas = val(GetVar(DatPath & "Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
50                .AtaqueProyectiles = val(GetVar(DatPath & "Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
60                .AtaqueWrestling = val(GetVar(DatPath & "Balance.dat", "MODATAQUEWRESTLING", ListaClases(i)))
70                .DañoArmas = val(GetVar(DatPath & "Balance.dat", "MODDAÑOARMAS", ListaClases(i)))
80                .DañoProyectiles = val(GetVar(DatPath & "Balance.dat", "MODDAÑOPROYECTILES", ListaClases(i)))
90                .DañoWrestling = val(GetVar(DatPath & "Balance.dat", "MODDAÑOWRESTLING", ListaClases(i)))
100               .Escudo = val(GetVar(DatPath & "Balance.dat", "MODESCUDO", ListaClases(i)))
110           End With
120       Next i
          
          'Modificadores de Raza
130       For i = 1 To NUMRAZAS
140           With ModRaza(i)
150               .Fuerza = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Fuerza"))
160               .Agilidad = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Agilidad"))
170               .Inteligencia = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Inteligencia"))
180               .Carisma = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Carisma"))
190               .Constitucion = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Constitucion"))
200           End With
210       Next i
          
          'Modificadores de Vida
220       For i = 1 To NUMCLASES
230           ModVida(i) = val(GetVar(DatPath & "Balance.dat", "MODVIDA", ListaClases(i)))
240       Next i
          
          'Distribución de Vida
250       For i = 1 To 5
260           DistribucionEnteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "E" + CStr(i)))
270       Next i
280       For i = 1 To 4
290           DistribucionSemienteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "S" + CStr(i)))
300       Next i
          
          'Extra
310       PorcentajeRecuperoMana = val(GetVar(DatPath & "Balance.dat", "EXTRA", "PorcentajeRecuperoMana"))

          ' Recompensas faccionarias
330       For i = 1 To NUM_RANGOS_FACCION
340           RecompensaFacciones(i - 1) = val(GetVar(DatPath & "Balance.dat", "RECOMPENSAFACCION", "Rango" & i))
350       Next i
          
End Sub

Sub LoadObjCarpintero()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim N As Integer, lc As Integer
          
10        N = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))
          
20        ReDim Preserve ObjCarpintero(1 To N) As Integer
          
30        For lc = 1 To N
40            ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
50        Next lc

End Sub



Sub LoadOBJData()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      '###################################################
      '#               ATENCION PELIGRO                  #
      '###################################################
      '
      '¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
      '
      'El que ose desafiar esta LEY, se las tendrá que ver
      'con migo. Para leer desde el OBJ.DAT se deberá usar
      'la nueva clase clsLeerInis.
      '
      'Alejo
      '
      '###################################################

      'Call LogTarea("Sub LoadOBJData")

10    On Error GoTo Errhandler

20        If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."
          
          '*****************************************************************
          'Carga la lista de objetos
          '*****************************************************************
          Dim Object As Integer
          Dim Leer As New clsIniManager
          
30        Call Leer.Initialize(DatPath & "Obj.dat")
          
          'obtiene el numero de obj
40        NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))
          
50        frmCargando.cargar.min = 0
60        frmCargando.cargar.max = NumObjDatas
70        frmCargando.cargar.Value = 0
          
          
80        ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
          
          
          'Llena la lista
90        For Object = 1 To NumObjDatas
100           With ObjData(Object)
110               .Name = Leer.GetValue("OBJ" & Object, "Name")
                  
                  'Pablo (ToxicWaste) Log de Objetos.
120               .LOG = val(Leer.GetValue("OBJ" & Object, "Log"))
130               .NoLog = val(Leer.GetValue("OBJ" & Object, "NoLog"))
                  '07/09/07
                  
140               .GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
150               If .GrhIndex = 0 Then
160                   .GrhIndex = .GrhIndex
170               End If
                  
180               .ObjType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
                  
190               .Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))
                  
200               .MagiaSkill = val(Leer.GetValue("OBJ" & Object, "MagiaSkill"))
210               .RMSkill = val(Leer.GetValue("OBJ" & Object, "RMSkill"))
220   .ArmaSkill = val(Leer.GetValue("OBJ" & Object, "WeaponSkill"))
230   .EscudoSkill = val(Leer.GetValue("OBJ" & Object, "EscudoSkill"))
240   .ArmaduraSkill = val(Leer.GetValue("OBJ" & Object, "ArmaduraSkill"))
250   .ArcoSkill = val(Leer.GetValue("OBJ" & Object, "ArcoSkill"))
260   .DagaSkill = val(Leer.GetValue("OBJ" & Object, "DagaSkill"))
270   .Monturasskill = val(Leer.GetValue("OBJ" & Object, "MonturasSkill"))
280   .QuitaEnergia = val(Leer.GetValue("OBJ" & Object, "Energia"))
290               .Quince = val(Leer.GetValue("OBJ" & Object, "Quince"))
300               .Treinta = val(Leer.GetValue("OBJ" & Object, "Treinta"))
310               .HM = val(Leer.GetValue("OBJ" & Object, "HM"))
320               .UM = val(Leer.GetValue("OBJ" & Object, "UM"))
330               .MM = val(Leer.GetValue("OBJ" & Object, "MM"))
340               .VIP = val(Leer.GetValue("OBJ" & Object, "VIP"))
350               .VIPB = val(Leer.GetValue("OBJ" & Object, "VIPB"))
360               .VIPP = val(Leer.GetValue("OBJ" & Object, "VIPP"))
                  
370               Select Case .ObjType
                      Case eOBJType.otarmadura
380                       .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
390                       .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
400                       .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
410                       .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
420                       .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
430                       .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                      
440                   Case eOBJType.otescudo
450                       .ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
460                       .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
470                       .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
480                       .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
490                       .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
500                       .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
510                       .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                      
520                   Case eOBJType.otcasco
530                       .CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
540                       .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
550                       .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
560                       .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
570                       .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
580                       .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
590                       .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                      
600                   Case eOBJType.otWeapon
610                       .WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
620                       .Apuñala = val(Leer.GetValue("OBJ" & Object, "Apuñala"))
630                       .Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
640                       .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
650                       .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
660                       .proyectil = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
670                       .Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
680                       .StaffPower = val(Leer.GetValue("OBJ" & Object, "StaffPower"))
690                       .StaffDamageBonus = val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
700                       .Refuerzo = val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
                          
710                       .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
720                       .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
730                       .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
740                       .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
750                       .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
760                       .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                          
770                       .WeaponRazaEnanaAnim = val(Leer.GetValue("OBJ" & Object, "RazaEnanaAnim"))
                      
780                   Case eOBJType.otInstrumentos
790                       .Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
800                       .Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
810                       .Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
                          'Pablo (ToxicWaste)
820                       .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
830                       .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                      
840                   Case eOBJType.otMinerales
850                       .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                      
860                   Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
870                       .IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
880                       .IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
890                       .IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
                      
900                   Case otPociones
910                       .TipoPocion = val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
920                       .MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
930                       .MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
940                       .DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
                      
                      Case eOBJType.otGemaTelep
                          .TelepMap = val(Leer.GetValue("OBJ" & Object, "TelepMap"))
                          .TelepX = val(Leer.GetValue("OBJ" & Object, "TelepX"))
                          .TelepY = val(Leer.GetValue("OBJ" & Object, "TelepY"))
                          .TelepTime = val(Leer.GetValue("OBJ" & Object, "TelepTime"))
                          
950                   Case eOBJType.otBarcos
960                       .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
970                       .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
980                       .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                      
990                   Case eOBJType.otAnilloNpc
1000                      .NpcTipo = val(Leer.GetValue("OBJ" & Object, "NpcTipo"))
                          
1010                  Case eOBJType.otMonturas
1020                      .Velocidad = val(Leer.GetValue("OBJ" & Object, "Velocidad"))
1030                      .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
1040                      .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
1050                      .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                      
1060                    Case eOBJType.otMonturasDraco
1070                       .Velocidad = val(Leer.GetValue("OBJ" & Object, "Velocidad"))
1080                      .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
1090                      .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
1100                      .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                      
1110                  Case eOBJType.otFlechas
1120                      .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
1130                      .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
1140                      .Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
1150                      .Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
                          
1160                  Case eOBJType.otAnillo 'Pablo (ToxicWaste)
1170                      .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
1180                      .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
1190                      .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
1200                      .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
1210                      .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
1220                      .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                          
1230                  Case eOBJType.otTeleport
1240                      .Radio = val(Leer.GetValue("OBJ" & Object, "Radio"))
                          
1250                  Case eOBJType.otMochilas
1260                      .MochilaType = val(Leer.GetValue("OBJ" & Object, "MochilaType"))
                          
1270                  Case eOBJType.otForos
1280                      Call AddForum(Leer.GetValue("OBJ" & Object, "ID"))
                          
1290              End Select
                  
1300              .Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
1310              .HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
                  
1320              .LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
                  
1330              .MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
                  
1340              .MaxHp = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
1350              .MinHp = val(Leer.GetValue("OBJ" & Object, "MinHP"))
                  
1360              .Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
1370              .Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
                  
1380              .MinHam = val(Leer.GetValue("OBJ" & Object, "MinHam"))
1390              .MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
                  
1400              .MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
1410              .MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
1420              .def = (.MinDef + .MaxDef) / 2
                  
1430              .RazaEnana = val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
1440              .RazaDrow = val(Leer.GetValue("OBJ" & Object, "RazaDrow"))
1450              .RazaElfa = val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
1460              .RazaGnoma = val(Leer.GetValue("OBJ" & Object, "RazaGnoma"))
1470              .RazaHumana = val(Leer.GetValue("OBJ" & Object, "RazaHumana"))
                  
1480              .valor = val(Leer.GetValue("OBJ" & Object, "Valor"))
1490              ObjData(Object).Premium = val(Leer.GetValue("OBJ" & Object, "PREMIUM"))
1500              .copaS = val(Leer.GetValue("OBJ" & Object, "DsP"))
1510              .Eldhir = val(Leer.GetValue("OBJ" & Object, "Eldhires"))
                  
1520              .Crucial = val(Leer.GetValue("OBJ" & Object, "Crucial"))
                  
1530              .Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta"))
1540              If .Cerrada = 1 Then
1550                  .Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
1560                  .clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
1570              End If
                  
                  'Puertas y llaves
1580              .clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
                  
1590              .texto = Leer.GetValue("OBJ" & Object, "Texto")
1600              .GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
                  
1610              .Agarrable = val(Leer.GetValue("OBJ" & Object, "Agarrable"))
1620              .ForoID = Leer.GetValue("OBJ" & Object, "ID")
                  
1630              .Acuchilla = val(Leer.GetValue("OBJ" & Object, "Acuchilla"))
                  
1640              .Guante = val(Leer.GetValue("OBJ" & Object, "Guante"))
                  
                  'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico
                  Dim i As Integer
                  Dim N As Integer
                  Dim S As String
1650              For i = 1 To NUMCLASES
1660                  S = UCase$(Leer.GetValue("OBJ" & Object, "CP" & i))
1670                  N = 1
1680                  Do While LenB(S) > 0 And UCase$(ListaClases(N)) <> S
1690                      N = N + 1
1700                  Loop
1710                  .ClaseProhibida(i) = IIf(LenB(S) > 0, N, 0)
1720              Next i
                  
1730              .DefensaMagicaMax = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
1740              .DefensaMagicaMin = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
                  
1750              .SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
                  
1760              If .SkCarpinteria > 0 Then _
                      .Madera = val(Leer.GetValue("OBJ" & Object, "Madera"))
1770                  .MaderaElfica = val(Leer.GetValue("OBJ" & Object, "MaderaElfica"))
                  
                  'Bebidas
1780              .MinSta = val(Leer.GetValue("OBJ" & Object, "MinST"))
                  
1790              .NoSeCae = val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
                  
1800              .Upgrade = val(Leer.GetValue("OBJ" & Object, "Upgrade"))
                  
1810              frmCargando.cargar.Value = frmCargando.cargar.Value + 1
1820          End With
1830      Next Object
          
          
1840      Set Leer = Nothing
          
          ' Inicializo los foros faccionarios
1850      Call AddForum(FORO_CAOS_ID)
1860      Call AddForum(FORO_REAL_ID)
          
1870      Exit Sub

Errhandler:
1880      MsgBox "error cargando objetos " & Err.Number & ": " & Err.Description & " En el Objeto: " & Object


End Sub

Sub LoadUserStats(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)
      '*************************************************
      'Author: Unknown
      'Last modified: 11/19/2009
      '11/19/2009: Pato - Load the EluSkills and ExpSkills
      '*************************************************
      Dim LoopC As Long
      Dim tmpStr As String

10    With UserList(UserIndex)
              
20        For LoopC = 0 To MAX_LOGROS
30            tmpStr = UserFile.GetValue("STATS", "LOGROS")
40            .Logros(LoopC) = val(ReadField(LoopC + 1, tmpStr, Asc("-")))
50        Next LoopC
          
          
60        With .Stats
70            For LoopC = 1 To NUMATRIBUTOS
80                .UserAtributos(LoopC) = CInt(UserFile.GetValue("ATRIBUTOS", "AT" & LoopC))
90                .UserAtributosBackUP(LoopC) = .UserAtributos(LoopC)
100           Next LoopC
              
110           For LoopC = 1 To NUMSKILLS
120               .UserSkills(LoopC) = CInt(UserFile.GetValue("SKILLS", "SK" & LoopC))
130               .EluSkills(LoopC) = CInt(UserFile.GetValue("SKILLS", "ELUSK" & LoopC))
140               .ExpSkills(LoopC) = CInt(UserFile.GetValue("SKILLS", "EXPSK" & LoopC))
150           Next LoopC
              
160           For LoopC = 1 To MAXUSERHECHIZOS
170               .UserHechizos(LoopC) = CInt(UserFile.GetValue("Hechizos", "H" & LoopC))
180           Next LoopC

              
190           .Gld = CLng(UserFile.GetValue("STATS", "GLD"))
200           .Banco = CLng(UserFile.GetValue("STATS", "BANCO"))
              
210           .MaxHp = CInt(UserFile.GetValue("STATS", "MaxHP"))
220           .MinHp = CInt(UserFile.GetValue("STATS", "MinHP"))
              
230           .MinSta = CInt(UserFile.GetValue("STATS", "MinSTA"))
240           .MaxSta = CInt(UserFile.GetValue("STATS", "MaxSTA"))
              
250           .MaxMAN = CInt(UserFile.GetValue("STATS", "MaxMAN"))
260           .MinMAN = CInt(UserFile.GetValue("STATS", "MinMAN"))
              
270           .MaxHIT = CInt(UserFile.GetValue("STATS", "MaxHIT"))
280           .MinHIT = CInt(UserFile.GetValue("STATS", "MinHIT"))
              
290           .MaxAGU = CByte(UserFile.GetValue("STATS", "MaxAGU"))
300           .MinAGU = CByte(UserFile.GetValue("STATS", "MinAGU"))
              
310           .MaxHam = CByte(UserFile.GetValue("STATS", "MaxHAM"))
320           .MinHam = CByte(UserFile.GetValue("STATS", "MinHAM"))
              
330           .SkillPts = CInt(UserFile.GetValue("STATS", "SkillPtsLibres"))
              
340           .Exp = CDbl(UserFile.GetValue("STATS", "EXP"))
350           .ELU = CLng(UserFile.GetValue("STATS", "ELU"))
360           .ELV = CByte(UserFile.GetValue("STATS", "ELV"))
              
              
370           .UsuariosMatados = CLng(UserFile.GetValue("MUERTES", "UserMuertes"))
380           .NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))
              
390           .RetosGanados = CInt(UserFile.GetValue("RETOS", "RetosGanados"))
400           .RetosPerdidos = CInt(UserFile.GetValue("RETOS", "RetosPerdidos"))
410           .OroGanado = CLng(UserFile.GetValue("RETOS", "OroGanado"))
420           .OroPerdido = CLng(UserFile.GetValue("RETOS", "OroPerdido"))
              
430           .TorneosGanados = CInt(UserFile.GetValue("TORNEOS", "TorneosGanados"))
              .Points = CLng(UserFile.GetValue("STATS", "Canje"))
440       End With
          
450       With .flags
460           If CByte(UserFile.GetValue("CONSEJO", "PERTENECE")) Then _
                  .Privilegios = .Privilegios Or PlayerType.RoyalCouncil
              
470           If CByte(UserFile.GetValue("CONSEJO", "PERTENECECAOS")) Then _
                  .Privilegios = .Privilegios Or PlayerType.ChaosCouncil
480       End With
490   End With
End Sub

Sub LoadUserReputacion(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With UserList(UserIndex).Reputacion
20            .AsesinoRep = val(UserFile.GetValue("REP", "Asesino"))
30            .BandidoRep = val(UserFile.GetValue("REP", "Bandido"))
40            .BurguesRep = val(UserFile.GetValue("REP", "Burguesia"))
50            .LadronesRep = val(UserFile.GetValue("REP", "Ladrones"))
60            .NobleRep = val(UserFile.GetValue("REP", "Nobles"))
70            .PlebeRep = val(UserFile.GetValue("REP", "Plebe"))
80            .Promedio = val(UserFile.GetValue("REP", "Promedio"))
90        End With
          
End Sub

Sub LoadUserInit(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)
      '*************************************************
      'Author: Unknown
      'Last modified: 19/11/2006
      'Loads the Users records
      '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
      '23/01/2007 Pablo (ToxicWaste) - Quito CriminalesMatados de Stats porque era redundante.
      '*************************************************
          Dim LoopC As Long
          Dim ln As String
          
10        With UserList(UserIndex)
20            With .Faccion
30                .ArmadaReal = CByte(UserFile.GetValue("FACCIONES", "EjercitoReal"))
40                .FuerzasCaos = CByte(UserFile.GetValue("FACCIONES", "EjercitoCaos"))
50                .CiudadanosMatados = CLng(UserFile.GetValue("FACCIONES", "CiudMatados"))
60                .CriminalesMatados = CLng(UserFile.GetValue("FACCIONES", "CrimMatados"))
70                .RecibioArmaduraCaos = CByte(UserFile.GetValue("FACCIONES", "rArCaos"))
80                .RecibioArmaduraReal = CByte(UserFile.GetValue("FACCIONES", "rArReal"))
90                .RecibioExpInicialCaos = CByte(UserFile.GetValue("FACCIONES", "rExCaos"))
100               .RecibioExpInicialReal = CByte(UserFile.GetValue("FACCIONES", "rExReal"))
110               .RecompensasCaos = CLng(UserFile.GetValue("FACCIONES", "recCaos"))
120               .RecompensasReal = CLng(UserFile.GetValue("FACCIONES", "recReal"))
130               .Reenlistadas = CByte(UserFile.GetValue("FACCIONES", "Reenlistadas"))
140               .NivelIngreso = CInt(UserFile.GetValue("FACCIONES", "NivelIngreso"))
150               .FechaIngreso = UserFile.GetValue("FACCIONES", "FechaIngreso")
160               .MatadosIngreso = CInt(UserFile.GetValue("FACCIONES", "MatadosIngreso"))
170               .NextRecompensa = CInt(UserFile.GetValue("FACCIONES", "NextRecompensa"))
180           End With
              
              Dim i As Integer
              
              .Account = UserFile.GetValue("INIT", "ACCOUNT")
              
300           For i = 1 To MAX_GUILDS_DISOLVED
310               .GuildDisolved(i) = CInt(UserFile.GetValue("DISOLVED", "Disolved" & i))
320           Next i
              
330           With .flags
340               .Muerto = CByte(UserFile.GetValue("FLAGS", "Muerto"))
350                .Premium = CByte(UserFile.GetValue("FLAGS", "PREMIUM"))
360               .Escondido = CByte(UserFile.GetValue("FLAGS", "Escondido"))
370               .BonosHP = CByte(UserFile.GetValue("FLAGS", "BonosHP"))
380               .Oro = CByte(UserFile.GetValue("FLAGS", "ORO"))
390               .Plata = CByte(UserFile.GetValue("FLAGS", "PLATA"))
400               .Bronce = CByte(UserFile.GetValue("FLAGS", "BRONCE"))
410               .DiosTerrenal = CByte(UserFile.GetValue("FLAGS", "DiosTerrenal"))
420               .Hambre = CByte(UserFile.GetValue("FLAGS", "Hambre"))
430               .Sed = CByte(UserFile.GetValue("FLAGS", "Sed"))
440               .Desnudo = CByte(UserFile.GetValue("FLAGS", "Desnudo"))
450               .Navegando = CByte(UserFile.GetValue("FLAGS", "Navegando"))
460               .Montando = CByte(UserFile.GetValue("FLAGS", "Montando"))
470               .Envenenado = CByte(UserFile.GetValue("FLAGS", "Envenenado"))
480               .Paralizado = CByte(UserFile.GetValue("FLAGS", "Paralizado"))
                  'Matrix
490               .lastMap = CInt(UserFile.GetValue("FLAGS", "LastMap"))
500           End With
              
510           If .flags.Paralizado = 1 Then
520               .Counters.Paralisis = IntervaloParalizado
530           End If
              
              
540           .Counters.Pena = CLng(UserFile.GetValue("COUNTERS", "Pena"))
              .Counters.TimeTelep = CLng(UserFile.GetValue("COUNTERS", "TimeTelep"))
550           .Counters.AsignedSkills = CByte(val(UserFile.GetValue("COUNTERS", "SkillsAsignados")))
              
560           .Email = UserFile.GetValue("CONTACTO", "Email")
570           .Pin = UserFile.GetValue("INIT", "Pin")
580           .Genero = UserFile.GetValue("INIT", "Genero")
590           .clase = UserFile.GetValue("INIT", "Clase")
600           .raza = UserFile.GetValue("INIT", "Raza")
610           .Hogar = UserFile.GetValue("INIT", "Hogar")
620           .Char.Heading = CInt(UserFile.GetValue("INIT", "Heading"))
              
              
630           With .OrigChar
640               .Head = CInt(UserFile.GetValue("INIT", "Head"))
650               .body = CInt(UserFile.GetValue("INIT", "Body"))
660               .WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
670               .ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
680               .CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))
                  
690               .Heading = eHeading.SOUTH
700           End With
              
        #If ConUpTime Then
710               .UpTime = CLng(UserFile.GetValue("INIT", "UpTime"))
        #End If
              
720     If UserList(UserIndex).flags.Muerto = 0 Then
730       UserList(UserIndex).Char = UserList(UserIndex).OrigChar
      'Cuerpo Legión & Normal
740   ElseIf UserList(UserIndex).Faccion.FuerzasCaos <> 0 Then 'Es caos
750       UserList(UserIndex).Char.body = iCuerpoMuertoCrimi
760       UserList(UserIndex).Char.Head = iCabezaMuertoCrimi
770       UserList(UserIndex).Char.WeaponAnim = NingunArma
780       UserList(UserIndex).Char.ShieldAnim = NingunEscudo
790       UserList(UserIndex).Char.CascoAnim = NingunCasco
800   Else
810       UserList(UserIndex).Char.body = iCuerpoMuerto
820       UserList(UserIndex).Char.Head = iCabezaMuerto
830       UserList(UserIndex).Char.WeaponAnim = NingunArma
840       UserList(UserIndex).Char.ShieldAnim = NingunEscudo
850       UserList(UserIndex).Char.CascoAnim = NingunCasco
860   End If
              
870           .desc = UserFile.GetValue("INIT", "Desc")
              
880           .Pos.map = CInt(ReadField(1, UserFile.GetValue("INIT", "Position"), 45))
890           .Pos.X = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
900           .Pos.Y = CInt(ReadField(3, UserFile.GetValue("INIT", "Position"), 45))
              
910           .Invent.NroItems = CInt(UserFile.GetValue("Inventory", "CantidadItems"))
              
              '[KEVIN]--------------------------------------------------------------------
              '***********************************************************************************
920           .BancoInvent.NroItems = CInt(UserFile.GetValue("BancoInventory", "CantidadItems"))
              'Lista de objetos del banco
930           For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
940               ln = UserFile.GetValue("BancoInventory", "Obj" & LoopC)
950               .BancoInvent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
960               .BancoInvent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
970           Next LoopC
              '------------------------------------------------------------------------------------
              '[/KEVIN]*****************************************************************************
              
              
              'Lista de objetos
980           For LoopC = 1 To MAX_INVENTORY_SLOTS
990               ln = UserFile.GetValue("Inventory", "Obj" & LoopC)
1000              .Invent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
1010              .Invent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
1020              .Invent.Object(LoopC).Equipped = CByte(ReadField(3, ln, 45))
1030          Next LoopC
              
              'Obtiene el indice-objeto del arma
1040          .Invent.WeaponEqpSlot = CByte(UserFile.GetValue("Inventory", "WeaponEqpSlot"))
1050          If .Invent.WeaponEqpSlot > 0 Then
1060              .Invent.WeaponEqpObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex
1070          End If
              
              'Obtiene el indice-objeto montura
1080          .Invent.MonturaSlot = CByte(UserFile.GetValue("Inventory", "MonturaSlot"))
1090          If .Invent.MonturaSlot > 0 Then
1100          .Invent.MonturaObjIndex = .Invent.Object(.Invent.MonturaSlot).ObjIndex
1110          End If
              
1120          .Invent.AnilloNpcSlot = CByte(UserFile.GetValue("Inventory", "AnilloNpcSlot"))
1130          If .Invent.AnilloNpcSlot > 0 Then
1140              .Invent.AnilloNpcObjIndex = .Invent.Object(.Invent.AnilloNpcSlot).ObjIndex
1150          End If
              
              'Obtiene el indice-objeto del armadura
1160          .Invent.ArmourEqpSlot = CByte(UserFile.GetValue("Inventory", "ArmourEqpSlot"))
1170          If .Invent.ArmourEqpSlot > 0 Then
1180              .Invent.ArmourEqpObjIndex = .Invent.Object(.Invent.ArmourEqpSlot).ObjIndex
1190              .flags.Desnudo = 0
1200          Else
1210              .flags.Desnudo = 1
1220          End If
              
              'Obtiene el indice-objeto del escudo
1230          .Invent.EscudoEqpSlot = CByte(UserFile.GetValue("Inventory", "EscudoEqpSlot"))
1240          If .Invent.EscudoEqpSlot > 0 Then
1250              .Invent.EscudoEqpObjIndex = .Invent.Object(.Invent.EscudoEqpSlot).ObjIndex
1260          End If
              
              'Obtiene el indice-objeto del casco
1270          .Invent.CascoEqpSlot = CByte(UserFile.GetValue("Inventory", "CascoEqpSlot"))
1280          If .Invent.CascoEqpSlot > 0 Then
1290              .Invent.CascoEqpObjIndex = .Invent.Object(.Invent.CascoEqpSlot).ObjIndex
1300          End If
              
              'Obtiene el indice-objeto barco
1310          .Invent.BarcoSlot = CByte(UserFile.GetValue("Inventory", "BarcoSlot"))
1320          If .Invent.BarcoSlot > 0 Then
1330              .Invent.BarcoObjIndex = .Invent.Object(.Invent.BarcoSlot).ObjIndex
1340          End If
              
              'Obtiene el indice-objeto municion
1350          .Invent.MunicionEqpSlot = CByte(UserFile.GetValue("Inventory", "MunicionSlot"))
1360          If .Invent.MunicionEqpSlot > 0 Then
1370              .Invent.MunicionEqpObjIndex = .Invent.Object(.Invent.MunicionEqpSlot).ObjIndex
1380          End If
              
              '[Alejo]
              'Obtiene el indice-objeto anilo
1390          .Invent.AnilloEqpSlot = CByte(UserFile.GetValue("Inventory", "AnilloSlot"))
1400          If .Invent.AnilloEqpSlot > 0 Then
1410              .Invent.AnilloEqpObjIndex = .Invent.Object(.Invent.AnilloEqpSlot).ObjIndex
1420          End If
              
1430          .Invent.MochilaEqpSlot = CByte(UserFile.GetValue("Inventory", "MochilaSlot"))
1440          If .Invent.MochilaEqpSlot > 0 Then
1450              .Invent.MochilaEqpObjIndex = .Invent.Object(.Invent.MochilaEqpSlot).ObjIndex
1460          End If
              
1470          .NroMascotas = CInt(UserFile.GetValue("MASCOTAS", "NroMascotas"))
              Dim NpcIndex As Integer
1480          For LoopC = 1 To MAXMASCOTAS
1490              .MascotasType(LoopC) = val(UserFile.GetValue("MASCOTAS", "MAS" & LoopC))
1500          Next LoopC
              
1510          ln = UserFile.GetValue("Guild", "GUILDINDEX")
1520          If IsNumeric(ln) Then
1530              .GuildIndex = CInt(ln)
1540          Else
1550              .GuildIndex = 0
1560          End If
1570      End With

End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim sSpaces As String ' This will hold the input that the program will retrieve
          Dim szReturn As String ' This will be the defaul value if the string is not found
            
10        szReturn = vbNullString
            
20        sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
            
            
30        GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
            
40        GetVar = RTrim$(sSpaces)
50        GetVar = Left$(GetVar, Len(GetVar) - 1)
        
End Function

Sub CargarBackUp()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."
          
          Dim map As Integer
          Dim TempInt As Integer
          Dim tFileName As String
          Dim npcfile As String
          
20        On Error GoTo man
              
30            NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
40            Call InitAreas
              
50            frmCargando.cargar.min = 0
60            frmCargando.cargar.max = NumMaps
70            frmCargando.cargar.Value = 0
              
80            MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
              
              
90            ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
100           ReDim MapInfo(1 To NumMaps) As MapInfo
              
110           For map = 1 To NumMaps
120               If val(GetVar(App.Path & MapPath & "Mapa" & map & ".Dat", "Mapa" & map, "BackUp")) <> 0 Then
130                   tFileName = App.Path & "\WorldBackUp\Mapa" & map
                      
140                   If Not FileExist(tFileName & ".*") Then 'Miramos que exista al menos uno de los 3 archivos, sino lo cargamos de la carpeta de los mapas
150                       tFileName = App.Path & MapPath & "Mapa" & map
160                   End If
170               Else
180                   tFileName = App.Path & MapPath & "Mapa" & map
190               End If
                  
200               Call CargarMapa(map, tFileName)
                  
210               frmCargando.cargar.Value = frmCargando.cargar.Value + 1
220               DoEvents
230           Next map
          
240       Exit Sub

man:
250       MsgBox ("Error durante la carga de mapas, el mapa " & map & " contiene errores")
260       Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)
       
End Sub

Sub LoadMapData()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."
          
          Dim map As Integer
          Dim TempInt As Integer
          Dim tFileName As String
          Dim npcfile As String
          
20        On Error GoTo man
              
30            NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
40            Call InitAreas
              
50            frmCargando.cargar.min = 0
60            frmCargando.cargar.max = NumMaps
70            frmCargando.cargar.Value = 0
              
80            MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
              
              
90            ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
100           ReDim MapInfo(1 To NumMaps) As MapInfo
                
110           For map = 1 To NumMaps
                  
120               tFileName = App.Path & MapPath & "Mapa" & map
130               Call CargarMapa(map, tFileName)
                  
140               frmCargando.cargar.Value = frmCargando.cargar.Value + 1
150               DoEvents
160           Next map
          
170       Exit Sub

man:
180       MsgBox ("Error durante la carga de mapas, el mapa " & map & " contiene errores")
190       Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)

End Sub

Public Sub CargarMapa(ByVal map As Long, ByRef MAPFl As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: 10/08/2010
      '10/08/2010 - Pato: Implemento el clsByteBuffer y el clsIniManager para la carga de mapa
      '***************************************************

10    On Error GoTo errh
          Dim hFile As Integer
          Dim X As Long
          Dim Y As Long
          Dim ByFlags As Byte
          Dim npcfile As String
          Dim Leer As clsIniManager
          Dim MapReader As clsByteBuffer
          Dim InfReader As clsByteBuffer
          Dim Buff() As Byte
          
20        Set MapReader = New clsByteBuffer
30        Set InfReader = New clsByteBuffer
40        Set Leer = New clsIniManager
          
50        npcfile = DatPath & "NPCs.dat"
          
60        hFile = FreeFile
          
          
70        Open MAPFl & ".map" For Binary As #hFile
80            Seek hFile, 1

90            ReDim Buff(LOF(hFile) - 1) As Byte
          
100           Get #hFile, , Buff
110       Close hFile
          
120       Call MapReader.initializeReader(Buff)

          'inf
130       Open MAPFl & ".inf" For Binary As #hFile
140           Seek hFile, 1

150           ReDim Buff(LOF(hFile) - 1) As Byte
          
160           Get #hFile, , Buff
170       Close hFile
          
180       Call InfReader.initializeReader(Buff)
          
          'map Header
190       MapInfo(map).MapVersion = MapReader.getInteger
          
200       MiCabecera.desc = MapReader.getString(Len(MiCabecera.desc))
210       MiCabecera.CRC = MapReader.getLong
220       MiCabecera.MagicWord = MapReader.getLong
          
230       Call MapReader.getDouble

          'inf Header
240       Call InfReader.getDouble
250       Call InfReader.getInteger

260       For Y = YMinMapSize To YMaxMapSize
270           For X = XMinMapSize To XMaxMapSize
280               With MapData(map, X, Y)
                      '.map file
290                   ByFlags = MapReader.getByte

300                   If ByFlags And 1 Then .Blocked = 1

310                   .Graphic(1) = MapReader.getInteger

                      'Layer 2 used?
320                   If ByFlags And 2 Then .Graphic(2) = MapReader.getInteger

                      'Layer 3 used?
330                   If ByFlags And 4 Then .Graphic(3) = MapReader.getInteger

                      'Layer 4 used?
340                   If ByFlags And 8 Then .Graphic(4) = MapReader.getInteger

                      'Trigger used?
350                   If ByFlags And 16 Then .trigger = MapReader.getInteger

                      '.inf file
360                   ByFlags = InfReader.getByte

370                   If ByFlags And 1 Then
380                       .TileExit.map = InfReader.getInteger
390                       .TileExit.X = InfReader.getInteger
400                       .TileExit.Y = InfReader.getInteger
410                   End If

420                   If ByFlags And 2 Then
                          'Get and make NPC
430                        .NpcIndex = InfReader.getInteger

440                       If .NpcIndex > 0 Then
                              'Si el npc debe hacer respawn en la pos
                              'original la guardamos
450                           If val(GetVar(npcfile, "NPC" & .NpcIndex, "PosOrig")) = 1 Then
460                               .NpcIndex = OpenNPC(.NpcIndex)
470                               Npclist(.NpcIndex).Orig.map = map
480                               Npclist(.NpcIndex).Orig.X = X
490                               Npclist(.NpcIndex).Orig.Y = Y
500                           Else
510                               .NpcIndex = OpenNPC(.NpcIndex)
520                           End If

530                           Npclist(.NpcIndex).Pos.map = map
540                           Npclist(.NpcIndex).Pos.X = X
550                           Npclist(.NpcIndex).Pos.Y = Y

560                           Call MakeNPCChar(True, 0, .NpcIndex, map, X, Y)
570                       End If
580                   End If

590                   If ByFlags And 4 Then
                          'Get and make Object
600                       .ObjInfo.ObjIndex = InfReader.getInteger
610                       .ObjInfo.Amount = InfReader.getInteger
620                   End If
630               End With
640           Next X
650       Next Y
          
660       Call Leer.Initialize(MAPFl & ".dat")
          
670       With MapInfo(map)
680           .Name = Leer.GetValue("Mapa" & map, "Name")
690           .Music = Leer.GetValue("Mapa" & map, "MusicNum")
700           .StartPos.map = val(ReadField(1, Leer.GetValue("Mapa" & map, "StartPos"), Asc("-")))
710           .StartPos.X = val(ReadField(2, Leer.GetValue("Mapa" & map, "StartPos"), Asc("-")))
720           .StartPos.Y = val(ReadField(3, Leer.GetValue("Mapa" & map, "StartPos"), Asc("-")))
              
730           .OnDeathGoTo.map = val(ReadField(1, Leer.GetValue("Mapa" & map, "OnDeathGoTo"), Asc("-")))
740           .OnDeathGoTo.X = val(ReadField(2, Leer.GetValue("Mapa" & map, "OnDeathGoTo"), Asc("-")))
750           .OnDeathGoTo.Y = val(ReadField(3, Leer.GetValue("Mapa" & map, "OnDeathGoTo"), Asc("-")))
              
760           .MagiaSinEfecto = val(Leer.GetValue("Mapa" & map, "MagiaSinEfecto"))
770           .InviSinEfecto = val(Leer.GetValue("Mapa" & map, "InviSinEfecto"))
780           .ResuSinEfecto = val(Leer.GetValue("Mapa" & map, "ResuSinEfecto"))
790           .OcultarSinEfecto = val(Leer.GetValue("Mapa" & map, "OcultarSinEfecto"))
800           .InvocarSinEfecto = val(Leer.GetValue("Mapa" & map, "InvocarSinEfecto"))
              
810           .NoEncriptarMP = val(Leer.GetValue("Mapa" & map, "NoEncriptarMP"))

820           .RoboNpcsPermitido = val(Leer.GetValue("Mapa" & map, "RoboNpcsPermitido"))
830            .levelmin = val(GetVar(MAPFl & ".dat", "Mapa" & map, "ELVMIN", Asc("-")))
840           .levelmax = val(GetVar(MAPFl & ".dat", "Mapa" & map, "ELVMAX", Asc("-")))
850           If .levelmax < 48 Then .levelmax = 48
             
860           If val(Leer.GetValue("Mapa" & map, "Pk")) = 0 Then
870               .Pk = True
880           Else
890               .Pk = False
900           End If
              
910           .Terreno = Leer.GetValue("Mapa" & map, "Terreno")
920           .Zona = Leer.GetValue("Mapa" & map, "Zona")
930           .Restringir = Leer.GetValue("Mapa" & map, "Restringir")
940           .BackUp = val(Leer.GetValue("Mapa" & map, "BACKUP"))
950       End With
          
960       Set MapReader = Nothing
970       Set InfReader = Nothing
980       Set Leer = Nothing
          
990       Erase Buff
1000  Exit Sub

errh:
          'Call LogError("Error cargando mapa: " & map & " - Pos: " & X & "," & Y & "." & Err.description)

1010      Set MapReader = Nothing
1020      Set InfReader = Nothing
1030      Set Leer = Nothing
End Sub
Sub LoadSini()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim Temporal As Long
          
10        If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."
          
20        BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))
          
          'Misc
    #If SeguridadAlkon Then
          
30        Call Security.SetServerIp(GetVar(IniPath & "Server.ini", "INIT", "ServerIp"))
          
    #End If
          
          
40        Puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
50        LastSockListen = val(GetVar(IniPath & "Server.ini", "INIT", "LastSockListen"))
60        HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
70        AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
80        IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
          'Lee la version correcta del cliente
90        ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")
          
100       PuedeCrearPersonajes = val(GetVar(IniPath & "Server.ini", "INIT", "PuedeCrearPersonajes"))
110       ServerSoloGMs = val(GetVar(IniPath & "Server.ini", "init", "ServerSoloGMs"))
          'INTERVALO_AUTO_GP = val(GetVar(IniPath & "Server.ini", "INIT", "AutoGP"))
          
120       ArmaduraImperial1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial1"))
130       ArmaduraImperial2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial2"))
140       ArmaduraImperial3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial3"))
150       TunicaMagoImperial = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperial"))
160       TunicaMagoImperialEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperialEnanos"))
170       ArmaduraCaos1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos1"))
180       ArmaduraCaos2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos2"))
190       ArmaduraCaos3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos3"))
200       TunicaMagoCaos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaos"))
210       TunicaMagoCaosEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaosEnanos"))
          
220       VestimentaImperialHumano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaImperialHumano"))
230       VestimentaImperialEnano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaImperialEnano"))
240       TunicaConspicuaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaConspicuaHumano"))
250       TunicaConspicuaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaConspicuaEnano"))
260       ArmaduraNobilisimaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraNobilisimaHumano"))
270       ArmaduraNobilisimaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraNobilisimaEnano"))
280       ArmaduraGranSacerdote = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraGranSacerdote"))
          
290       VestimentaLegionHumano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaLegionHumano"))
300       VestimentaLegionEnano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaLegionEnano"))
310       TunicaLobregaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaLobregaHumano"))
320       TunicaLobregaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaLobregaEnano"))
330       TunicaEgregiaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaEgregiaHumano"))
340       TunicaEgregiaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaEgregiaEnano"))
350       SacerdoteDemoniaco = val(GetVar(IniPath & "Server.ini", "INIT", "SacerdoteDemoniaco"))
          
360       MAPA_PRETORIANO = val(GetVar(IniPath & "Server.ini", "INIT", "MapaPretoriano"))
          
370       EnTesting = val(GetVar(IniPath & "Server.ini", "INIT", "Testing"))
          
          'Intervalos
380       SanaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))
390       FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar
          
400       StaminaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar"))
410       FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar
          
420       SanaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))
430       FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar
          
440       StaminaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar"))
450       FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar
          
460       IntervaloSed = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed"))
470       FrmInterv.txtIntervaloSed.Text = IntervaloSed
          
480       IntervaloHambre = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre"))
490       FrmInterv.txtIntervaloHambre.Text = IntervaloHambre
          
500       IntervaloVeneno = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno"))
510       FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno
          
520       IntervaloParalizado = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
530       FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado
          
540       IntervaloInvisible = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))
550       FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible
          
560       IntervaloFrio = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
570       FrmInterv.txtIntervaloFrio.Text = IntervaloFrio
          
580       IntervaloWavFx = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX"))
590       FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx
          
600       IntervaloInvocacion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
610       FrmInterv.txtInvocacion.Text = IntervaloInvocacion
          
620       IntervaloParaConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion"))
630       FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion
          
          '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
          
640       IntervaloPuedeSerAtacado = 5000 ' Cargar desde balance.dat
650       IntervaloAtacable = 60000 ' Cargar desde balance.dat
660       IntervaloOwnedNpc = 18000 ' Cargar desde balance.dat
          
670       IntervaloUserPuedeCastear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo"))
680       FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear
          
690       frmMain.TIMER_AI.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI"))
700       FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval
          
710       frmMain.npcataca.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar"))
720       FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.Interval
          
730       IntervaloUserPuedeTrabajar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
740       FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar
          
750       IntervaloUserPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))
760       FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar
          
          'TODO : Agregar estos intervalos al form!!!
770       IntervaloMagiaGolpe = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloMagiaGolpe"))
780       IntervaloGolpeMagia = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeMagia"))
790       IntervaloGolpeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeUsar"))
          
800       MinutosWs = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS"))
          
810       MinutosGuardarUsuarios = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGuardarUsuarios"))
          
820       IntervaloCerrarConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))
830       IntervaloUserPuedeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsar"))
840       IntervaloFlechasCazadores = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFlechasCazadores"))
          
850       IntervaloOculto = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloOculto"))
          
          '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
            
860       recordusuarios = val(GetVar(IniPath & "Server.ini", "INIT", "RECORD"))
          
          'Max users
870       Temporal = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
880       If MaxUsers = 0 Then
890           MaxUsers = Temporal
900           ReDim UserList(1 To MaxUsers) As User
910       End If
          
          '&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
          'Se agregó en LoadBalance y en el Balance.dat
          'PorcentajeRecuperoMana = val(GetVar(IniPath & "Server.ini", "BALANCE", "PorcentajeRecuperoMana"))
          
          ''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
920       Call Statistics.Initialize
          
930       Ullathorpe.map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
940       Ullathorpe.X = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
950       Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")
          
960       Nix.map = GetVar(DatPath & "Ciudades.dat", "Nix", "Mapa")
970       Nix.X = GetVar(DatPath & "Ciudades.dat", "Nix", "X")
980       Nix.Y = GetVar(DatPath & "Ciudades.dat", "Nix", "Y")
          
990       Banderbill.map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
1000      Banderbill.X = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
1010      Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")
          
1020      Lindos.map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
1030      Lindos.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
1040      Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")
          
1050      Arghal.map = GetVar(DatPath & "Ciudades.dat", "Arghal", "Mapa")
1060      Arghal.X = GetVar(DatPath & "Ciudades.dat", "Arghal", "X")
1070      Arghal.Y = GetVar(DatPath & "Ciudades.dat", "Arghal", "Y")
          
1080      Arkhein.map = GetVar(DatPath & "Ciudades.dat", "Arkhein", "Mapa")
1090      Arkhein.X = GetVar(DatPath & "Ciudades.dat", "Arkhein", "X")
1100      Arkhein.Y = GetVar(DatPath & "Ciudades.dat", "Arkhein", "Y")
          
1110      Nemahuak.map = GetVar(DatPath & "Ciudades.dat", "Nemahuak", "Mapa")
1120      Nemahuak.X = GetVar(DatPath & "Ciudades.dat", "Nemahuak", "X")
1130      Nemahuak.Y = GetVar(DatPath & "Ciudades.dat", "Nemahuak", "Y")

          
1140      Ciudades(eCiudad.cUllathorpe) = Ullathorpe
1150      Ciudades(eCiudad.cNix) = Nix
1160      Ciudades(eCiudad.cBanderbill) = Banderbill
1170      Ciudades(eCiudad.cLindos) = Lindos
1180      Ciudades(eCiudad.cArghal) = Arghal

1190      Call MD5sCarga
          
1200      Set ConsultaPopular = New ConsultasPopulares
1210      Call ConsultaPopular.LoadData

#If SeguridadAlkon Then
1220      Encriptacion.StringValidacion = Encriptacion.ArmarStringValidacion
#End If
          
           Call loadAdministrativeUsers
          
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      'Escribe VAR en un archivo
      '***************************************************

10    writeprivateprofilestring Main, Var, Value, File
          
End Sub

Sub SaveUser(ByVal UserIndex As Integer, ByVal UserFile As String, Optional ByVal SaveTimeOnline As Boolean = True)
      '*************************************************
      'Author: Unknown
      'Last modified: 10/10/2010 (Pato)
      'Saves the Users RECORDs
      '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
      '11/19/2009: Pato - Save the EluSkills and ExpSkills
      '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
      '10/10/2010: Pato - Saco el WriteVar e implemento la clase clsIniManager
      '*************************************************

10    On Error GoTo Errhandler

      Dim Manager As clsIniManager
      Dim Existe As Boolean

20    With UserList(UserIndex)

          'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
          'clase=0 es el error, porq el enum empieza de 1!!
30        If .clase = 0 Or .Stats.ELV = 0 Then
40            Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & .Name)
50            Exit Sub
60        End If
          
70        Set Manager = New clsIniManager
          
80        If FileExist(UserFile) Then
90            Call Manager.Initialize(UserFile)
              
100           If FileExist(UserFile & ".bk") Then Call Kill(UserFile & ".bk")
110           Name UserFile As UserFile & ".bk"
              
120           Existe = True
130       End If
          
140       If .flags.Mimetizado = 1 Then
150           .Char.body = .CharMimetizado.body
160           .Char.Head = .CharMimetizado.Head
170           .Char.CascoAnim = .CharMimetizado.CascoAnim
180           .Char.ShieldAnim = .CharMimetizado.ShieldAnim
190           .Char.WeaponAnim = .CharMimetizado.WeaponAnim
200           .Counters.Mimetismo = 0
210           .flags.Mimetizado = 0
              ' Se fue el efecto del mimetismo, puede ser atacado por npcs
220           .flags.Ignorado = False
230       End If
          
          
          Call Manager.ChangeValue("INIT", "ACCOUNT", CStr(.Account))
          
          Dim i As Long
          
240       For i = 1 To MAX_GUILDS_DISOLVED
250           Call Manager.ChangeValue("DISOLVED", "DISOLVED" & i, CStr(.GuildDisolved(i)))
260       Next i

          
270       Call Manager.ChangeValue("FLAGS", "Muerto", CStr(.flags.Muerto))
280       Call Manager.ChangeValue("FLAGS", "PREMIUM", CStr(.flags.Premium))
290       Call Manager.ChangeValue("FLAGS", "Escondido", CStr(.flags.Escondido))
300       Call Manager.ChangeValue("FLAGS", "Hambre", CStr(.flags.Hambre))
310       Call Manager.ChangeValue("FLAGS", "Sed", CStr(.flags.Sed))
320       Call Manager.ChangeValue("FLAGS", "Desnudo", CStr(.flags.Desnudo))
330       Call Manager.ChangeValue("FLAGS", "Ban", CStr(.flags.Ban))
340       Call Manager.ChangeValue("FLAGS", "Navegando", CStr(.flags.Navegando))
350       Call Manager.ChangeValue("FLAGS", "Montando", CStr(.flags.Montando))
360       Call Manager.ChangeValue("FLAGS", "BonosHP", CStr(.flags.BonosHP))
370       Call Manager.ChangeValue("FLAGS", "Envenenado", CStr(.flags.Envenenado))
380       Call Manager.ChangeValue("FLAGS", "Paralizado", CStr(.flags.Paralizado))
390       Call Manager.ChangeValue("FLAGS", "ORO", CStr(.flags.Oro))
400       Call Manager.ChangeValue("FLAGS", "Plata", CStr(.flags.Plata))
410       Call Manager.ChangeValue("FLAGS", "Bronce", CStr(.flags.Bronce))
420       Call Manager.ChangeValue("FLAGS", "DiosTerrenal", CStr(.flags.DiosTerrenal))
          'Matrix
430       Call Manager.ChangeValue("FLAGS", "LastMap", CStr(.flags.lastMap))
          
440       Call Manager.ChangeValue("CONSEJO", "PERTENECE", IIf(.flags.Privilegios And PlayerType.RoyalCouncil, "1", "0"))
450       Call Manager.ChangeValue("CONSEJO", "PERTENECECAOS", IIf(.flags.Privilegios And PlayerType.ChaosCouncil, "1", "0"))
          
          
          Call Manager.ChangeValue("COUNTERS", "TimeTelep", CStr(.Counters.TimeTelep))
460       Call Manager.ChangeValue("COUNTERS", "Pena", CStr(.Counters.Pena))
470       Call Manager.ChangeValue("COUNTERS", "SkillsAsignados", CStr(.Counters.AsignedSkills))
          
480       Call Manager.ChangeValue("FACCIONES", "EjercitoReal", CStr(.Faccion.ArmadaReal))
490       Call Manager.ChangeValue("FACCIONES", "EjercitoCaos", CStr(.Faccion.FuerzasCaos))
500       Call Manager.ChangeValue("FACCIONES", "CiudMatados", CStr(.Faccion.CiudadanosMatados))
510       Call Manager.ChangeValue("FACCIONES", "CrimMatados", CStr(.Faccion.CriminalesMatados))
520       Call Manager.ChangeValue("FACCIONES", "rArCaos", CStr(.Faccion.RecibioArmaduraCaos))
530       Call Manager.ChangeValue("FACCIONES", "rArReal", CStr(.Faccion.RecibioArmaduraReal))
540       Call Manager.ChangeValue("FACCIONES", "rExCaos", CStr(.Faccion.RecibioExpInicialCaos))
550       Call Manager.ChangeValue("FACCIONES", "rExReal", CStr(.Faccion.RecibioExpInicialReal))
560       Call Manager.ChangeValue("FACCIONES", "recCaos", CStr(.Faccion.RecompensasCaos))
570       Call Manager.ChangeValue("FACCIONES", "recReal", CStr(.Faccion.RecompensasReal))
580       Call Manager.ChangeValue("FACCIONES", "Reenlistadas", CStr(.Faccion.Reenlistadas))
590       Call Manager.ChangeValue("FACCIONES", "NivelIngreso", CStr(.Faccion.NivelIngreso))
600       Call Manager.ChangeValue("FACCIONES", "FechaIngreso", .Faccion.FechaIngreso)
610       Call Manager.ChangeValue("FACCIONES", "MatadosIngreso", CStr(.Faccion.MatadosIngreso))
620       Call Manager.ChangeValue("FACCIONES", "NextRecompensa", CStr(.Faccion.NextRecompensa))
          
          
          Dim LoopC As Long
              
              
          '¿Fueron modificados los atributos del usuario?
740       If Not .flags.TomoPocion Then
750           For LoopC = 1 To UBound(.Stats.UserAtributos)
760               Call Manager.ChangeValue("ATRIBUTOS", "AT" & LoopC, CStr(.Stats.UserAtributos(LoopC)))
770           Next LoopC
780       Else
790           For LoopC = 1 To UBound(.Stats.UserAtributos)
                  '.Stats.UserAtributos(LoopC) = .Stats.UserAtributosBackUP(LoopC)
800               Call Manager.ChangeValue("ATRIBUTOS", "AT" & LoopC, CStr(.Stats.UserAtributosBackUP(LoopC)))
810           Next LoopC
820       End If
          
830       For LoopC = 1 To UBound(.Stats.UserSkills)
840           Call Manager.ChangeValue("SKILLS", "SK" & LoopC, CStr(.Stats.UserSkills(LoopC)))
850           Call Manager.ChangeValue("SKILLS", "ELUSK" & LoopC, CStr(.Stats.EluSkills(LoopC)))
860           Call Manager.ChangeValue("SKILLS", "EXPSK" & LoopC, CStr(.Stats.ExpSkills(LoopC)))
870       Next LoopC
          
          
880       Call Manager.ChangeValue("CONTACTO", "Email", .Email)
890       Call Manager.ChangeValue("INIT", "Genero", .Genero)
900       Call Manager.ChangeValue("INIT", "Raza", .raza)
910       Call Manager.ChangeValue("INIT", "Hogar", .Hogar)
920       Call Manager.ChangeValue("INIT", "Clase", .clase)
930       Call Manager.ChangeValue("INIT", "Desc", .desc)
          
940       Call Manager.ChangeValue("INIT", "Heading", CStr(.Char.Heading))
950       Call Manager.ChangeValue("INIT", "Head", CStr(.OrigChar.Head))
          
960       If .flags.Muerto = 0 Then
970           If .Char.body <> 0 Then
980               Call Manager.ChangeValue("INIT", "Body", CStr(.Char.body))
990           End If
1000      End If
          
1010      Call Manager.ChangeValue("INIT", "Arma", CStr(.Char.WeaponAnim))
1020      Call Manager.ChangeValue("INIT", "Escudo", CStr(.Char.ShieldAnim))
1030      Call Manager.ChangeValue("INIT", "Casco", CStr(.Char.CascoAnim))
          
#If ConUpTime Then
          
1040      If SaveTimeOnline Then
              Dim TempDate As Date
1050          TempDate = Now - .LogOnTime
1060          .LogOnTime = Now
1070          .UpTime = .UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
1080          .UpTime = .UpTime
1090          Call Manager.ChangeValue("INIT", "UpTime", .UpTime)
1100      End If
#End If
          
          'First time around?
1110      If Manager.GetValue("INIT", "LastIP1") = vbNullString Then
1120          Call Manager.ChangeValue("INIT", "LastIP1", .ip & " - " & Date & ":" & time)
          'Is it a different ip from last time?
1130      ElseIf .ip <> Left$(Manager.GetValue("INIT", "LastIP1"), InStr(1, Manager.GetValue("INIT", "LastIP1"), " ") - 1) Then
1140          For i = 5 To 2 Step -1
1150              Call Manager.ChangeValue("INIT", "LastIP" & i, Manager.GetValue("INIT", "LastIP" & CStr(i - 1)))
1160          Next i
1170          Call Manager.ChangeValue("INIT", "LastIP1", .ip & " - " & Date & ":" & time)
          'Same ip, just update the date
1180      Else
1190          Call Manager.ChangeValue("INIT", "LastIP1", .ip & " - " & Date & ":" & time)
1200      End If
          
          
          
1210      Call Manager.ChangeValue("INIT", "Position", .Pos.map & "-" & .Pos.X & "-" & .Pos.Y)
          
          
          
          Dim tmpStr As String
          
1220      For LoopC = 0 To MAX_LOGROS
1230          If LoopC = 0 Then
1240              tmpStr = tmpStr & .Logros(LoopC)
1250          Else
1260              tmpStr = tmpStr & "-" & .Logros(LoopC)
1270          End If
1280      Next LoopC
          
          
1290      Call Manager.ChangeValue("STATS", "LOGROS", CStr(tmpStr))
          
1300      Call Manager.ChangeValue("STATS", "GLD", CStr(.Stats.Gld))
1310      Call Manager.ChangeValue("STATS", "BANCO", CStr(.Stats.Banco))
          
1320      Call Manager.ChangeValue("STATS", "MaxHP", CStr(.Stats.MaxHp))
1330      Call Manager.ChangeValue("STATS", "MinHP", CStr(.Stats.MinHp))
          
1340      Call Manager.ChangeValue("STATS", "MaxSTA", CStr(.Stats.MaxSta))
1350      Call Manager.ChangeValue("STATS", "MinSTA", CStr(.Stats.MinSta))
          
1360      Call Manager.ChangeValue("STATS", "MaxMAN", CStr(.Stats.MaxMAN))
1370      Call Manager.ChangeValue("STATS", "MinMAN", CStr(.Stats.MinMAN))
          
1380      Call Manager.ChangeValue("STATS", "MaxHIT", CStr(.Stats.MaxHIT))
1390      Call Manager.ChangeValue("STATS", "MinHIT", CStr(.Stats.MinHIT))
          
1400      Call Manager.ChangeValue("STATS", "MaxAGU", CStr(.Stats.MaxAGU))
1410      Call Manager.ChangeValue("STATS", "MinAGU", CStr(.Stats.MinAGU))
          
1420      Call Manager.ChangeValue("STATS", "MaxHAM", CStr(.Stats.MaxHam))
1430      Call Manager.ChangeValue("STATS", "MinHAM", CStr(.Stats.MinHam))
          
1440      Call Manager.ChangeValue("STATS", "SkillPtsLibres", CStr(.Stats.SkillPts))
            
1450      Call Manager.ChangeValue("STATS", "EXP", CStr(.Stats.Exp))
1460      Call Manager.ChangeValue("STATS", "ELV", CStr(.Stats.ELV))
          
          
1470      Call Manager.ChangeValue("STATS", "ELU", CStr(.Stats.ELU))
1480      Call Manager.ChangeValue("MUERTES", "UserMuertes", CStr(.Stats.UsuariosMatados))
          'Call Manager.ChangeValue( "MUERTES", "CrimMuertes", CStr(.Stats.CriminalesMatados))
1490      Call Manager.ChangeValue("MUERTES", "NpcsMuertes", CStr(.Stats.NPCsMuertos))
          
1500      Call Manager.ChangeValue("RETOS", "RetosGanados", CStr(.Stats.RetosGanados))
1510      Call Manager.ChangeValue("RETOS", "RetosPerdidos", CStr(.Stats.RetosPerdidos))
1520      Call Manager.ChangeValue("RETOS", "OroGanado", CStr(.Stats.OroGanado))
1530      Call Manager.ChangeValue("RETOS", "OroPerdido", CStr(.Stats.OroPerdido))
1540      Call Manager.ChangeValue("TORNEOS", "TorneosGanados", CStr(.Stats.TorneosGanados))
          Call Manager.ChangeValue("STATS", "Canje", CStr(.Stats.Points))
          
          '[KEVIN]----------------------------------------------------------------------------
          '*******************************************************************************************
1550      Call Manager.ChangeValue("BancoInventory", "CantidadItems", val(.BancoInvent.NroItems))
          Dim loopd As Integer
1560      For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
1570          Call Manager.ChangeValue("BancoInventory", "Obj" & loopd, .BancoInvent.Object(loopd).ObjIndex & "-" & .BancoInvent.Object(loopd).Amount)
1580      Next loopd
          '*******************************************************************************************
          '[/KEVIN]-----------
            
          'Save Inv
1590      Call Manager.ChangeValue("Inventory", "CantidadItems", val(.Invent.NroItems))
          
1600      For LoopC = 1 To MAX_INVENTORY_SLOTS
1610          Call Manager.ChangeValue("Inventory", "Obj" & LoopC, .Invent.Object(LoopC).ObjIndex & "-" & .Invent.Object(LoopC).Amount & "-" & .Invent.Object(LoopC).Equipped)
1620      Next LoopC
          
1630      Call Manager.ChangeValue("Inventory", "WeaponEqpSlot", CStr(.Invent.WeaponEqpSlot))
1640      Call Manager.ChangeValue("Inventory", "ArmourEqpSlot", CStr(.Invent.ArmourEqpSlot))
1650      Call Manager.ChangeValue("Inventory", "CascoEqpSlot", CStr(.Invent.CascoEqpSlot))
1660      Call Manager.ChangeValue("Inventory", "EscudoEqpSlot", CStr(.Invent.EscudoEqpSlot))
1670      Call Manager.ChangeValue("Inventory", "BarcoSlot", CStr(.Invent.BarcoSlot))
1680      Call Manager.ChangeValue("Inventory", "MonturaSlot", CStr(.Invent.MonturaSlot))
1690      Call Manager.ChangeValue("Inventory", "AnilloNpcSlot", CStr(.Invent.AnilloNpcSlot))
1700      Call Manager.ChangeValue("Inventory", "MunicionSlot", CStr(.Invent.MunicionEqpSlot))
1710      Call Manager.ChangeValue("Inventory", "MochilaSlot", CStr(.Invent.MochilaEqpSlot))
          '/Nacho
          
1720      Call Manager.ChangeValue("Inventory", "AnilloSlot", CStr(.Invent.AnilloEqpSlot))
          
          
          'Reputacion
1730      Call Manager.ChangeValue("REP", "Asesino", CStr(.Reputacion.AsesinoRep))
1740      Call Manager.ChangeValue("REP", "Bandido", CStr(.Reputacion.BandidoRep))
1750      Call Manager.ChangeValue("REP", "Burguesia", CStr(.Reputacion.BurguesRep))
1760      Call Manager.ChangeValue("REP", "Ladrones", CStr(.Reputacion.LadronesRep))
1770      Call Manager.ChangeValue("REP", "Nobles", CStr(.Reputacion.NobleRep))
1780      Call Manager.ChangeValue("REP", "Plebe", CStr(.Reputacion.PlebeRep))
          
          Dim L As Long
1790      L = (-.Reputacion.AsesinoRep) + _
              (-.Reputacion.BandidoRep) + _
              .Reputacion.BurguesRep + _
              (-.Reputacion.LadronesRep) + _
              .Reputacion.NobleRep + _
              .Reputacion.PlebeRep
1800      L = L / 6
1810      Call Manager.ChangeValue("REP", "Promedio", CStr(L))
          
          Dim cad As String
          
1820      For LoopC = 1 To MAXUSERHECHIZOS
1830          cad = .Stats.UserHechizos(LoopC)
1840          Call Manager.ChangeValue("HECHIZOS", "H" & LoopC, cad)
1850      Next
          
1860      Call SaveQuestStats(UserIndex, Manager)
          
          Dim NroMascotas As Long
1870      NroMascotas = .NroMascotas
          
1880      For LoopC = 1 To MAXMASCOTAS
              ' Mascota valida?
1890          If .MascotasIndex(LoopC) > 0 Then
                  ' Nos aseguramos que la criatura no fue invocada
1900              If Npclist(.MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
1910                  cad = .MascotasType(LoopC)
1920              Else 'Si fue invocada no la guardamos
1930                  cad = "0"
1940                  NroMascotas = NroMascotas - 1
1950              End If
1960              Call Manager.ChangeValue("MASCOTAS", "MAS" & LoopC, cad)
1970          Else
1980              cad = .MascotasType(LoopC)
1990              Call Manager.ChangeValue("MASCOTAS", "MAS" & LoopC, cad)
2000          End If
          
2010      Next
          
2020      Call Manager.ChangeValue("MASCOTAS", "NroMascotas", CStr(NroMascotas))
          
          'Devuelve el head de muerto
2030      If .flags.Muerto = 1 Then
2040          .Char.Head = iCabezaMuerto
2050      End If
2060  End With

2070  Call Manager.DumpFile(UserFile)

2080  Set Manager = Nothing

2090  If Existe Then Call Kill(UserFile & ".bk")

2100  Exit Sub

Errhandler:
2110  Call LogError("Error en SaveUser")
2120  Set Manager = Nothing

End Sub

Function criminal(ByVal UserIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim L As Long
          
10        With UserList(UserIndex).Reputacion
20            L = (-.AsesinoRep) + _
                  (-.BandidoRep) + _
                  .BurguesRep + _
                  (-.LadronesRep) + _
                  .NobleRep + _
                  .PlebeRep
30            L = L / 6
40            criminal = (L < 0)
50        End With

End Function

Sub BackUPnPc(ByVal NpcIndex As Integer, ByVal hFile As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 10/09/2010
      '10/09/2010 - Pato: Optimice el BackUp de NPCs
      '***************************************************

          Dim LoopC As Integer
          
10        Print #hFile, "[NPC" & Npclist(NpcIndex).Numero & "]"
          
20        With Npclist(NpcIndex)
              'General
30            Print #hFile, "Name=" & .Name
40            Print #hFile, "Desc=" & .desc
50            Print #hFile, "Head=" & val(.Char.Head)
60            Print #hFile, "Body=" & val(.Char.body)
70            Print #hFile, "Heading=" & val(.Char.Heading)
80            Print #hFile, "Movement=" & val(.Movement)
90            Print #hFile, "Attackable=" & val(.Attackable)
100           Print #hFile, "Comercia=" & val(.Comercia)
110           Print #hFile, "TipoItems=" & val(.TipoItems)
120           Print #hFile, "Hostil=" & val(.Hostile)
130           Print #hFile, "GiveEXP=" & val(.GiveEXP)
140           Print #hFile, "GiveGLD=" & val(.GiveGLD)
150           Print #hFile, "InvReSpawn=" & val(.InvReSpawn)
160           Print #hFile, "NpcType=" & val(.NPCtype)
              
              'Stats
170           Print #hFile, "Alineacion=" & val(.Stats.Alineacion)
180           Print #hFile, "DEF=" & val(.Stats.def)
190           Print #hFile, "MaxHit=" & val(.Stats.MaxHIT)
200           Print #hFile, "MaxHp=" & val(.Stats.MaxHp)
210           Print #hFile, "MinHit=" & val(.Stats.MinHIT)
220           Print #hFile, "MinHp=" & val(.Stats.MinHp)
              
              'Flags
230           Print #hFile, "ReSpawn=" & val(.flags.Respawn)
240           Print #hFile, "BackUp=" & val(.flags.BackUp)
250           Print #hFile, "Domable=" & val(.flags.Domable)
              
              'Inventario
260           Print #hFile, "NroItems=" & val(.Invent.NroItems)
270           If .Invent.NroItems > 0 Then
280              For LoopC = 1 To .Invent.NroItems
290                   Print #hFile, "Obj" & LoopC & "=" & .Invent.Object(LoopC).ObjIndex & "-" & .Invent.Object(LoopC).Amount
300              Next LoopC
310           End If
              
320           Print #hFile, ""
330       End With

End Sub

Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          'Status
10        If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"
          
          Dim npcfile As String
          
          'If NpcNumber > 499 Then
          '    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
          'Else
20            npcfile = DatPath & "bkNPCs.dat"
          'End If
          
30        With Npclist(NpcIndex)
          
40            .Numero = NpcNumber
50            .Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
60            .desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
70            .Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
80            .flags.OldMovement = .Movement
90            .NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))
              
100           .flags.AguaValida = val(GetVar(npcfile, "NPC" & NpcNumber, "AguaValida"))
110           .flags.TierraInvalida = val(GetVar(npcfile, "NPC" & NpcNumber, "TierraInValida"))
120           .flags.Faccion = val(GetVar(npcfile, "NPC" & NpcNumber, "Faccion"))
130           .flags.AtacaDoble = val(GetVar(npcfile, "NPC" & NpcNumber, "AtacaDoble"))
              
140           .Char.body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
150           .Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
160           .Char.ShieldAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "EscudoAnim"))
170           .Char.WeaponAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "ArmaAnim"))
180           .Char.CascoAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "CascoAnim"))
190           .Char.Heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))
              
              
200           .Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
210           .Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
220           .Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
230           .GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP")) * Expc
              
              
240           .GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))
              
250           .InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))
              
260           .Stats.MaxHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
270           .Stats.MinHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
280           .Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
290           .Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
300           .Stats.def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
310           .Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
              
              
              
              Dim LoopC As Integer
              Dim ln As String
320           .Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
330           If .Invent.NroItems > 0 Then
340               For LoopC = 1 To MAX_INVENTORY_SLOTS
350                   ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
360                   .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
370                   .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
                     
380               Next LoopC
390           Else
400               For LoopC = 1 To MAX_INVENTORY_SLOTS
410                   .Invent.Object(LoopC).ObjIndex = 0
420                   .Invent.Object(LoopC).Amount = 0
430               Next LoopC
440           End If
              
450           For LoopC = 1 To MAX_NPC_DROPS
460               ln = GetVar(npcfile, "NPC" & NpcNumber, "Drop" & LoopC)
470               .Drop(LoopC).ObjIndex = val(ReadField(1, ln, 45))
480               .Drop(LoopC).Amount = val(ReadField(2, ln, 45))
490               .Drop(LoopC).Probability = val(ReadField(3, ln, 45))
500           Next LoopC
              
510           .flags.NPCActive = True
520           .flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
530           .flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
540           .flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
550           .flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))
              
              'Tipo de items con los que comercia
560           .TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))
570       End With

End Sub


Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal motivo As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
20        Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "Reason", motivo)
          
          'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
          Dim mifile As Integer
30        mifile = FreeFile
40        Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
50        Print #mifile, UserList(BannedIndex).Name
60        Close #mifile

End Sub


Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal motivo As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
20        Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)
          
          'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
          Dim mifile As Integer
30        mifile = FreeFile
40        Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
50        Print #mifile, BannedName
60        Close #mifile

End Sub


Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal motivo As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
20        Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)
          
          
          'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
          Dim mifile As Integer
30        mifile = FreeFile
40        Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
50        Print #mifile, BannedName
60        Close #mifile

End Sub

Public Sub CargaApuestas()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
20        Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
30        Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

End Sub

Public Sub generateMatrix(ByVal Mapa As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim i As Integer
      Dim j As Integer
      Dim X As Integer
      Dim Y As Integer

10    ReDim distanceToCities(1 To NumMaps) As HomeDistance

20    For j = 1 To NUMCIUDADES
30        For i = 1 To NumMaps
40            distanceToCities(i).distanceToCity(j) = -1
50        Next i
60    Next j

70    For j = 1 To NUMCIUDADES
80        For i = 1 To 4
90            Select Case i
                  Case eHeading.NORTH
100                   Call setDistance(getLimit(Ciudades(j).map, eHeading.NORTH), j, i, 0, 1)
110               Case eHeading.EAST
120                   Call setDistance(getLimit(Ciudades(j).map, eHeading.EAST), j, i, 1, 0)
130               Case eHeading.SOUTH
140                   Call setDistance(getLimit(Ciudades(j).map, eHeading.SOUTH), j, i, 0, 1)
150               Case eHeading.WEST
160                   Call setDistance(getLimit(Ciudades(j).map, eHeading.WEST), j, i, -1, 0)
170           End Select
180       Next i
190   Next j

End Sub

Public Sub setDistance(ByVal Mapa As Integer, ByVal city As Byte, ByVal side As Integer, Optional ByVal X As Integer = 0, Optional ByVal Y As Integer = 0)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim i As Integer
      Dim lim As Integer

10    If Mapa <= 0 Or Mapa > NumMaps Then Exit Sub

20    If distanceToCities(Mapa).distanceToCity(city) >= 0 Then Exit Sub

30    If Mapa = Ciudades(city).map Then
40        distanceToCities(Mapa).distanceToCity(city) = 0
50    Else
60        distanceToCities(Mapa).distanceToCity(city) = Abs(X) + Abs(Y)
70    End If

80    For i = 1 To 4
90        lim = getLimit(Mapa, i)
100       If lim > 0 Then
110           Select Case i
                  Case eHeading.NORTH
120                   Call setDistance(lim, city, i, X, Y + 1)
130               Case eHeading.EAST
140                   Call setDistance(lim, city, i, X + 1, Y)
150               Case eHeading.SOUTH
160                   Call setDistance(lim, city, i, X, Y - 1)
170               Case eHeading.WEST
180                   Call setDistance(lim, city, i, X - 1, Y)
190           End Select
200       End If
210   Next i
End Sub

Public Function getLimit(ByVal Mapa As Integer, ByVal side As Byte) As Integer
      '***************************************************
      'Author: Budi
      'Last Modification: 31/01/2010
      'Retrieves the limit in the given side in the given map.
      'TODO: This should be set in the .inf map file.
      '***************************************************
      Dim i, X, Y As Integer

10    If Mapa <= 0 Then Exit Function

20    For X = 15 To 87
30        For Y = 0 To 3
40            Select Case side
                  Case eHeading.NORTH
50                    getLimit = MapData(Mapa, X, 7 + Y).TileExit.map
60                Case eHeading.EAST
70                    getLimit = MapData(Mapa, 92 - Y, X).TileExit.map
80                Case eHeading.SOUTH
90                    getLimit = MapData(Mapa, X, 94 - Y).TileExit.map
100               Case eHeading.WEST
110                   getLimit = MapData(Mapa, 9 + Y, X).TileExit.map
120           End Select
130           If getLimit > 0 Then Exit Function
140       Next Y
150   Next X
End Function


Public Sub LoadArmadurasFaccion()
      '***************************************************
      'Author: ZaMa
      'Last Modification: 15/04/2010
      '
      '***************************************************
          Dim ClassIndex As Long
          Dim RaceIndex As Long
          
          Dim ArmaduraIndex As Integer
          
          
10        For ClassIndex = 1 To NUMCLASES
          
              ' Defensa minima para armadas altos
20            ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinArmyAlto"))
              
30            ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
40            ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
50            ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
              
              ' Defensa minima para armadas bajos
60            ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinArmyBajo"))
              
70            ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
80            ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
              
              ' Defensa minima para caos altos
90            ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinCaosAlto"))
              
100           ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
110           ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
120           ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
              
              ' Defensa minima para caos bajos
130           ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinCaosBajo"))
              
140           ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
150           ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
          
          
              ' Defensa media para armadas altos
160           ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedArmyAlto"))
              
170           ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
180           ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
190           ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
              
              ' Defensa media para armadas bajos
200           ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedArmyBajo"))
              
210           ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
220           ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
              
              ' Defensa media para caos altos
230           ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedCaosAlto"))
              
240           ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
250           ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
260           ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
              
              ' Defensa media para caos bajos
270           ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedCaosBajo"))
              
280           ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
290           ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
          
          
              ' Defensa alta para armadas altos
300           ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaArmyAlto"))
              
310           ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
320           ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
330           ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
              
              ' Defensa alta para armadas bajos
340           ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaArmyBajo"))
              
350           ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
360           ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
              
              ' Defensa alta para caos altos
370           ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaCaosAlto"))
              
380           ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
390           ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
400           ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
              
              ' Defensa alta para caos bajos
410           ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaCaosBajo"))
              
420           ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
430           ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
          
440       Next ClassIndex
          
End Sub

