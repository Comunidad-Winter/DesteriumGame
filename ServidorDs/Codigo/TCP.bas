Attribute VB_Name = "TCP"
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

#If UsarQueSocket = 0 Then
' General constants used with most of the controls
Public Const INVALID_HANDLE As Integer = -1
Public Const CONTROL_ERRIGNORE As Integer = 0
Public Const CONTROL_ERRDISPLAY As Integer = 1


' SocietWrench Control Actions
Public Const SOCKET_OPEN As Integer = 1
Public Const SOCKET_CONNECT As Integer = 2
Public Const SOCKET_LISTEN As Integer = 3
Public Const SOCKET_ACCEPT As Integer = 4
Public Const SOCKET_CANCEL As Integer = 5
Public Const SOCKET_FLUSH As Integer = 6
Public Const SOCKET_CLOSE As Integer = 7
Public Const SOCKET_DISCONNECT As Integer = 7
Public Const SOCKET_ABORT As Integer = 8

' SocketWrench Control States
Public Const SOCKET_NONE As Integer = 0
Public Const SOCKET_IDLE As Integer = 1
Public Const SOCKET_LISTENING As Integer = 2
Public Const SOCKET_CONNECTING As Integer = 3
Public Const SOCKET_ACCEPTING As Integer = 4
Public Const SOCKET_RECEIVING As Integer = 5
Public Const SOCKET_SENDING As Integer = 6
Public Const SOCKET_CLOSING As Integer = 7

' Societ Address Families
Public Const AF_UNSPEC As Integer = 0
Public Const AF_UNIX As Integer = 1
Public Const AF_INET As Integer = 2

' Societ Types
Public Const SOCK_STREAM As Integer = 1
Public Const SOCK_DGRAM As Integer = 2
Public Const SOCK_RAW As Integer = 3
Public Const SOCK_RDM As Integer = 4
Public Const SOCK_SEQPACKET As Integer = 5

' Protocol Types
Public Const IPPROTO_IP As Integer = 0
Public Const IPPROTO_ICMP As Integer = 1
Public Const IPPROTO_GGP As Integer = 2
Public Const IPPROTO_TCP As Integer = 6
Public Const IPPROTO_PUP As Integer = 12
Public Const IPPROTO_UDP As Integer = 17
Public Const IPPROTO_IDP As Integer = 22
Public Const IPPROTO_ND As Integer = 77
Public Const IPPROTO_RAW As Integer = 255
Public Const IPPROTO_MAX As Integer = 256


' Network Addpesses
Public Const INADDR_ANY As String = "0.0.0.0"
Public Const INADDR_LOOPBACK As String = "127.0.0.1"
Public Const INADDR_NONE As String = "255.055.255.255"

' Shutdown Values
Public Const SOCKET_READ As Integer = 0
Public Const SOCKET_WRITE As Integer = 1
Public Const SOCKET_READWRITE As Integer = 2

' SocketWrench Error Pesponse
Public Const SOCKET_ERRIGNORE As Integer = 0
Public Const SOCKET_ERRDISPLAY As Integer = 1

' SocketWrench Error Codes
Public Const WSABASEERR As Integer = 24000
Public Const WSAEINTR As Integer = 24004
Public Const WSAEBADF As Integer = 24009
Public Const WSAEACCES As Integer = 24013
Public Const WSAEFAULT As Integer = 24014
Public Const WSAEINVAL As Integer = 24022
Public Const WSAEMFILE As Integer = 24024
Public Const WSAEWOULDBLOCK As Integer = 24035
Public Const WSAEINPROGRESS As Integer = 24036
Public Const WSAEALREADY As Integer = 24037
Public Const WSAENOTSOCK As Integer = 24038
Public Const WSAEDESTADDRREQ As Integer = 24039
Public Const WSAEMSGSIZE As Integer = 24040
Public Const WSAEPROTOTYPE As Integer = 24041
Public Const WSAENOPROTOOPT As Integer = 24042
Public Const WSAEPROTONOSUPPORT As Integer = 24043
Public Const WSAESOCKTNOSUPPORT As Integer = 24044
Public Const WSAEOPNOTSUPP As Integer = 24045
Public Const WSAEPFNOSUPPORT As Integer = 24046
Public Const WSAEAFNOSUPPORT As Integer = 24047
Public Const WSAEADDRINUSE As Integer = 24048
Public Const WSAEADDRNOTAVAIL As Integer = 24049
Public Const WSAENETDOWN As Integer = 24050
Public Const WSAENETUNREACH As Integer = 24051
Public Const WSAENETRESET As Integer = 24052
Public Const WSAECONNABORTED As Integer = 24053
Public Const WSAECONNRESET As Integer = 24054
Public Const WSAENOBUFS As Integer = 24055
Public Const WSAEISCONN As Integer = 24056
Public Const WSAENOTCONN As Integer = 24057
Public Const WSAESHUTDOWN As Integer = 24058
Public Const WSAETOOMANYREFS As Integer = 24059
Public Const WSAETIMEDOUT As Integer = 24060
Public Const WSAECONNREFUSED As Integer = 24061
Public Const WSAELOOP As Integer = 24062
Public Const WSAENAMETOOLONG As Integer = 24063
Public Const WSAEHOSTDOWN As Integer = 24064
Public Const WSAEHOSTUNREACH As Integer = 24065
Public Const WSAENOTEMPTY As Integer = 24066
Public Const WSAEPROCLIM As Integer = 24067
Public Const WSAEUSERS As Integer = 24068
Public Const WSAEDQUOT As Integer = 24069
Public Const WSAESTALE As Integer = 24070
Public Const WSAEREMOTE As Integer = 24071
Public Const WSASYSNOTREADY As Integer = 24091
Public Const WSAVERNOTSUPPORTED As Integer = 24092
Public Const WSANOTINITIALISED As Integer = 24093
Public Const WSAHOST_NOT_FOUND As Integer = 25001
Public Const WSATRY_AGAIN As Integer = 25002
Public Const WSANO_RECOVERY As Integer = 25003
Public Const WSANO_DATA As Integer = 25004
Public Const WSANO_ADDRESS As Integer = 2500
#End If

Sub DarCuerpoYCabeza(ByVal userindex As Integer)
      '*************************************************
      'Author: Nacho (Integer)
      'Last modified: 14/03/2007
      'Elije una cabeza para el usuario y le da un body
      '*************************************************
      Dim NewBody As Integer
      Dim NewHead As Integer
      Dim UserRaza As Byte
      Dim UserGenero As Byte
10    UserGenero = UserList(userindex).Genero
20    UserRaza = UserList(userindex).raza
30    Select Case UserGenero
         Case eGenero.Hombre
40            Select Case UserRaza
                  Case eRaza.Humano
50                    NewHead = RandomNumber(1, 25)
60                    NewBody = 1
70                Case eRaza.Elfo
80                    NewHead = RandomNumber(102, 111)
90                    NewBody = 2
100               Case eRaza.Drow
110                   NewHead = RandomNumber(201, 205)
120                   NewBody = 3
130               Case eRaza.Enano
140                   NewHead = RandomNumber(301, 305)
150                   NewBody = 52
160               Case eRaza.Gnomo
170                   NewHead = RandomNumber(401, 405)
180                   NewBody = 52
190           End Select
200      Case eGenero.Mujer
210           Select Case UserRaza
                  Case eRaza.Humano
220                   NewHead = RandomNumber(71, 75)
230                   NewBody = 44
240               Case eRaza.Elfo
250                   NewHead = RandomNumber(170, 174)
260                   NewBody = 44
270               Case eRaza.Drow
280                   NewHead = RandomNumber(270, 276)
290                   NewBody = 44
300               Case eRaza.Gnomo
310                   NewHead = RandomNumber(471, 475)
320                   NewBody = 138
330               Case eRaza.Enano
340                   NewHead = RandomNumber(370, 371)
350                   NewBody = 138
360           End Select
370   End Select
380   UserList(userindex).Char.Head = NewHead
390   UserList(userindex).Char.body = NewBody
End Sub
Function AsciiValidos(ByVal cad As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim Car As Byte
      Dim i As Integer

10    cad = LCase$(cad)

20    For i = 1 To Len(cad)
30        Car = Asc(mid$(cad, i, 1))
          
40        If (Car < 97 Or Car > 122) And (Car <> 255) And (Car <> 32) Then
50            AsciiValidos = False
60            Exit Function
70        End If
          
80    Next i

90    AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim Car As Byte
      Dim i As Integer

10    cad = LCase$(cad)

20    For i = 1 To Len(cad)
30        Car = Asc(mid$(cad, i, 1))
          
40        If (Car < 48 Or Car > 57) Then
50            Numeric = False
60            Exit Function
70        End If
          
80    Next i

90    Numeric = True

End Function


Function NombrePermitido(ByVal Nombre As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim i As Integer

10    For i = 1 To UBound(ForbidenNames)
20        If InStr(Nombre, ForbidenNames(i)) Then
30                NombrePermitido = False
40                Exit Function
50        End If
60    Next i

70    NombrePermitido = True

End Function
Function ValidateSkills(ByVal userindex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim LoopC As Integer

10    For LoopC = 1 To NUMSKILLS
20        If UserList(userindex).Stats.UserSkills(LoopC) < 0 Then
30            Exit Function
40            If UserList(userindex).Stats.UserSkills(LoopC) > 100 Then UserList(userindex).Stats.UserSkills(LoopC) = 100
50        End If
60    Next LoopC

70    ValidateSkills = True
          
End Function
Public Sub KillCharINFO(ByVal User As String)
10    On Error Resume Next
       
      Dim c As String
      Dim d As String
      Dim f As String
      Dim g As String
      Dim h As Byte
      Dim i As String
      Dim j As String
       
       
20    c = GetVar(App.Path & "\CHARFILE\" & User & ".chr", "GUILD", "GUILDINDEX")
30    d = GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & c, "founder")
40    f = GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & c, "GuildName")
50    g = GetVar(App.Path & "\guilds\" & f & "-members.mem", "INIT", "NroMembers")
60    j = GetVar(App.Path & "\guilds\" & f & "-members.mem", "Members", "Member" & g)
       
       
       
70        If c = "" Then
80            Kill (App.Path & "\CHARFILE\" & User & ".chr")
90        Else
         
100           If d <> User Then
110               guilds(c).ExpulsarMiembro (User)
120           Else
130               For h = 1 To g
                 
140       i = GetVar(App.Path & "\guilds\" & f & "-members.mem", "Members", "Member" & h)
       
150                   If i = User Then Call WriteVar(App.Path & "\guilds\" & f & "-members.mem", "Members", "Member" & h, j): Call WriteVar(App.Path & "\guilds\" & f & "-members.mem", "INIT", "NroMembers", g - 1)
       
160       Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & c, "EleccionesAbiertas", "1")
170       Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & c, "EleccionesFinalizan", DateAdd("d", 1, Now))
180   Call WriteVar(App.Path & "\guilds\" & f & "-votaciones.vot", "INIT", "NumVotos", "0")
190               Next h
                 
200           End If
210       Kill (App.Path & "\CHARFILE\" & User & ".chr")
220      End If
       
End Sub
       
           
Public Function GenerateRandomKey(ByVal User As String) As Integer
10    On Error Resume Next
       
20            GenerateRandomKey = RandomNumber(500, 1600) + RandomNumber(5000, 16000)
30            Call WriteVar(App.Path & "\CHARFILE\" & User & ".chr", "INIT", "Password", GenerateRandomKey)
       
End Function


Sub ConnectNewUser(ByVal userindex As Integer, ByRef Name As String, ByVal UserClase As eClass, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero)
      '*************************************************
      'Author: Unknown
      'Last modified: 3/12/2009
      'Conecta un nuevo Usuario
      '23/01/2007 Pablo (ToxicWaste) - Agregué ResetFaccion al crear usuario
      '24/01/2007 Pablo (ToxicWaste) - Agregué el nuevo mana inicial de los magos.
      '12/02/2007 Pablo (ToxicWaste) - Puse + 1 de const al Elfo normal.
      '20/04/2007 Pablo (ToxicWaste) - Puse -1 de fuerza al Elfo.
      '09/01/2008 Pablo (ToxicWaste) - Ahora los modificadores de Raza se controlan desde Balance.dat
      '11/19/2009: Pato - Modifico la maná inicial del bandido.
      '11/19/2009: Pato - Asigno los valores iniciales de ExpSkills y EluSkills.
      '03/12/2009: Budi - Optimización del código.
      '*************************************************
      Dim i As Long

10    With UserList(userindex)

20        If Not AsciiValidos(Name) Or LenB(Name) = 0 Then
30            Call WriteErrorMsg(userindex, "Nombre inválido.")
40            Exit Sub
50        End If
          
60        If UserList(userindex).flags.UserLogged Then
70            Call LogCheating("El usuario " & UserList(userindex).Name & " ha intentado crear a " & Name & " desde la IP " & UserList(userindex).ip)
              
              'Kick player ( and leave character inside :D )!
80            Call CloseSocketSL(userindex)
90            Call Cerrar_Usuario(userindex)
              
100           Exit Sub
110       End If
          
          '¿Existe el personaje?
120       If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = True Then
130           Call WriteErrorMsg(userindex, "Ya existe el personaje.")
140           Exit Sub
150       End If
          
          'Tiró los dados antes de llegar acá??
160       If .Stats.UserAtributos(eAtributos.Fuerza) = 0 Then
170           Call WriteErrorMsg(userindex, "Debe tirar los dados antes de poder crear un personaje.")
180           Exit Sub
190       End If
          

200       .flags.Muerto = 0
210       .flags.Oro = 0
220       .flags.Plata = 0
230       .flags.Bronce = 0
240       .flags.DiosTerrenal = 0
250       .flags.Escondido = 0
260       .flags.ModoCombate = 0
270       .Reputacion.AsesinoRep = 0
280       .Reputacion.BandidoRep = 0
290       .Reputacion.BurguesRep = 0
300       .Reputacion.LadronesRep = 0
310       .Reputacion.PlebeRep = 30
          
320       .Reputacion.Promedio = 30 / 6
          
          
330       .Name = Name
340       .clase = UserClase
350       .raza = UserRaza
360       .Genero = UserSexo
380       .Hogar = cUllathorpe
          
400       .flags.Premium = 0
          
410       '.CPU_ID = CPU_ID
          
420       Call ResetMercado(userindex)
          
          '[Pablo (Toxic Waste) 9/01/08]
430       .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + ModRaza(UserRaza).Fuerza
440       .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + ModRaza(UserRaza).Agilidad
450       .Stats.UserAtributos(eAtributos.Inteligencia) = .Stats.UserAtributos(eAtributos.Inteligencia) + ModRaza(UserRaza).Inteligencia
460       .Stats.UserAtributos(eAtributos.Carisma) = .Stats.UserAtributos(eAtributos.Carisma) + ModRaza(UserRaza).Carisma
470       .Stats.UserAtributos(eAtributos.Constitucion) = .Stats.UserAtributos(eAtributos.Constitucion) + ModRaza(UserRaza).Constitucion
          '[/Pablo (Toxic Waste)]
          
480       For i = 1 To NUMSKILLS
490           .Stats.UserSkills(i) = 0
500           Call CheckEluSkill(userindex, i, True)
510       Next i
          
520       .Stats.SkillPts = 0
          
530       .Char.Heading = eHeading.SOUTH
          
540   Call DarCuerpoYCabeza(userindex)
          
550       .OrigChar = .Char
          
          Dim MiInt As Long
560       MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Constitucion) \ 3)
          
570       .Stats.MaxHp = 15 + MiInt
580       .Stats.MinHp = 15 + MiInt

590       MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Agilidad) \ 6)
600       If MiInt = 1 Then MiInt = 2
          
610       .Stats.MaxSta = 20 * MiInt
620       .Stats.MinSta = 20 * MiInt
          
          
630       .Stats.MaxAGU = 100
640       .Stats.MinAGU = 100
          
650       .Stats.MaxHam = 100
660       .Stats.MinHam = 100
          
          
670       .Stats.RetosGanados = 0
680       .Stats.RetosPerdidos = 0
690       .Stats.OroGanado = 0
700       .Stats.OroPerdido = 0
710       .Stats.OldHp = 0
          
720       .Stats.TorneosGanados = 0
          .Stats.Points = 0
          
          '<-----------------MANA----------------------->
730       If UserClase = eClass.Mage Then 'Cambio en mana inicial (ToxicWaste)
740           MiInt = RandomNumber(100, 106)
750           .Stats.MaxMAN = MiInt
760           .Stats.MinMAN = MiInt
770       ElseIf UserClase = eClass.Cleric Or UserClase = eClass.Druid _
              Or UserClase = eClass.Bard Or UserClase = eClass.Assasin Then
780               .Stats.MaxMAN = 50
790               .Stats.MinMAN = 50
         ' ElseIf UserClase = eClass.Bandit Then 'Mana Inicial del Bandido (ToxicWaste)
                '  .Stats.MaxMAN = 50
                 ' .Stats.MinMAN = 50
800       Else
810           .Stats.MaxMAN = 0
820           .Stats.MinMAN = 0
830       End If
          
840       If UserClase = eClass.Mage Or UserClase = eClass.Cleric Or _
             UserClase = eClass.Druid Or UserClase = eClass.Bard Or _
             UserClase = eClass.Assasin Then
850               .Stats.UserHechizos(1) = 2
              
                  'If UserClase = eClass.Druid Then .Stats.UserHechizos(2) = 46
860       End If
          
870       .Stats.MaxHIT = 2
880       .Stats.MinHIT = 1
          
890       .Stats.Gld = 5000
          
900       .Stats.Exp = 0
910       .Stats.ELU = 300
920       .Stats.ELV = 1
930       .flags.BonosHP = 0
          
          '???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
          Dim Slot As Byte
          Dim IsPaladin As Boolean
          
940       IsPaladin = UserClase = eClass.Paladin
          
          'Pociones Rojas (Newbie)
950       Slot = 1
960       .Invent.Object(Slot).ObjIndex = 461
970       .Invent.Object(Slot).Amount = 50
          
          'Pociones azules (Newbie)
980       If .Stats.MaxMAN > 0 Or IsPaladin Then
990           Slot = Slot + 1
1000          .Invent.Object(Slot).ObjIndex = 462
1010          .Invent.Object(Slot).Amount = 50
          
1020      End If
          
          ' Ropa (Newbie)
1030      Slot = Slot + 1
1040      Select Case UserRaza
              Case eRaza.Humano
1050              .Invent.Object(Slot).ObjIndex = 463
1060          Case eRaza.Elfo
1070              .Invent.Object(Slot).ObjIndex = 464
1080          Case eRaza.Drow
1090              .Invent.Object(Slot).ObjIndex = 465
1100          Case eRaza.Enano
1110              .Invent.Object(Slot).ObjIndex = 466
1120          Case eRaza.Gnomo
1130              .Invent.Object(Slot).ObjIndex = 466
1140      End Select
          
          ' Equipo ropa
1150      .Invent.Object(Slot).Amount = 1
1160      .Invent.Object(Slot).Equipped = 1
          
1170      .Invent.ArmourEqpSlot = Slot
1180      .Invent.ArmourEqpObjIndex = .Invent.Object(Slot).ObjIndex

          'Arma (Newbie)
1190      Slot = Slot + 1
1200      Select Case UserClase
              Case eClass.Hunter
                  ' Arco (Newbie)
1210              .Invent.Object(Slot).ObjIndex = 460
1220          Case eClass.Worker
                  ' Herramienta (Newbie)
1230              .Invent.Object(Slot).ObjIndex = 460
1240          Case Else
                  ' Daga (Newbie)
1250              .Invent.Object(Slot).ObjIndex = 460
1260      End Select
          
          ' Equipo arma
1270      .Invent.Object(Slot).Amount = 1
1280      .Invent.Object(Slot).Equipped = 1
          
1290      .Invent.WeaponEqpObjIndex = .Invent.Object(Slot).ObjIndex
1300      .Invent.WeaponEqpSlot = Slot
          
1310      .Char.WeaponAnim = GetWeaponAnim(userindex, .Invent.WeaponEqpObjIndex)

          ' Manzanas (Newbie)
1320      Slot = Slot + 1
1330      .Invent.Object(Slot).ObjIndex = 467
1340      .Invent.Object(Slot).Amount = 100
          
          ' Jugos (Nwbie)
1350      Slot = Slot + 1
1360      .Invent.Object(Slot).ObjIndex = 468
1370      .Invent.Object(Slot).Amount = 100
          
          ' Sin casco y escudo
1380      .Char.ShieldAnim = NingunEscudo
1390      .Char.CascoAnim = NingunCasco
          
          ' Total Items
1400      .Invent.NroItems = Slot
          
    #If ConUpTime Then
1410          .LogOnTime = Now
1420          .UpTime = 0
    #End If

1430   If .flags.ModoCombate = False Then
1440          WriteConsoleMsg userindex, "No estás en modo combate, para activarlo presiona la tecla C.", FontTypeNames.FONTTYPE_INFO
1450  End If
1460  End With



      'Valores Default de facciones al Activar nuevo usuario
1470  Call ResetFaccionCaos(userindex)
1480  Call ResetFaccionReal(userindex)

1490  'Call WriteVar(CharPath & UCase$(Name) & ".chr", "INIT", "Password", Password) 'grabamos el password aqui afuera, para no mantenerlo cargado en memoria

1500  'Call WriteVar(CharPath & UCase$(Name) & ".chr", "INIT", "Pin", Pin)

1510  Call SaveUser(userindex, CharPath & UCase$(Name) & ".chr")
      Dim SABPO As String
1520  SABPO = CharPath & UCase$(Name) & ".chr"

      'Open User
1530  Call ConnectUser(userindex, Name)
        
End Sub

#If UsarQueSocket = 1 Or UsarQueSocket = 2 Then

Sub CloseSocket(ByVal userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler
          
20        Call aDos.RestarConexion(UserList(userindex).ip)
          
30        If userindex = LastUser Then
40            Do Until UserList(LastUser).flags.UserLogged
50                LastUser = LastUser - 1
60                If LastUser < 1 Then Exit Do
70            Loop
80        End If
          
90        If UserList(userindex).flags.Montando = 0 Then
          
100       End If

         ' @@ Si un GM me tiene fichado entonces..
       
110       If UserList(userindex).flags.ElPedidorSeguimiento > 0 Then
       
120           Call WriteConsoleMsg(UserList(userindex).flags.ElPedidorSeguimiento, "El usuario que estás siguiendo se desconectó, re putin", FontTypeNames.fonttype_dios)
             
130           Call Protocol.WriteShowPanelSeguimiento(UserList(userindex).flags.ElPedidorSeguimiento, 2)
       
             
140       End If

          ' Si tenemos una transformación la sacamos
            If UserList(userindex).Invent.AnilloNpcObjIndex > 0 Then
                Call Desequipar(userindex, UserList(userindex).Invent.AnilloNpcSlot)
            End If
    
           AbandonateEvent userindex, , True
              
           If UserList(userindex).flags.SlotReto > 0 Then
               Call mRetos.UserdieFight(userindex, 0, True)
           End If
              
           If UserList(userindex).flags.InCVC Then
               CloseUserCvc userindex
           End If

          Call SecurityIp.IpRestarConexion(GetLongIp(UserList(userindex).ip))

150       If UserList(userindex).ConnID <> -1 Then
160           Call CloseSocketSL(userindex)
170       End If
          
          'Es el mismo user al que está revisando el centinela??
          'IMPORTANTE!!! hacerlo antes de resetear así todavía sabemos el nombre del user
          ' y lo podemos loguear
180       If Centinela.RevisandoUserIndex = userindex Then _
              Call modCentinela.CentinelaUserLogout
          
          
          'mato los comercios seguros
190       If UserList(userindex).ComUsu.DestUsu > 0 Then
200           If UserList(UserList(userindex).ComUsu.DestUsu).flags.UserLogged Then
210               If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.DestUsu = userindex Then
220                   Call WriteConsoleMsg(UserList(userindex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
230                   Call FinComerciarUsu(UserList(userindex).ComUsu.DestUsu)
240                   Call FlushBuffer(UserList(userindex).ComUsu.DestUsu)
250               End If
260           End If
270       End If
          
          'Empty buffer for reuse
280       Call UserList(userindex).incomingData.ReadASCIIStringFixed(UserList(userindex).incomingData.length)
          
290       If UserList(userindex).flags.UserLogged Then
300           If NumUsers > 0 Then NumUsers = NumUsers - 1
310           Call CloseUser(userindex)
320           Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
330       Else
340           Call ResetUserSlot(userindex)
350       End If
          
360       UserList(userindex).ConnID = -1
370       UserList(userindex).ConnIDValida = False
          
380   Exit Sub

Errhandler:
390       UserList(userindex).ConnID = -1
400       UserList(userindex).ConnIDValida = False
410       Call ResetUserSlot(userindex)

420       Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.Description & " - UserIndex = " & userindex)

End Sub

#ElseIf UsarQueSocket = 0 Then

Sub CloseSocket(ByVal userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler
    
    UserList(userindex).ConnID = -1


    If userindex = LastUser And LastUser > 1 Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop
    End If

    If UserList(userindex).flags.UserLogged Then
            If NumUsers <> 0 Then NumUsers = NumUsers - 1
            Call CloseUser(userindex)
    End If

    frmMain.Socket2(userindex).Cleanup
    Unload frmMain.Socket2(userindex)
    Call ResetUserSlot(userindex)

Exit Sub

Errhandler:
    UserList(userindex).ConnID = -1
    Call ResetUserSlot(userindex)
End Sub


#ElseIf UsarQueSocket = 3 Then

Sub CloseSocket(ByVal userindex As Integer, Optional ByVal cerrarlo As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo Errhandler

Dim NURestados As Boolean
Dim CoNnEcTiOnId As Long


    NURestados = False
    CoNnEcTiOnId = UserList(userindex).ConnID
    
    'call logindex(UserIndex, "******> Sub CloseSocket. ConnId: " & CoNnEcTiOnId & " Cerrarlo: " & Cerrarlo)
    
    
  
    UserList(userindex).ConnID = -1 'inabilitamos operaciones en socket
    
    If userindex = LastUser And LastUser > 1 Then
        Do
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop While UserList(LastUser).flags.UserLogged = True
    End If

    If UserList(userindex).flags.UserLogged Then
            If NumUsers <> 0 Then NumUsers = NumUsers - 1
            NURestados = True
            Call CloseUser(userindex)
    End If
    
    Call ResetUserSlot(userindex)
    
    'limpiada la userlist... reseteo el socket, si me lo piden
    'Me lo piden desde: cerrada intecional del servidor (casi todas
    'las llamadas a CloseSocket del codigo)
    'No me lo piden desde: disconnect remoto (el on_close del control
    'de alejo realiza la desconexion automaticamente). Esto puede pasar
    'por ejemplo, si el cliente cierra el AO.
    If cerrarlo Then Call frmMain.TCPServ.CerrarSocket(CoNnEcTiOnId)

Exit Sub

Errhandler:
    Call LogError("CLOSESOCKETERR: " & Err.Description & " UI:" & userindex)
    
    If Not NURestados Then
        If UserList(userindex).flags.UserLogged Then
            If NumUsers > 0 Then
                NumUsers = NumUsers - 1
            End If
            Call LogError("Cerre sin grabar a: " & UserList(userindex).Name)
        End If
    End If
    
    Call LogError("El usuario no guardado tenía connid " & CoNnEcTiOnId & ". Socket no liberado.")
    Call ResetUserSlot(userindex)
    
    

End Sub


#End If

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

#If UsarQueSocket = 1 Then

If UserList(userindex).ConnID <> -1 And UserList(userindex).ConnIDValida Then
    Call BorraSlotSock(UserList(userindex).ConnID)
    Call WSApiCloseSocket(UserList(userindex).ConnID)
    UserList(userindex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 0 Then

If UserList(userindex).ConnID <> -1 And UserList(userindex).ConnIDValida Then
    frmMain.Socket2(userindex).Cleanup
    Unload frmMain.Socket2(userindex)
    UserList(userindex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 2 Then

If UserList(userindex).ConnID <> -1 And UserList(userindex).ConnIDValida Then
    Call frmMain.Serv.CerrarSocket(UserList(userindex).ConnID)
    UserList(userindex).ConnIDValida = False
End If

#End If
End Sub

''
' Send an string to a Slot
'
' @param userIndex The index of the User
' @param Datos The string that will be send
' @remarks If UsarQueSocket is 3 it won`t use the clsByteQueue

Public Function EnviarDatosASlot(ByVal userindex As Integer, ByRef Datos As String) As Long
      '***************************************************
      'Author: Unknown
      'Last Modification: 01/10/07
      'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
      'Now it uses the clsByteQueue class and don`t make a FIFO Queue of String
      '***************************************************

#If UsarQueSocket = 1 Then '**********************************************
10        On Error GoTo Err
          
          Dim ret As Long
          
20        ret = WsApiEnviar(userindex, Datos)
          
30        If ret <> 0 And ret <> WSAEWOULDBLOCK Then
              ' Close the socket avoiding any critical error
40            Call CloseSocketSL(userindex)
50            Call Cerrar_Usuario(userindex)
60        End If
70    Exit Function
          
Err:

#ElseIf UsarQueSocket = 0 Then '**********************************************
          
80        If frmMain.Socket2(userindex).Write(Datos, Len(Datos)) < 0 Then
90            If frmMain.Socket2(userindex).LastError = WSAEWOULDBLOCK Then
                  ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
100               Call UserList(userindex).outgoingData.WriteASCIIStringFixed(Datos)
110           Else
                  'Close the socket avoiding any critical error
120               Call Cerrar_Usuario(userindex)
130           End If
140       End If
#ElseIf UsarQueSocket = 2 Then '**********************************************

          'Return value for this Socket:
          '--0) OK
          '--1) WSAEWOULDBLOCK
          '--2) ERROR
          
          Dim ret As Long

150       ret = frmMain.Serv.Enviar(.ConnID, Datos, Len(Datos))
                  
160       If ret = 1 Then
              ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
170           Call .outgoingData.WriteASCIIStringFixed(Datos)
180       ElseIf ret = 2 Then
              'Close socket avoiding any critical error
190           Call CloseSocketSL(userindex)
200           Call Cerrar_Usuario(userindex)
210       End If
          

#ElseIf UsarQueSocket = 3 Then
          'THIS SOCKET DOESN`T USE THE BYTE QUEUE CLASS
          Dim rv As Long
          'al carajo, esto encola solo!!! che, me aprobará los
          'parciales también?, este control hace todo solo!!!!
220       On Error GoTo ErrorHandler
              
230           If UserList(userindex).ConnID = -1 Then
240               Call LogError("TCP::EnviardatosASlot, se intento enviar datos a un userIndex con ConnId=-1")
250               Exit Function
260           End If
              
270           If frmMain.TCPServ.Enviar(UserList(userindex).ConnID, Datos, Len(Datos)) = 2 Then Call CloseSocket(userindex)

280   Exit Function
ErrorHandler:
290       Call LogError("TCP::EnviarDatosASlot. UI/ConnId/Datos: " & userindex & "/" & UserList(userindex).ConnID & "/" & Datos)
#End If '**********************************************

End Function
Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim X As Integer, Y As Integer
For Y = UserList(Index).Pos.Y - MinYBorder + 1 To UserList(Index).Pos.Y + MinYBorder - 1
        For X = UserList(Index).Pos.X - MinXBorder + 1 To UserList(Index).Pos.X + MinXBorder - 1

            If MapData(UserList(Index).Pos.map, X, Y).userindex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
        Next X
Next Y
EstaPCarea = False
End Function

Function HayPCarea(Pos As WorldPos) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim X As Integer, Y As Integer
10    For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
20            For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
30                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
40                    If MapData(Pos.map, X, Y).userindex > 0 Then
50                        HayPCarea = True
60                        Exit Function
70                    End If
80                End If
90            Next X
100   Next Y
110   HayPCarea = False
End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim X As Integer, Y As Integer
10    For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
20            For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
30                If MapData(Pos.map, X, Y).ObjInfo.ObjIndex = ObjIndex Then
40                    HayOBJarea = True
50                    Exit Function
60                End If
              
70            Next X
80    Next Y
90    HayOBJarea = False
End Function
Function ValidateChr(ByVal userindex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    ValidateChr = UserList(userindex).Char.Head <> 0 _
                      And UserList(userindex).Char.body <> 0 _
                      And ValidateSkills(userindex)

End Function

Sub ConnectUser(ByVal userindex As Integer, ByRef Name As String)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 24/07/2010 (ZaMa)
'26/03/2009: ZaMa - Agrego por default que el color de dialogo de los dioses, sea como el de su nick.
'12/06/2009: ZaMa - Agrego chequeo de nivel al loguear
'14/09/2009: ZaMa - Ahora el usuario esta protegido del ataque de npcs al loguear
'11/27/2009: Budi - Se envian los InvStats del personaje y su Fuerza y Agilidad
'03/12/2009: Budi - Optimización del código
'24/07/2010: ZaMa - La posicion de comienzo es namehuak, como se habia definido inicialmente.
'***************************************************
Dim n As Integer
Dim tStr As String

   On Error GoTo ConnectUser_Error

10    With UserList(userindex)

20        If .flags.UserLogged Then
30      Call LogCheating("El usuario " & .Name & " ha intentado loguear a " & Name & " desde la IP " & .ip)
        'Kick player ( and leave character inside :D )!
40      Call CloseSocketSL(userindex)
50      Call Cerrar_Usuario(userindex)
60      Exit Sub
70        End If
    
    
    'Sistema de Global Anti.Floodeo de consolas
80        UserList(userindex).ultimoGlobal = GetTickCount()
    
    'Reseteamos los FLAGS
90        .flags.Escondido = 0
100       .flags.TargetNPC = 0
110       .flags.TargetNpcTipo = eNPCType.Comun
120       .flags.TargetObj = 0
130       .flags.TargetUser = 0
140       .Char.FX = 0
150       '.HD = disco '//Disco.
160      ' .CPU_ID = CPU_ID
170       .flags.Compañero = 0
    
    
180       .flags.MenuCliente = 255
190       .flags.LastSlotClient = 255
200       .Stats.OldHp = 0
    
    'Controlamos no pasar el maximo de usuarios
210       If NumUsers >= MaxUsers Then
220     Call WriteErrorMsg(userindex, "El servidor ha alcanzado el máximo de usuarios soportado, por favor vuelva a intertarlo más tarde.")
230     Call FlushBuffer(userindex)
240     Call CloseSocket(userindex)
250     Exit Sub
260       End If
    
    
    '¿Este IP ya esta conectado?
270       If AllowMultiLogins = 0 Then
280     If CheckForSameIP(userindex, .ip) = True Then
290         Call WriteErrorMsg(userindex, "No es posible usar más de un personaje al mismo tiempo.")
300         Call FlushBuffer(userindex)
310         Call CloseSocket(userindex)
320         Exit Sub
330     End If
340       End If
    
    '¿Existe el personaje?
350       If Not FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) Then
360     Call WriteErrorMsg(userindex, "El personaje no existe.")
370     Call FlushBuffer(userindex)
380     Call CloseSocket(userindex)
390     Exit Sub
400       End If
    
        '¿Es el passwd valido?
410       'If UCase$(Password) <> UCase$(GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "Password")) Then
420    ' Call WriteErrorMsg(UserIndex, "Password incorrecto.")
430    ' Call FlushBuffer(UserIndex)
440    ' Call CloseSocket(UserIndex)
450    ' Exit Sub
460    '   End If
    
    '¿Ya esta conectado el personaje?
470       If CheckForSameName(Name) Then
480             If UserList(NameIndex(Name)).Counters.Saliendo Then
490                 Call WriteErrorMsg(userindex, "El usuario está saliendo.")
500             Else
510                 Call WriteErrorMsg(userindex, "Perdón, un usuario con el mismo nombre se ha logueado.")
520             End If

530             Call FlushBuffer(userindex)
540             Call CloseSocket(userindex)
550             Exit Sub
560       End If
    
    'Reseteamos los privilegios
570       .flags.Privilegios = 0
    
    'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
580       If EsAdmin(Name) Then
590     .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
600     Call LogGM(Name, "Se conecto con ip:" & .ip & "/" & .CPU_ID & "/" & .HD)
610       ElseIf EsDios(Name) Then
620     .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
630     Call LogGM(Name, "Se conecto con ip:" & .ip & "/" & .CPU_ID & "/" & .HD)
640       ElseIf EsSemiDios(Name) Then
650     .flags.Privilegios = .flags.Privilegios Or PlayerType.SemiDios
660     Call LogGM(Name, "Se conecto con ip:" & .ip & "/" & .CPU_ID & "/" & .HD)
670       ElseIf EsConsejero(Name) Then
680     .flags.Privilegios = .flags.Privilegios Or PlayerType.Consejero
690     Call LogGM(Name, "Se conecto con ip:" & .ip & "/" & .CPU_ID & "/" & .HD)
700       Else
710     .flags.Privilegios = .flags.Privilegios Or PlayerType.User
720     .flags.AdminPerseguible = True
730       End If
    
    'Add RM flag if needed
740       If EsRolesMaster(Name) Then
750     .flags.Privilegios = .flags.Privilegios Or PlayerType.RoleMaster
760       End If
    
    
770       Call LogUserConect(Name, "Se conecto con ip:" & .ip & "/" & .CPU_ID & "/" & .HD)
    
780       If ServerSoloGMs > 0 Then
790     If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
800         Call WriteErrorMsg(userindex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
810         Call FlushBuffer(userindex)
820         Call CloseSocket(userindex)
830         Exit Sub
840     End If
850       End If
        
    'Cargamos el personaje
     Dim Leer As clsIniManager
860       Set Leer = New clsIniManager
    
870       Call Leer.Initialize(CharPath & UCase$(Name) & ".chr")
    
    'Cargamos los datos del personaje
880       Call LoadUserInit(userindex, Leer)
    
890       Call LoadUserStats(userindex, Leer)
900       Call LoadQuestStats(userindex, Leer)
    
910       If Not ValidateChr(userindex) Then
920     Call WriteErrorMsg(userindex, "Error en el personaje.")
930     Call CloseSocket(userindex)
940     Exit Sub
950       End If
    
960       Call LoadUserReputacion(userindex, Leer)
    
        
    'ADM CHECK HD ---- ivanlisz
970     .StaticHD = Leer.GetValue("INIT", "StaticHD")

980     If LenB(.StaticHD) <> 0 Then
990         If Not .StaticHD = .HD Then
1000            Call WriteErrorMsg(userindex, "Este personaje está protegido.")
1010            Call FlushBuffer(userindex)
1020            Call CloseSocket(userindex)

1030            Exit Sub

1040        End If
1050    End If

1060      Set Leer = Nothing
    
1070      If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
1080      If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
1090      If .Invent.WeaponEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
    
    
1100      If .Invent.EscudoEqpSlot > 25 Then .Invent.EscudoEqpSlot = NingunEscudo
1110      If .Invent.CascoEqpSlot > 25 Then .Invent.CascoEqpSlot = NingunCasco
1120      If .Invent.WeaponEqpSlot > 25 Then .Invent.WeaponEqpSlot = NingunArma
    
    
1130      If .Invent.MochilaEqpSlot > 0 Then
1140    .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + ObjData(.Invent.Object(.Invent.MochilaEqpSlot).ObjIndex).MochilaType * 5
1150      Else
1160    .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
1170      End If
    
1180      If (.flags.Muerto = 0) Then
1190    .flags.SeguroResu = False
1200    Call WriteMultiMessage(userindex, eMessages.ResuscitationSafeOff)
1210      Else
1220    .flags.SeguroResu = True
1230    Call WriteMultiMessage(userindex, eMessages.ResuscitationSafeOn)
1240      End If
    
1250      Call UpdateUserInv(True, userindex, 0)
1260      Call UpdateUserHechizos(True, userindex, 0)
    
1270      If .flags.Paralizado Then
1280    Call WriteParalizeOK(userindex)
1290      End If
    
    Dim Mapa As Integer
1300      Mapa = .Pos.map
    
    'Posicion de comienzo
1310      If Mapa = 0 Then
1320    .Pos = Ullathorpe
1330    Mapa = Ullathorpe.map
1340      Else
    
1350    If Not MapaValido(Mapa) Then
1360        Call WriteErrorMsg(userindex, "El PJ se encuenta en un mapa inválido.")
1370        Call CloseSocket(userindex)
1380        Exit Sub
1390    End If
        
        ' If map has different initial coords, update it
        Dim StartMap As Integer
1400    StartMap = MapInfo(Mapa).StartPos.map
1410    If StartMap <> 0 Then
1420        If MapaValido(StartMap) Then
1430            .Pos = MapInfo(Mapa).StartPos
1440            Mapa = StartMap
1450        End If
1460    End If
        
1470      End If
    
    'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
    'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martín Sotuyo Dodero (Maraxus)
1480      If MapData(Mapa, .Pos.X, .Pos.Y).userindex <> 0 Or MapData(Mapa, .Pos.X, .Pos.Y).NpcIndex <> 0 Then
        Dim FoundPlace As Boolean
        Dim esAgua As Boolean
        Dim tX As Long
        Dim tY As Long
        
1490    FoundPlace = False
1500    esAgua = HayAgua(Mapa, .Pos.X, .Pos.Y)
        
1510    For tY = .Pos.Y - 1 To .Pos.Y + 1
1520        For tX = .Pos.X - 1 To .Pos.X + 1
1530            If esAgua Then
                    'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
1540                If LegalPos(Mapa, tX, tY, True, False) Then
1550                    FoundPlace = True
1560                    Exit For
1570                End If
1580            Else
                    'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
1590                If LegalPos(Mapa, tX, tY, False, True) Then
1600                    FoundPlace = True
1610                    Exit For
1620                End If
1630            End If
1640        Next tX
            
1650        If FoundPlace Then _
                Exit For
1660    Next tY
        
1670    If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
1680        .Pos.X = tX
1690        .Pos.Y = tY
1700    Else
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
1710        If MapData(Mapa, .Pos.X, .Pos.Y).userindex <> 0 Then
               'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
1720            If UserList(MapData(Mapa, .Pos.X, .Pos.Y).userindex).ComUsu.DestUsu > 0 Then
                    'Le avisamos al que estaba comerciando que se tuvo que ir.
1730                If UserList(UserList(MapData(Mapa, .Pos.X, .Pos.Y).userindex).ComUsu.DestUsu).flags.UserLogged Then
1740                    Call FinComerciarUsu(UserList(MapData(Mapa, .Pos.X, .Pos.Y).userindex).ComUsu.DestUsu)
1750                    Call WriteConsoleMsg(UserList(MapData(Mapa, .Pos.X, .Pos.Y).userindex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
1760                    Call FlushBuffer(UserList(MapData(Mapa, .Pos.X, .Pos.Y).userindex).ComUsu.DestUsu)
1770                End If
                    'Lo sacamos.
1780                If UserList(MapData(Mapa, .Pos.X, .Pos.Y).userindex).flags.UserLogged Then
1790                    Call FinComerciarUsu(MapData(Mapa, .Pos.X, .Pos.Y).userindex)
1800                    Call WriteErrorMsg(MapData(Mapa, .Pos.X, .Pos.Y).userindex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
1810                    Call FlushBuffer(MapData(Mapa, .Pos.X, .Pos.Y).userindex)
1820                End If
1830            End If
                
1840            Call CloseSocket(MapData(Mapa, .Pos.X, .Pos.Y).userindex)
1850        End If
1860    End If
1870      End If
    
    'Nombre de sistema
1880      .Name = Name
    
1890      .showName = True 'Por default los nombres son visibles
1900      Call Mod_AntiCheat.SetIntervalos(userindex)
    
    
   'If in the water, and has a boat, equip it!
1910      If .Invent.BarcoObjIndex > 0 And _
            (HayAgua(Mapa, .Pos.X, .Pos.Y) Or BodyIsBoat(.Char.body)) Then

1920    .Char.Head = 0
1930    If .flags.Muerto = 0 Then
1940        Call ToggleBoatBody(userindex)
1950    Else
1960        .Char.body = iFragataFantasmal
1970        .Char.ShieldAnim = NingunEscudo
1980        .Char.WeaponAnim = NingunArma
1990        .Char.CascoAnim = NingunCasco
2000    End If
        
2010    .flags.Navegando = 1
2020      End If
    
    
    
    'Info
2030      Call WriteUserIndexInServer(userindex) 'Enviamos el User index
2040       Call WriteChangeMap(userindex, .Pos.map, MapInfo(.Pos.map).MapVersion)
2050      Call WritePlayMidi(userindex, val(ReadField(1, MapInfo(.Pos.map).Music, 45)))
    
2060       If .flags.Privilegios = PlayerType.Admin Then
2070    .flags.ChatColor = RGB(250, 250, 150)
2080      ElseIf .flags.Privilegios = PlayerType.Dios Then
2090    .flags.ChatColor = RGB(255, 166, 0)
2100      ElseIf .flags.Privilegios <> PlayerType.User And .flags.Privilegios <> (PlayerType.User Or PlayerType.ChaosCouncil) And .flags.Privilegios <> (PlayerType.User Or PlayerType.RoyalCouncil) Then
2110    .flags.ChatColor = RGB(0, 255, 0)
2120      ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.RoyalCouncil) Then
2130    .flags.ChatColor = RGB(0, 255, 255)
2140      ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.ChaosCouncil) Then
2150    .flags.ChatColor = RGB(255, 128, 64)
2160      Else
2170    .flags.ChatColor = vbWhite
2180      End If
    
2190      .flags.ModoCombate = False
2200      .flags.ModoCombate = 0
    
    ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
    #If ConUpTime Then
2210    .LogOnTime = Now
    #End If
    'Crea  el personaje del usuario
2220      Call MakeUserChar(True, .Pos.map, userindex, .Pos.map, .Pos.X, .Pos.Y)
    
2230      If (.flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster)) = 0 Then
2240    Call DoAdminInvisible(userindex)
2250      End If
    
2260      Call WriteUserCharIndexInServer(userindex)
    ''[/el oso]
    
2270      Call DoTileEvents(userindex, .Pos.map, .Pos.X, .Pos.Y)
    
2280      Call CheckUserLevel(userindex)
2290      Call WriteUpdateUserStats(userindex)
    
2300      Call WriteUpdateHungerAndThirst(userindex)
2310      Call WriteUpdateStrenghtAndDexterity(userindex)
    
2320      If haciendoBK Then
2330    Call WritePauseToggle(userindex)
2340    Call WriteConsoleMsg(userindex, "Servidor> Por favor espera algunos segundos, el WorldSave está ejecutándose.", FontTypeNames.FONTTYPE_SERVER)
2350      End If
    
2360      If EnPausa Then
2370    Call WritePauseToggle(userindex)
2380    Call WriteConsoleMsg(userindex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER)
2390      End If
    
2400      If EnTesting And .Stats.ELV >= 18 Then
2410    Call WriteErrorMsg(userindex, "Servidor en Testing por unos minutos, conectese con PJs de nivel menor a 18. No se conecte con Pjs que puedan resultar importantes por ahora pues pueden arruinarse.")
2420    Call FlushBuffer(userindex)
2430    Call CloseSocket(userindex)
2440    Exit Sub
2450      End If
            
            FlushBuffer userindex
    'Actualiza el Num de usuarios
    'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
2460      NumUsers = NumUsers + 1
2470      .flags.UserLogged = True

    'usado para borrar Pjs
2480      Call WriteVar(CharPath & .Name & ".chr", "INIT", "Logged", "1")
    
2490      Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    
2500      MapInfo(.Pos.map).NumUsers = MapInfo(.Pos.map).NumUsers + 1
    
2510      If .Stats.SkillPts > 0 Then
2520    Call WriteSendSkills(userindex)
2530    Call WriteLevelUp(userindex, .Stats.SkillPts)
2540      End If
    
2550      If NumUsers > recordusuarios Then
2560    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("RECORD de usuarios conectados simultaneamente." & "Hay " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_INFO))
2570    recordusuarios = NumUsers
2580    Call WriteVar(IniPath & "Server.ini", "INIT", "RECORD", Str(recordusuarios))
        
2590    Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)
2600      End If
    
2610      If .NroMascotas > 0 And MapInfo(.Pos.map).Pk Then
        Dim i As Integer
2620    For i = 1 To MAXMASCOTAS
2630        If .MascotasType(i) > 0 Then
2640            .MascotasIndex(i) = SpawnNpc(.MascotasType(i), .Pos, True, True)
                
2650            If .MascotasIndex(i) > 0 Then
2660                Npclist(.MascotasIndex(i)).MaestroUser = userindex
2670                Call FollowAmo(.MascotasIndex(i))
2680            Else
2690                .MascotasIndex(i) = 0
2700            End If
2710        End If
2720    Next i
2730      End If
    
2740      If .flags.Navegando = 1 Then
2750            Call WriteNavigateToggle(userindex)
2760      End If
    
2770       If .flags.Montando = 1 Then
2780            Call WriteMontateToggle(userindex)
2790      End If
    
2800      If .ACT Then
2810            WriteMultiMessage userindex, eMessages.DragOnn
2820      .ACT = True
2830      Else
2840        WriteMultiMessage userindex, eMessages.DragOff
2850      .ACT = False
2860      End If

2870      If criminal(userindex) Then
2880    Call WriteMultiMessage(userindex, eMessages.SafeModeOff) 'Call WriteSafeModeOff(UserIndex)
2890    .flags.Seguro = False
2900      Else
2910    .flags.Seguro = True
2920    Call WriteMultiMessage(userindex, eMessages.SafeModeOn) 'Call WriteSafeModeOn(UserIndex)
2930      End If
    
2940      If .GuildIndex > 0 Then
        'welcome to the show baby...
2950    If Not modGuilds.m_ConectarMiembroAClan(userindex, .GuildIndex) Then
2960        Call WriteConsoleMsg(userindex, "Tu estado no te permite entrar al clan.", FontTypeNames.FONTTYPE_GUILD)
2970    End If
2980      End If
    
2990      Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
            .ClaveAC = 0
            Dim xxx As Long
            For xxx = 1 To 7
                .ClaveAC = .ClaveAC + RandomNumber(1, 35)
            Next
            
            'activamos seguro de clan
            UserList(userindex).SeguroClan = True
3000      Call WriteLoggedMessage(userindex)

3010      Call modGuilds.SendGuildNews(userindex)
    
    ' Esta protegido del ataque de npcs por 5 segundos, si no realiza ninguna accion
3020      Call IntervaloPermiteSerAtacado(userindex, True)
    
3030      tStr = modGuilds.a_ObtenerRechazoDeChar(.Name)
    
3040      If LenB(tStr) <> 0 Then
3050    Call WriteShowMessageBox(userindex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
        ' @ Ingresamos al ranking de DS
3060      End If
    
3070      CheckRankingUser userindex, TopFrags
3080      CheckRankingUser userindex, TopLevel
3090      CheckRankingUser userindex, TopOro
3100      CheckRankingUser userindex, TopRetos
3110      CheckRankingUser userindex, TopTorneos
3120      CheckRankingUser userindex, TopClanes
    
    
          WriteUpdatePoints userindex
    'Load the user statistics
3130      Call Statistics.UserConnected(userindex)
       
    
3140      Call MostrarNumUsers

3150      n = FreeFile
3160      Open App.Path & "\logs\numusers.log" For Output As n
3170      Print #n, NumUsers
3180      Close #n
    
3190      n = FreeFile
    'Log
3200      Open App.Path & "\logs\Connect.log" For Append Shared As #n
3210      Print #n, .Name & " ha entrado al juego. UserIndex:" & userindex & " " & time & " " & Date
3220      Close #n
    
    ' Mensajes de entrada
3230      Call WriteShortMsj(userindex, 70, FontTypeNames.FONTTYPE_GUILD)
3240      Call WriteShortMsj(userindex, 71, FontTypeNames.FONTTYPE_GM)
3250      Call WriteShortMsj(userindex, 72, FontTypeNames.FONTTYPE_GUILD)
3260      Call WriteShortMsj(userindex, 73, FontTypeNames.FONTTYPE_CONSEJOVesA)
3270      Call WriteShortMsj(userindex, 74, FontTypeNames.FONTTYPE_EJECUCION)

3280       If .flags.ModoCombate = False Then
3290    WriteConsoleMsg userindex, "No estás en modo combate, para activarlo presiona la tecla C.", FontTypeNames.FONTTYPE_INFO
3300      End If
    
3310      Call LogUserConect(Name, "Se conecto con ip:" & .ip & "/" & .CPU_ID & "/" & .HD)
    
    
3320      If .Mercado.InList > 0 And _
                (.Mercado.Gld > 0 Or .Mercado.Dsp > 0) Then
3330            WriteShortMsj userindex, 75, FontTypeNames.fonttype_dios
3340      End If
     
    
3350      Call CheckLogros(userindex)
    
3360      If IsEffectDios(UCase$(.Name)) Then
3370    WriteShortMsj userindex, 76, FontTypeNames.FONTTYPE_WARNING
3380    .flags.IsDios = True
3390      End If
    
3400  End With

   On Error GoTo 0
   Exit Sub

ConnectUser_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure ConnectUser of Módulo TCP in line " & Erl
End Sub
Sub ResetFaccionCaos(ByVal userindex As Integer)
      '*************************************************
      'Author: Unknown
      'Last modified: 23/01/2007
      'Resetea todos los valores generales y las stats
      '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
      '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
      '*************************************************
10        With UserList(userindex).Faccion
20            .ArmadaReal = 0
30            .FuerzasCaos = 0
40            .FechaIngreso = "No ingresó a ninguna Facción"
50            .CriminalesMatados = 0
60            .RecibioArmaduraCaos = 0
70            .RecibioArmaduraReal = 0
80            .RecibioExpInicialCaos = 0
90            .RecibioExpInicialReal = 0
100           .RecompensasCaos = 0
110           .RecompensasReal = 0
120           .Reenlistadas = 0
130           .NivelIngreso = 0
140           .MatadosIngreso = 0
150           .NextRecompensa = 0
160       End With
End Sub
Sub ResetFaccionReal(ByVal userindex As Integer)
      '*************************************************
      'Author: Unknown
      'Last modified: 23/01/2007
      'Resetea todos los valores generales y las stats
      '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
      '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
      '*************************************************
10        With UserList(userindex).Faccion
20            .ArmadaReal = 0
30            .CiudadanosMatados = 0
40            .FuerzasCaos = 0
50            .FechaIngreso = "No ingresó a ninguna Facción"
60            .RecibioArmaduraCaos = 0
70            .RecibioArmaduraReal = 0
80            .RecibioExpInicialCaos = 0
90            .RecibioExpInicialReal = 0
100           .RecompensasCaos = 0
110           .RecompensasReal = 0
120           .Reenlistadas = 0
130           .NivelIngreso = 0
140           .MatadosIngreso = 0
150           .NextRecompensa = 0
160       End With
End Sub

Sub ResetContadores(ByVal userindex As Integer)
      '*************************************************
      'Author: Unknown
      'Last modified: 03/15/2006
      'Resetea todos los valores generales y las stats
      '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
      '05/20/2007 Integer - Agregue todas las variables que faltaban.
      '*************************************************
10        With UserList(userindex).Counters
20            .TimeFight = 0
30            .TimeAntiFriz = 0
40            .TimeCastleMode = 0
50            .TimePin = 0
60            .TimePotFull = 0
70            .AGUACounter = 0
80            .AttackCounter = 0
90            .Ceguera = 0
100           .COMCounter = 0
110           .Estupidez = 0
120           .Frio = 0
130           .HPCounter = 0
140           .IdleCount = 0
150           .TimeCastleMode = 0
160           .TimeFight = 0
170           .Invisibilidad = 0
180           .Paralisis = 0
190           .Pena = 0
200           .PiqueteC = 0
210           .STACounter = 0
220           .Veneno = 0
230           .Trabajando = 0
240           .Ocultando = 0
              '.bPuedeMeditar = False
250           .Lava = 0
260           .Mimetismo = 0
270           .Saliendo = False
280           .Salir = 0
290           .TiempoOculto = 0
300           .TimerMagiaGolpe = 0
310           .TimerGolpeMagia = 0
320           .TimerLanzarSpell = 0
330           .TimerPuedeAtacar = 0
340           .TimerPuedeUsarArco = 0
350           .TimerPuedeTrabajar = 0
360           .TimerUsar = 0
370           .TimerUsarClick = 0
380           .failedUsageAttempts = 0
390           .goHome = 0
400           .AsignedSkills = 0
410       End With
End Sub

Sub ResetCharInfo(ByVal userindex As Integer)
      '*************************************************
      'Author: Unknown
      'Last modified: 03/15/2006
      'Resetea todos los valores generales y las stats
      '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
      '*************************************************
10        With UserList(userindex).Char
20            .body = 0
30            .CascoAnim = 0
40            .CharIndex = 0
50            .FX = 0
60            .Head = 0
70            .loops = 0
80            .Heading = 0
90            .loops = 0
100           .ShieldAnim = 0
110           .WeaponAnim = 0
120       End With
End Sub

Sub ResetBasicUserInfo(ByVal userindex As Integer)
      '*************************************************
      'Author: Unknown
      'Last modified: 03/15/2006
      'Resetea todos los valores generales y las stats
      '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
      '*************************************************
10        With UserList(userindex)
20            .Name = vbNullString
30            .desc = vbNullString
40            .DescRM = vbNullString
50            .Pos.map = 0
60            .Pos.X = 0
70            .Pos.Y = 0
80            .ip = vbNullString
90            .clase = 0
100           .Email = vbNullString
110           .Genero = 0
120           .Hogar = 0
130           .raza = 0
              
140           .GroupIndex = 0
150           .GroupRequired = 0
              .GroupSlotUser = 0
              
160           With .Stats
170               .Banco = 0
180               .ELV = 0
190               .ELU = 0
200               .Exp = 0
210               .def = 0
                  '.CriminalesMatados = 0
220               .NPCsMuertos = 0
230               .UsuariosMatados = 0
240               .SkillPts = 0
250               .Gld = 0
260               .UserAtributos(1) = 0
270               .UserAtributos(2) = 0
280               .UserAtributos(3) = 0
290               .UserAtributos(4) = 0
300               .UserAtributos(5) = 0
310               .UserAtributosBackUP(1) = 0
320               .UserAtributosBackUP(2) = 0
330               .UserAtributosBackUP(3) = 0
340               .UserAtributosBackUP(4) = 0
350               .UserAtributosBackUP(5) = 0
360           End With
              
370       End With
End Sub

Sub ResetReputacion(ByVal userindex As Integer)
      '*************************************************
      'Author: Unknown
      'Last modified: 03/15/2006
      'Resetea todos los valores generales y las stats
      '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
      '*************************************************
10        With UserList(userindex).Reputacion
20            .AsesinoRep = 0
30            .BandidoRep = 0
40            .BurguesRep = 0
50            .LadronesRep = 0
60            .NobleRep = 0
70            .PlebeRep = 0
80            .NobleRep = 0
90            .Promedio = 0
100       End With
End Sub

Sub ResetGuildInfo(ByVal userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If UserList(userindex).EscucheClan > 0 Then
20            Call modGuilds.GMDejaDeEscucharClan(userindex, UserList(userindex).EscucheClan)
30            UserList(userindex).EscucheClan = 0
40        End If
50        If UserList(userindex).GuildIndex > 0 Then
60            Call modGuilds.m_DesconectarMiembroDelClan(userindex, UserList(userindex).GuildIndex)
70        End If
80        UserList(userindex).GuildIndex = 0
End Sub
Sub ResetUserLogros(ByVal userindex As Integer)
10        With UserList(userindex)
              Dim LoopC As Integer
20            For LoopC = 0 To MAX_LOGROS
30                .Logros(LoopC) = 0
40            Next LoopC
50        End With
End Sub
Sub ResetUserFlags(ByVal userindex As Integer)
      '*************************************************
      'Author: Unknown
      'Last modified: 06/28/2008
      'Resetea todos los valores generales y las stats
      '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
      '03/29/2006 Maraxus - Reseteo el CentinelaOK también.
      '06/28/2008 NicoNZ - Agrego el flag Inmovilizado
      '*************************************************

10        With UserList(userindex)
              .KeyPotas = 0
20            .PotFull = False
30            .FailedPot = 0
40            .PaquetesBasura = 0
50        End With
          
60        With UserList(userindex).PosAnt
70            .map = 0
80            .X = 0
90            .Y = 0
100       End With
          
110       With UserList(userindex).flags
120           .RequestData = vbNullString
130           .InCVC = False
140           .IsDios = False
150           .SlotEvent = 0
160           .SlotReto = 0
170           .SlotUserEvent = 0
180           .SlotRetoUser = 255
190           .SelectedEvent = 0
200           .FightTeam = 0
210           .ElPedidorSeguimiento = 0
220           .MenuCliente = 0 'Esto de MenuCliente es de coso, Hispano i guess.
230           .LastSlotClient = 0
240           .Siguiendo = 0
250           .EnEvento = 0
260           .Compañero = 0
270           .Comerciando = False
280           .Ban = 0
290           .Escondido = 0
300           .DuracionEfecto = 0
310           .NpcInv = 0
320           .StatsChanged = 0
330           .TargetNPC = 0
340           .TargetNpcTipo = eNPCType.Comun
350           .TargetObj = 0
360           .TargetObjMap = 0
370           .TargetObjX = 0
380           .TargetObjY = 0
390           .TargetUser = 0
400           .TipoPocion = 0
410           .TomoPocion = False
420           .Descuento = vbNullString
430           .Hambre = 0
440           .Sed = 0
450           .Descansar = False
460           .ModoCombate = False
470           .Vuela = 0
480           .Navegando = 0
490           .Montando = 0
500           .Oculto = 0
510           .Envenenado = 0
520           .invisible = 0
530           .Paralizado = 0
540           .Inmovilizado = 0
550           .Maldicion = 0
560           .Bendicion = 0
570           .Meditando = 0
580           .Privilegios = 0
590           .PrivEspecial = False
600           .PuedeMoverse = 0
610           .OldBody = 0
620           .OldHead = 0
630           .AdminInvisible = 0
640           .ValCoDe = 0
650           .Hechizo = 0
660           .TimesWalk = 0
670           .StartWalk = 0
680           .CountSH = 0
690           .Silenciado = 0
700           .CentinelaOK = False
710           .AdminPerseguible = False
720           .lastMap = 0
730           .Traveling = 0
740           .AtacablePor = 0
750           .AtacadoPorNpc = 0
760           .AtacadoPorUser = 0
770           .NoPuedeSerAtacado = False
780           .OwnedNpc = 0
790           .ShareNpcWith = 0
800           .EnConsulta = False
810           .Ignorado = False
820       End With
End Sub

Sub ResetUserSpells(ByVal userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim LoopC As Long
10        For LoopC = 1 To MAXUSERHECHIZOS
20            UserList(userindex).Stats.UserHechizos(LoopC) = 0
30        Next LoopC
End Sub

Sub ResetUserPets(ByVal userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim LoopC As Long
          
10        UserList(userindex).NroMascotas = 0
              
20        For LoopC = 1 To MAXMASCOTAS
30            UserList(userindex).MascotasIndex(LoopC) = 0
40            UserList(userindex).MascotasType(LoopC) = 0
50        Next LoopC
End Sub

Sub ResetUserBanco(ByVal userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim LoopC As Long
          
10        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
20              UserList(userindex).BancoInvent.Object(LoopC).Amount = 0
30              UserList(userindex).BancoInvent.Object(LoopC).Equipped = 0
40              UserList(userindex).BancoInvent.Object(LoopC).ObjIndex = 0
50        Next LoopC
          
60        UserList(userindex).BancoInvent.NroItems = 0
End Sub

Public Sub LimpiarComercioSeguro(ByVal userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With UserList(userindex).ComUsu
20            If .DestUsu > 0 Then
30                Call FinComerciarUsu(.DestUsu)
40                Call FinComerciarUsu(userindex)
50            End If
60        End With
End Sub

Sub ResetUserSlot(ByVal userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim i As Long

10    UserList(userindex).ConnIDValida = False
20    UserList(userindex).ConnID = -1
        UserList(userindex).SeguroClan = True

        'elsanto encriptacion dinamica
        UserList(userindex).ClaveAC = 0
        UserList(userindex).CounterAC = 0
30    Call LimpiarComercioSeguro(userindex)
40    Call ResetFaccionCaos(userindex)
50    Call ResetFaccionReal(userindex)
60    Call ResetContadores(userindex)
70    Call ResetGuildInfo(userindex)
80    Call ResetCharInfo(userindex)
90    Call ResetBasicUserInfo(userindex)
100   Call ResetReputacion(userindex)
110   Call ResetUserFlags(userindex)
      Call ResetKeyPackets(userindex)
      Call ResetPointer(userindex, Point_Inv)
      Call ResetPointer(userindex, Point_Spell)
120   Call ResetUserLogros(userindex)
130   Call LimpiarInventario(userindex)
140   Call ResetUserSpells(userindex)
150   Call ResetUserPets(userindex)
160   Call ResetUserBanco(userindex)
170   Call ResetQuestStats(userindex)
180   With UserList(userindex).ComUsu
190       .Acepto = False
          
200       For i = 1 To MAX_OFFER_SLOTS
210           .cant(i) = 0
220           .Objeto(i) = 0
230       Next i
          
240       .GoldAmount = 0
250       .DestNick = vbNullString
260       .DestUsu = 0
270   End With
       
End Sub

Sub CloseUser(ByVal userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

      Dim n As Integer
      Dim map As Integer
      Dim Name As String
      Dim i As Integer

      Dim aN As Integer

20    With UserList(userindex)

          
70        aN = .flags.AtacadoPorNpc
80        If aN > 0 Then
90              Npclist(aN).Movement = Npclist(aN).flags.OldMovement
100             Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
110             Npclist(aN).flags.AttackedBy = vbNullString
120       End If
          
130       aN = .flags.NPCAtacado
140       If aN > 0 Then
150           If Npclist(aN).flags.AttackedFirstBy = .Name Then
160               Npclist(aN).flags.AttackedFirstBy = vbNullString
170           End If
180       End If
190       .flags.AtacadoPorNpc = 0
200       .flags.NPCAtacado = 0
          
210       map = .Pos.map
220       Name = UCase$(.Name)
          
230       .Char.FX = 0
240       .Char.loops = 0
250       Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
          
260       .flags.UserLogged = False
270       .Counters.Saliendo = False
          

          
          'Le devolvemos el body y head originales
350       If .flags.AdminInvisible = 1 Then
360           .Char.body = .flags.OldBody
370           .Char.Head = .flags.OldHead
380           .flags.AdminInvisible = 0
390       End If
          
          'si esta en party le devolvemos la experiencia
400       If .GroupIndex > 0 Then Call mGroup.AbandonateGroup(userindex)
          
          'Save statistics
410       Call Statistics.UserDisconnected(userindex)
          
          ' Grabamos el personaje del usuario
420       Call SaveUser(userindex, CharPath & Name & ".chr")
          
          'usado para borrar Pjs
430       Call WriteVar(CharPath & .Name & ".chr", "INIT", "Logged", "0")

          
          'Quitar el dialogo
          'If MapInfo(Map).NumUsers > 0 Then
          '    Call SendToUserArea(UserIndex, "QDL" & .Char.charindex)
          'End If
          
440       If MapInfo(map).NumUsers > 0 Then
450           Call SendData(SendTarget.ToPCAreaButIndex, userindex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
460       End If
          
          'Borrar el personaje
470       If .Char.CharIndex > 0 Then
480           Call EraseUserChar(userindex, .flags.AdminInvisible = 1)
490       End If
          
          'Borrar mascotas
500       For i = 1 To MAXMASCOTAS
510           If .MascotasIndex(i) > 0 Then
520               If Npclist(.MascotasIndex(i)).flags.NPCActive Then _
                      Call QuitarNPC(.MascotasIndex(i))
530           End If
540       Next i
          
          'Update Map Users
550       MapInfo(map).NumUsers = MapInfo(map).NumUsers - 1
          
560       If MapInfo(map).NumUsers < 0 Then
570           MapInfo(map).NumUsers = 0
580       End If
          
          ' Si el usuario habia dejado un msg en la gm's queue lo borramos
590       If Ayuda.Existe(.Name) Then Call Ayuda.Quitar(.Name)
          
600       Call ResetUserSlot(userindex)
          
610       Call MostrarNumUsers
          
620       n = FreeFile(1)
630       Open App.Path & "\logs\Connect.log" For Append Shared As #n
640           Print #n, Name & " ha dejado el juego. " & "User Index:" & userindex & " " & time & " " & Date
650       Close #n
660   End With

670   Exit Sub

Errhandler:
680   Call LogError("Error en CloseUser. Número " & Err.Number & " Descripción: " & Err.Description)

End Sub

Sub ReloadSokcet()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler
#If UsarQueSocket = 1 Then

20        Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
          
30        If NumUsers <= 0 Then
40            Call WSApiReiniciarSockets
50        Else
      '       Call apiclosesocket(SockListen)
      '       SockListen = ListenForConnect(Puerto, hWndMsg, "")
60        End If

#ElseIf UsarQueSocket = 0 Then

70        frmMain.Socket1.Cleanup
80        Call ConfigListeningSocket(frmMain.Socket1, Puerto)
          
#ElseIf UsarQueSocket = 2 Then

          

#End If

90    Exit Sub
Errhandler:
100       Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.Description)

End Sub

Public Sub EnviarNoche(ByVal userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        Call WriteSendNight(userindex, IIf(DeNoche And (MapInfo(UserList(userindex).Pos.map).Zona = Campo Or MapInfo(UserList(userindex).Pos.map).Zona = Ciudad), True, False))
20        Call WriteSendNight(userindex, IIf(DeNoche, True, False))
End Sub

Public Sub EcharPjsNoPrivilegiados()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim LoopC As Long
          
10        For LoopC = 1 To LastUser
20            If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
30                If UserList(LoopC).flags.Privilegios And PlayerType.User Then
40                    Call CloseSocket(LoopC)
50                End If
60            End If
70        Next LoopC

End Sub
