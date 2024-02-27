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

Sub DarCuerpoYCabeza(ByVal UserIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 14/03/2007
'Elije una cabeza para el usuario y le da un body
'*************************************************
Dim NewBody As Integer
Dim NewHead As Integer
Dim UserRaza As Byte
Dim UserGenero As Byte
UserGenero = UserList(UserIndex).Genero
UserRaza = UserList(UserIndex).Raza
Select Case UserGenero
   Case eGenero.Hombre
        Select Case UserRaza
            Case eRaza.Humano
                NewHead = RandomNumber(1, 25)
                NewBody = 1
            Case eRaza.Elfo
                NewHead = RandomNumber(102, 111)
                NewBody = 2
            Case eRaza.Drow
                NewHead = RandomNumber(201, 205)
                NewBody = 3
            Case eRaza.Enano
                NewHead = RandomNumber(301, 305)
                NewBody = 52
            Case eRaza.Gnomo
                NewHead = RandomNumber(401, 405)
                NewBody = 52
        End Select
   Case eGenero.Mujer
        Select Case UserRaza
            Case eRaza.Humano
                NewHead = RandomNumber(71, 77)
                NewBody = 44
            Case eRaza.Elfo
                NewHead = RandomNumber(170, 174)
                NewBody = 44
            Case eRaza.Drow
                NewHead = RandomNumber(270, 276)
                NewBody = 44
            Case eRaza.Gnomo
                NewHead = RandomNumber(471, 475)
                NewBody = 138
            Case eRaza.Enano
                NewHead = RandomNumber(370, 371)
                NewBody = 138
        End Select
End Select
UserList(UserIndex).Char.Head = NewHead
UserList(UserIndex).Char.body = NewBody
End Sub
Function AsciiValidos(ByVal cad As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim Car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    Car = Asc(mid$(cad, i, 1))
    
    If (Car < 97 Or Car > 122) And (Car <> 255) And (Car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim Car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    Car = Asc(mid$(cad, i, 1))
    
    If (Car < 48 Or Car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next i

Numeric = True

End Function


Function NombrePermitido(ByVal Nombre As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim i As Integer

For i = 1 To UBound(ForbidenNames)
    If InStr(Nombre, ForbidenNames(i)) Then
            NombrePermitido = False
            Exit Function
    End If
Next i

NombrePermitido = True

End Function
Function ValidateSkills(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then
        Exit Function
        If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100
    End If
Next LoopC

ValidateSkills = True
    
End Function
Public Sub KillCharINFO(ByVal User As String)
On Error Resume Next
 
Dim c As String
Dim d As String
Dim f As String
Dim g As String
Dim h As Byte
Dim i As String
Dim j As String
 
 
c = GetVar(App.Path & "\CHARFILE\" & User & ".chr", "GUILD", "GUILDINDEX")
d = GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & c, "founder")
f = GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & c, "GuildName")
g = GetVar(App.Path & "\guilds\" & f & "-members.mem", "INIT", "NroMembers")
j = GetVar(App.Path & "\guilds\" & f & "-members.mem", "Members", "Member" & g)
 
 
 
    If c = "" Then
        Kill (App.Path & "\CHARFILE\" & User & ".chr")
    Else
   
        If d <> User Then
            guilds(c).ExpulsarMiembro (User)
        Else
            For h = 1 To g
           
    i = GetVar(App.Path & "\guilds\" & f & "-members.mem", "Members", "Member" & h)
 
                If i = User Then Call WriteVar(App.Path & "\guilds\" & f & "-members.mem", "Members", "Member" & h, j): Call WriteVar(App.Path & "\guilds\" & f & "-members.mem", "INIT", "NroMembers", g - 1)
 
    Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & c, "EleccionesAbiertas", "1")
    Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & c, "EleccionesFinalizan", DateAdd("d", 1, Now))
Call WriteVar(App.Path & "\guilds\" & f & "-votaciones.vot", "INIT", "NumVotos", "0")
            Next h
           
        End If
    Kill (App.Path & "\CHARFILE\" & User & ".chr")
   End If
 
End Sub
       
           
Public Function GenerateRandomKey(ByVal User As String) As Integer
On Error Resume Next
 
        GenerateRandomKey = RandomNumber(500, 1600) + RandomNumber(5000, 16000)
        Call WriteVar(App.Path & "\CHARFILE\" & User & ".chr", "INIT", "Password", GenerateRandomKey)
 
End Function

Sub ConnectNewUser(ByVal UserIndex As Integer, ByRef name As String, ByRef Password As String, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal UserClase As eClass, _
                    ByRef UserEmail As String, ByVal Hogar As eCiudad, ByVal Head As Integer, Pin As String, disco As String, CPU_ID As String)
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

With UserList(UserIndex)

    If Not AsciiValidos(name) Or LenB(name) = 0 Then
        Call WriteErrorMsg(UserIndex, "Nombre inválido.")
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.UserLogged Then
        Call LogCheating("El usuario " & UserList(UserIndex).name & " ha intentado crear a " & name & " desde la IP " & UserList(UserIndex).ip)
        
        'Kick player ( and leave character inside :D )!
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
        
        Exit Sub
    End If
    
    '¿Existe el personaje?
    If FileExist(CharPath & UCase$(name) & ".chr", vbNormal) = True Then
        Call WriteErrorMsg(UserIndex, "Ya existe el personaje.")
        Exit Sub
    End If
    
    'Tiró los dados antes de llegar acá??
    If .Stats.UserAtributos(eAtributos.Fuerza) = 0 Then
        Call WriteErrorMsg(UserIndex, "Debe tirar los dados antes de poder crear un personaje.")
        Exit Sub
    End If
    

    .flags.Muerto = 0
    .flags.Oro = 0
    .flags.Plata = 0
    .flags.Bronce = 0
    .flags.DiosTerrenal = 0
    .flags.Escondido = 0
    .flags.ModoCombate = 0
    .Reputacion.AsesinoRep = 0
    .Reputacion.BandidoRep = 0
    .Reputacion.BurguesRep = 0
    .Reputacion.LadronesRep = 0
    .Reputacion.PlebeRep = 30
    
    .Reputacion.Promedio = 30 / 6
    
    
    .name = name
    .clase = UserClase
    .Raza = UserRaza
    .Genero = UserSexo
    .Email = UserEmail
    .Hogar = Hogar
    
    .Pin = Pin
    .flags.Premium = 0
    
    .CPU_ID = CPU_ID
    
    '[Pablo (Toxic Waste) 9/01/08]
    .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + ModRaza(UserRaza).Fuerza
    .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + ModRaza(UserRaza).Agilidad
    .Stats.UserAtributos(eAtributos.Inteligencia) = .Stats.UserAtributos(eAtributos.Inteligencia) + ModRaza(UserRaza).Inteligencia
    .Stats.UserAtributos(eAtributos.Carisma) = .Stats.UserAtributos(eAtributos.Carisma) + ModRaza(UserRaza).Carisma
    .Stats.UserAtributos(eAtributos.Constitucion) = .Stats.UserAtributos(eAtributos.Constitucion) + ModRaza(UserRaza).Constitucion
    '[/Pablo (Toxic Waste)]
    
    For i = 1 To NUMSKILLS
        .Stats.UserSkills(i) = 0
        Call CheckEluSkill(UserIndex, i, True)
    Next i
    
    .Stats.SkillPts = 10
    
    .Char.Heading = eHeading.SOUTH
    
Call DarCuerpoYCabeza(UserIndex)
    
    .OrigChar = .Char
    
    Dim MiInt As Long
    MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Constitucion) \ 3)
    
    .Stats.MaxHp = 15 + MiInt
    .Stats.MinHp = 15 + MiInt

    MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Agilidad) \ 6)
    If MiInt = 1 Then MiInt = 2
    
    .Stats.MaxSta = 20 * MiInt
    .Stats.MinSta = 20 * MiInt
    
    
    .Stats.MaxAGU = 100
    .Stats.MinAGU = 100
    
    .Stats.MaxHam = 100
    .Stats.MinHam = 100
    
    
    .Stats.RetosGanados = 0
    .Stats.RetosPerdidos = 0
    .Stats.OroGanado = 0
    .Stats.OroPerdido = 0
    
    .Stats.TorneosGanados = 0
    
    
    '<-----------------MANA----------------------->
    If UserClase = eClass.Mage Then 'Cambio en mana inicial (ToxicWaste)
        MiInt = RandomNumber(100, 106)
        .Stats.MaxMAN = MiInt
        .Stats.MinMAN = MiInt
    ElseIf UserClase = eClass.Cleric Or UserClase = eClass.Druid _
        Or UserClase = eClass.Bard Or UserClase = eClass.Assasin Then
            .Stats.MaxMAN = 50
            .Stats.MinMAN = 50
    ElseIf UserClase = eClass.Bandit Then 'Mana Inicial del Bandido (ToxicWaste)
            .Stats.MaxMAN = 50
            .Stats.MinMAN = 50
    Else
        .Stats.MaxMAN = 0
        .Stats.MinMAN = 0
    End If
    
    If UserClase = eClass.Mage Or UserClase = eClass.Cleric Or _
       UserClase = eClass.Druid Or UserClase = eClass.Bard Or _
       UserClase = eClass.Assasin Then
            .Stats.UserHechizos(1) = 2
        
            'If UserClase = eClass.Druid Then .Stats.UserHechizos(2) = 46
    End If
    
    .Stats.MaxHIT = 2
    .Stats.MinHIT = 1
    
    .Stats.GLD = 30000
    
    .Stats.Exp = 0
    .Stats.ELU = 300
    .Stats.ELV = 1
    .flags.BonosHP = 0
    
    '???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
    Dim Slot As Byte
    Dim IsPaladin As Boolean
    
    IsPaladin = UserClase = eClass.Paladin
    
    'Pociones Rojas (Newbie)
    Slot = 1
    .Invent.Object(Slot).objindex = 461
    .Invent.Object(Slot).Amount = 50
    
    'Pociones azules (Newbie)
    If .Stats.MaxMAN > 0 Or IsPaladin Then
        Slot = Slot + 1
        .Invent.Object(Slot).objindex = 462
        .Invent.Object(Slot).Amount = 50
    
    End If
    
    ' Ropa (Newbie)
    Slot = Slot + 1
    Select Case UserRaza
        Case eRaza.Humano
            .Invent.Object(Slot).objindex = 463
        Case eRaza.Elfo
            .Invent.Object(Slot).objindex = 464
        Case eRaza.Drow
            .Invent.Object(Slot).objindex = 465
        Case eRaza.Enano
            .Invent.Object(Slot).objindex = 466
        Case eRaza.Gnomo
            .Invent.Object(Slot).objindex = 466
    End Select
    
    ' Equipo ropa
    .Invent.Object(Slot).Amount = 1
    .Invent.Object(Slot).Equipped = 1
    
    .Invent.ArmourEqpSlot = Slot
    .Invent.ArmourEqpObjIndex = .Invent.Object(Slot).objindex

    'Arma (Newbie)
    Slot = Slot + 1
    Select Case UserClase
        Case eClass.Hunter
            ' Arco (Newbie)
            .Invent.Object(Slot).objindex = 460
        Case eClass.Worker
            ' Herramienta (Newbie)
            .Invent.Object(Slot).objindex = 460
        Case Else
            ' Daga (Newbie)
            .Invent.Object(Slot).objindex = 460
    End Select
    
    ' Equipo arma
    .Invent.Object(Slot).Amount = 1
    .Invent.Object(Slot).Equipped = 1
    
    .Invent.WeaponEqpObjIndex = .Invent.Object(Slot).objindex
    .Invent.WeaponEqpSlot = Slot
    
    .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Invent.WeaponEqpObjIndex)

    ' Manzanas (Newbie)
    Slot = Slot + 1
    .Invent.Object(Slot).objindex = 467
    .Invent.Object(Slot).Amount = 100
    
    ' Jugos (Nwbie)
    Slot = Slot + 1
    .Invent.Object(Slot).objindex = 468
    .Invent.Object(Slot).Amount = 100
    
    ' Sin casco y escudo
    .Char.ShieldAnim = NingunEscudo
    .Char.CascoAnim = NingunCasco
    
    ' Total Items
    .Invent.NroItems = Slot
    
    #If ConUpTime Then
        .LogOnTime = Now
        .UpTime = 0
    #End If

 If .flags.ModoCombate = False Then
        WriteConsoleMsg UserIndex, "No estás en modo combate, para activarlo presiona la tecla C.", FontTypeNames.FONTTYPE_INFO
End If
End With



'Valores Default de facciones al Activar nuevo usuario
Call ResetFaccionCaos(UserIndex)
Call ResetFaccionReal(UserIndex)

Call WriteVar(CharPath & UCase$(name) & ".chr", "INIT", "Password", Password) 'grabamos el password aqui afuera, para no mantenerlo cargado en memoria

Call WriteVar(CharPath & UCase$(name) & ".chr", "INIT", "Pin", Pin)

Call SaveUser(UserIndex, CharPath & UCase$(name) & ".chr")
Dim SABPO As String
SABPO = CharPath & UCase$(name) & ".chr"

'Open User
Call ConnectUser(UserIndex, name, Password, disco, CPU_ID)
  
End Sub

#If UsarQueSocket = 1 Or UsarQueSocket = 2 Then

Sub CloseSocket(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler
    
    Call aDos.RestarConexion(UserList(UserIndex).ip)
    
    If UserIndex = LastUser Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser < 1 Then Exit Do
        Loop
    End If
     
    If UserList(UserIndex).flags.Plantico = True Then
    Call Rondas_plantadorDesconecta(UserIndex)
    End If
    
    If UserList(UserIndex).flags.Montando = 0 Then
    End If
     
     If UserList(UserIndex).flags.automatico = True Then
     Call Rondas_UsuarioDesconecta(UserIndex)
     End If
     

                        ' @@ Si un GM me tiene fichado entonces..
 
    If UserList(UserIndex).flags.ElPedidorSeguimiento > 0 Then
 
        Call WriteConsoleMsg(UserList(UserIndex).flags.ElPedidorSeguimiento, "El usuario que estás siguiendo se desconectó, re putin", FontTypeNames.fonttype_dios)
       
        Call Protocol.WriteShowPanelSeguimiento(UserList(UserIndex).flags.ElPedidorSeguimiento, 2)
 
       
    End If

'1vs1
  With UserList(UserIndex)
    If .Reto1vs1.RetoIndex <> 0 Then Retos1vs1.Muere UserIndex, True
     End With
     
Call eventClose(UserIndex)
     


If UserList(UserIndex).flags.Angel = 1 Then
'A resetear variables...
UserList(UserIndex).Stats.MaxHp = UserList(UserIndex).Stats.ViejaHP
UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.ViejaMan
UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.ViejaminHP
UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.ViejaminMan
UserList(UserIndex).flags.Angel = 0
End If

        If UserList(UserIndex).flags.Infectado = 1 Then
'A resetear variables...
UserList(UserIndex).Stats.MaxHp = UserList(UserIndex).Stats.ViejaHP
UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.ViejaMan
UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.ViejaminHP
UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.ViejaminMan
UserList(UserIndex).flags.Infectado = 0
End If

        If UserList(UserIndex).flags.Demonio = 1 Then
'A resetear variables...
UserList(UserIndex).Stats.MaxHp = UserList(UserIndex).Stats.ViejaHP
UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.ViejaMan
UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.ViejaminHP
UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.ViejaminMan
UserList(UserIndex).flags.Demonio = 0
End If

If UserList(UserIndex).death Then Call ModDeath.MuereUser(UserIndex)
If UserList(UserIndex).hungry Then Call Mod_Jdh.MuereUser(UserIndex)



'Deathmatch Automatico
    
    'Call SecurityIp.IpRestarConexion(GetLongIp(UserList(UserIndex).ip))

    If UserList(UserIndex).ConnID <> -1 Then
        Call CloseSocketSL(UserIndex)
    End If
    
    'Es el mismo user al que está revisando el centinela??
    'IMPORTANTE!!! hacerlo antes de resetear así todavía sabemos el nombre del user
    ' y lo podemos loguear
    If Centinela.RevisandoUserIndex = UserIndex Then _
        Call modCentinela.CentinelaUserLogout
    
    
    'mato los comercios seguros
    If UserList(UserIndex).ComUsu.DestUsu > 0 Then
        If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
            If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call WriteConsoleMsg(UserList(UserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                Call FlushBuffer(UserList(UserIndex).ComUsu.DestUsu)
            End If
        End If
    End If
    
    'Empty buffer for reuse
    Call UserList(UserIndex).incomingData.ReadASCIIStringFixed(UserList(UserIndex).incomingData.length)
    
    If UserList(UserIndex).flags.UserLogged Then
        If NumUsers > 0 Then NumUsers = NumUsers - 1
        Call CloseUser(UserIndex)
        Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    Else
        Call ResetUserSlot(UserIndex)
    End If
    
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False
    
Exit Sub

errhandler:
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False
    Call ResetUserSlot(UserIndex)

    Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.description & " - UserIndex = " & UserIndex)

End Sub

#ElseIf UsarQueSocket = 0 Then

Sub CloseSocket(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler
    
    UserList(UserIndex).ConnID = -1


    If UserIndex = LastUser And LastUser > 1 Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop
    End If

    If UserList(UserIndex).flags.UserLogged Then
            If NumUsers <> 0 Then NumUsers = NumUsers - 1
            Call CloseUser(UserIndex)
    End If

    frmMain.Socket2(UserIndex).Cleanup
    Unload frmMain.Socket2(UserIndex)
    Call ResetUserSlot(UserIndex)

Exit Sub

errhandler:
    UserList(UserIndex).ConnID = -1
    Call ResetUserSlot(UserIndex)
End Sub


#ElseIf UsarQueSocket = 3 Then

Sub CloseSocket(ByVal UserIndex As Integer, Optional ByVal cerrarlo As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo errhandler

Dim NURestados As Boolean
Dim CoNnEcTiOnId As Long


    NURestados = False
    CoNnEcTiOnId = UserList(UserIndex).ConnID
    
    'call logindex(UserIndex, "******> Sub CloseSocket. ConnId: " & CoNnEcTiOnId & " Cerrarlo: " & Cerrarlo)
    
    
  
    UserList(UserIndex).ConnID = -1 'inabilitamos operaciones en socket

    If UserIndex = LastUser And LastUser > 1 Then
        Do
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop While UserList(LastUser).flags.UserLogged = True
    End If

    If UserList(UserIndex).flags.UserLogged Then
            If NumUsers <> 0 Then NumUsers = NumUsers - 1
            NURestados = True
            Call CloseUser(UserIndex)
    End If
    
    Call ResetUserSlot(UserIndex)
    
    'limpiada la userlist... reseteo el socket, si me lo piden
    'Me lo piden desde: cerrada intecional del servidor (casi todas
    'las llamadas a CloseSocket del codigo)
    'No me lo piden desde: disconnect remoto (el on_close del control
    'de alejo realiza la desconexion automaticamente). Esto puede pasar
    'por ejemplo, si el cliente cierra el AO.
    If cerrarlo Then Call frmMain.TCPServ.CerrarSocket(CoNnEcTiOnId)

Exit Sub

errhandler:
    Call LogError("CLOSESOCKETERR: " & Err.description & " UI:" & UserIndex)
    
    If Not NURestados Then
        If UserList(UserIndex).flags.UserLogged Then
            If NumUsers > 0 Then
                NumUsers = NumUsers - 1
            End If
            Call LogError("Cerre sin grabar a: " & UserList(UserIndex).name)
        End If
    End If
    
    Call LogError("El usuario no guardado tenía connid " & CoNnEcTiOnId & ". Socket no liberado.")
    Call ResetUserSlot(UserIndex)
    
    

End Sub


#End If

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

#If UsarQueSocket = 1 Then

If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
    Call BorraSlotSock(UserList(UserIndex).ConnID)
    Call WSApiCloseSocket(UserList(UserIndex).ConnID)
    UserList(UserIndex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 0 Then

If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
    frmMain.Socket2(UserIndex).Cleanup
    Unload frmMain.Socket2(UserIndex)
    UserList(UserIndex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 2 Then

If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
    Call frmMain.Serv.CerrarSocket(UserList(UserIndex).ConnID)
    UserList(UserIndex).ConnIDValida = False
End If

#End If
End Sub

''
' Send an string to a Slot
'
' @param userIndex The index of the User
' @param Datos The string that will be send
' @remarks If UsarQueSocket is 3 it won`t use the clsByteQueue

Public Function EnviarDatosASlot(ByVal UserIndex As Integer, ByRef Datos As String) As Long
'***************************************************
'Author: Unknown
'Last Modification: 01/10/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Now it uses the clsByteQueue class and don`t make a FIFO Queue of String
'***************************************************

#If UsarQueSocket = 1 Then '**********************************************
    On Error GoTo Err
    
    Dim ret As Long
    
    ret = WsApiEnviar(UserIndex, Datos)
    
    If ret <> 0 And ret <> WSAEWOULDBLOCK Then
        ' Close the socket avoiding any critical error
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
    End If
Exit Function
    
Err:

#ElseIf UsarQueSocket = 0 Then '**********************************************
    
    If frmMain.Socket2(UserIndex).Write(Datos, Len(Datos)) < 0 Then
        If frmMain.Socket2(UserIndex).LastError = WSAEWOULDBLOCK Then
            ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
            Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(Datos)
        Else
            'Close the socket avoiding any critical error
            Call Cerrar_Usuario(UserIndex)
        End If
    End If
#ElseIf UsarQueSocket = 2 Then '**********************************************

    'Return value for this Socket:
    '--0) OK
    '--1) WSAEWOULDBLOCK
    '--2) ERROR
    
    Dim ret As Long

    ret = frmMain.Serv.Enviar(.ConnID, Datos, Len(Datos))
            
    If ret = 1 Then
        ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
        Call .outgoingData.WriteASCIIStringFixed(Datos)
    ElseIf ret = 2 Then
        'Close socket avoiding any critical error
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
    End If
    

#ElseIf UsarQueSocket = 3 Then
    'THIS SOCKET DOESN`T USE THE BYTE QUEUE CLASS
    Dim rv As Long
    'al carajo, esto encola solo!!! che, me aprobará los
    'parciales también?, este control hace todo solo!!!!
    On Error GoTo ErrorHandler
        
        If UserList(UserIndex).ConnID = -1 Then
            Call LogError("TCP::EnviardatosASlot, se intento enviar datos a un userIndex con ConnId=-1")
            Exit Function
        End If
        
        If frmMain.TCPServ.Enviar(UserList(UserIndex).ConnID, Datos, Len(Datos)) = 2 Then Call CloseSocket(UserIndex)

Exit Function
ErrorHandler:
    Call LogError("TCP::EnviarDatosASlot. UI/ConnId/Datos: " & UserIndex & "/" & UserList(UserIndex).ConnID & "/" & Datos)
#End If '**********************************************

End Function
Function EstaPCarea(index As Integer, Index2 As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim X As Integer, Y As Integer
For Y = UserList(index).Pos.Y - MinYBorder + 1 To UserList(index).Pos.Y + MinYBorder - 1
        For X = UserList(index).Pos.X - MinXBorder + 1 To UserList(index).Pos.X + MinXBorder - 1

            If MapData(UserList(index).Pos.map, X, Y).UserIndex = Index2 Then
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
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(Pos.map, X, Y).UserIndex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next X
Next Y
HayPCarea = False
End Function

Function HayOBJarea(Pos As WorldPos, objindex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If MapData(Pos.map, X, Y).ObjInfo.objindex = objindex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next X
Next Y
HayOBJarea = False
End Function
Function ValidateChr(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

ValidateChr = UserList(UserIndex).Char.Head <> 0 _
                And UserList(UserIndex).Char.body <> 0 _
                And ValidateSkills(UserIndex)

End Function

Sub ConnectUser(ByVal UserIndex As Integer, ByRef name As String, ByRef Password As String, disco As String, CPU_ID As String)
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
Dim N As Integer
Dim tStr As String

With UserList(UserIndex)

    If .flags.UserLogged Then
        Call LogCheating("El usuario " & .name & " ha intentado loguear a " & name & " desde la IP " & .ip)
        'Kick player ( and leave character inside :D )!
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
        Exit Sub
    End If
    
    
    Call WriteConsoleMsg(UserIndex, ">Bienvenido a Desterium AO noble aventurero.", FontTypeNames.FONTTYPE_GUILD)
                     Call WriteConsoleMsg(UserIndex, ">El máximo de oro es 50.000.000 por personaje.", FontTypeNames.FONTTYPE_GM)
                     Call WriteConsoleMsg(UserIndex, ">Recuerda visitar nuestro sitio oficial: http://desterium-ao.com/.", FontTypeNames.FONTTYPE_GUILD)
                     Call WriteConsoleMsg(UserIndex, ">NO INSULTES. NO CHITEES. NO JUEGUES SUCIO. NO VALE LA PENA.", FontTypeNames.FONTTYPE_CONSEJOVesA)
                     Call WriteConsoleMsg(UserIndex, ">El uso de cheats será sancionado de forma estricta por el Staff.", FontTypeNames.FONTTYPE_EJECUCION)
    
    'Sistema de Global Anti.Floodeo de consolas
    UserList(UserIndex).ultimoGlobal = GetTickCount()
    
    'Reseteamos los FLAGS
    .flags.Escondido = 0
    .flags.TargetNPC = 0
    .flags.TargetNpcTipo = eNPCType.Comun
    .flags.TargetObj = 0
    .flags.TargetUser = 0
    .Char.FX = 0
    .flags.indexmercado = 0
    .HD = disco '//Disco.
    .CPU_ID = CPU_ID
    .flags.Compañero = 0
    
            .flags.MenuCliente = 255
        .flags.LastSlotClient = 255
    
    'Controlamos no pasar el maximo de usuarios
    If NumUsers >= MaxUsers Then
        Call WriteErrorMsg(UserIndex, "El servidor ha alcanzado el máximo de usuarios soportado, por favor vuelva a intertarlo más tarde.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    '¿Este IP ya esta conectado?
    If AllowMultiLogins = 0 Then
        If CheckForSameHD(UserIndex, .HD) = True Then
            Call WriteErrorMsg(UserIndex, "No es posible usar más de un personaje al mismo tiempo.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    End If
    
    '¿Existe el personaje?
    If Not FileExist(CharPath & UCase$(name) & ".chr", vbNormal) Then
        Call WriteErrorMsg(UserIndex, "El personaje no existe.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    '¿Es el passwd valido?
    If UCase$(Password) <> UCase$(GetVar(CharPath & UCase$(name) & ".chr", "INIT", "Password")) Then
        Call WriteErrorMsg(UserIndex, "Password incorrecto.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    '¿Ya esta conectado el personaje?
    If CheckForSameName(name) Then
        If UserList(NameIndex(name)).Counters.Saliendo Then
            Call WriteErrorMsg(UserIndex, "El usuario está saliendo.")
        Else
            Call WriteErrorMsg(UserIndex, "Perdón, un usuario con el mismo nombre se ha logueado.")
        End If
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    'Reseteamos los privilegios
    .flags.Privilegios = 0
    
    'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
    If EsAdmin(name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
        Call LogGM(name, "Se conecto con ip:" & .ip & .CPU_ID)
    ElseIf EsDios(name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
        Call LogGM(name, "Se conecto con ip:" & .ip & .CPU_ID)
    ElseIf EsSemiDios(name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.SemiDios
        
        .flags.PrivEspecial = EsGmEspecial(name)
        
        Call LogGM(name, "Se conecto con ip:" & .ip & .CPU_ID)
    ElseIf EsConsejero(name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Consejero
        Call LogGM(name, "Se conecto con ip:" & .ip & .CPU_ID)
    ElseIf EsUsuario(name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.User
        Call LogUserConect(name, "Se conecto con ip:" & .ip & .CPU_ID)
        .flags.AdminPerseguible = True
    End If
    
    'Add RM flag if needed
    If EsRolesMaster(name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.RoleMaster
    End If
    
    If ServerSoloGMs > 0 Then
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
            Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    End If
    
    'Cargamos el personaje
     Dim Leer As clsIniManager
    Set Leer = New clsIniManager
    
    Call Leer.Initialize(CharPath & UCase$(name) & ".chr")
    
    'Cargamos los datos del personaje
    Call LoadUserInit(UserIndex, Leer)
    
    Call LoadUserStats(UserIndex, Leer)
  '  Call LoadQuestStats(Userindex, Leer)
    
    If Not ValidateChr(UserIndex) Then
        Call WriteErrorMsg(UserIndex, "Error en el personaje.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    Call LoadUserReputacion(UserIndex, Leer)
    
    Set Leer = Nothing
    
    If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
    If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
    If .Invent.WeaponEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
    
    If .Invent.MochilaEqpSlot > 0 Then
        .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + ObjData(.Invent.Object(.Invent.MochilaEqpSlot).objindex).MochilaType * 5
    Else
        .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
    End If
    If (.flags.Muerto = 0) Then
        .flags.SeguroResu = False
        Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff)
    Else
        .flags.SeguroResu = True
        Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn)
    End If
    
    Call UpdateUserInv(True, UserIndex, 0)
    Call UpdateUserHechizos(True, UserIndex, 0)
    
    If .flags.Paralizado Then
        Call WriteParalizeOK(UserIndex)
    End If
    
    Dim mapa As Integer
    mapa = .Pos.map
    
    'Posicion de comienzo
    If mapa = 0 Then
        .Pos = Ullathorpe
        mapa = Ullathorpe.map
    Else
    
        If Not MapaValido(mapa) Then
            Call WriteErrorMsg(UserIndex, "El PJ se encuenta en un mapa inválido.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
        ' If map has different initial coords, update it
        Dim StartMap As Integer
        StartMap = MapInfo(mapa).StartPos.map
        If StartMap <> 0 Then
            If MapaValido(StartMap) Then
                .Pos = MapInfo(mapa).StartPos
                mapa = StartMap
            End If
        End If
        
    End If

    
    'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
    'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martín Sotuyo Dodero (Maraxus)
    If MapData(mapa, .Pos.X, .Pos.Y).UserIndex <> 0 Or MapData(mapa, .Pos.X, .Pos.Y).NpcIndex <> 0 Then
        Dim FoundPlace As Boolean
        Dim esAgua As Boolean
        Dim tX As Long
        Dim tY As Long
        
        FoundPlace = False
        esAgua = HayAgua(mapa, .Pos.X, .Pos.Y)
        
        For tY = .Pos.Y - 1 To .Pos.Y + 1
            For tX = .Pos.X - 1 To .Pos.X + 1
                If esAgua Then
                    'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
                    If LegalPos(mapa, tX, tY, True, False) Then
                        FoundPlace = True
                        Exit For
                    End If
                Else
                    'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
                    If LegalPos(mapa, tX, tY, False, True) Then
                        FoundPlace = True
                        Exit For
                    End If
                End If
            Next tX
            
            If FoundPlace Then _
                Exit For
        Next tY
        
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            .Pos.X = tX
            .Pos.Y = tY
        Else
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
            If MapData(mapa, .Pos.X, .Pos.Y).UserIndex <> 0 Then
               'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu > 0 Then
                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                        Call FinComerciarUsu(UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu)
                        Call WriteConsoleMsg(UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                        Call FlushBuffer(UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu)
                    End If
                    'Lo sacamos.
                    If UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).flags.UserLogged Then
                        Call FinComerciarUsu(MapData(mapa, .Pos.X, .Pos.Y).UserIndex)
                        Call WriteErrorMsg(MapData(mapa, .Pos.X, .Pos.Y).UserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                        Call FlushBuffer(MapData(mapa, .Pos.X, .Pos.Y).UserIndex)
                    End If
                End If
                
                Call CloseSocket(MapData(mapa, .Pos.X, .Pos.Y).UserIndex)
            End If
        End If
    End If
    
    'Nombre de sistema
    .name = name
    
    .showName = True 'Por default los nombres son visibles
    Call Mod_AntiCheat.SetIntervalos(UserIndex)
    
    
   'If in the water, and has a boat, equip it!
    If .Invent.BarcoObjIndex > 0 And _
            (HayAgua(mapa, .Pos.X, .Pos.Y) Or BodyIsBoat(.Char.body)) Then

        .Char.Head = 0
        If .flags.Muerto = 0 Then
            Call ToggleBoatBody(UserIndex)
        Else
            .Char.body = iFragataFantasmal
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
        End If
        
        .flags.Navegando = 1
    End If
    
    
    
    'Info
    Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
     Call WriteChangeMap(UserIndex, .Pos.map, MapInfo(.Pos.map).MapVersion)
    Call WritePlayMidi(UserIndex, val(ReadField(1, MapInfo(.Pos.map).Music, 45)))
    
    If .flags.Privilegios = PlayerType.Dios Then
        .flags.ChatColor = RGB(255, 255, 255)
    ElseIf .flags.Privilegios <> PlayerType.User And .flags.Privilegios <> (PlayerType.User Or PlayerType.ChaosCouncil) And .flags.Privilegios <> (PlayerType.User Or PlayerType.RoyalCouncil) Then
        .flags.ChatColor = RGB(255, 255, 255)
    ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.RoyalCouncil) Then
        .flags.ChatColor = RGB(255, 255, 255)
    ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.ChaosCouncil) Then
        .flags.ChatColor = RGB(255, 255, 255)
    Else
        .flags.ChatColor = vbWhite
    End If
    
    .flags.ModoCombate = False
    .flags.ModoCombate = 0
    
    ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
    #If ConUpTime Then
        .LogOnTime = Now
    #End If
    'Crea  el personaje del usuario
    Call MakeUserChar(True, .Pos.map, UserIndex, .Pos.map, .Pos.X, .Pos.Y)
    
    If (.flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster)) = 0 Then
     '   Call DoAdminInvisible(Userindex)
    End If
    
    Call WriteUserCharIndexInServer(UserIndex)
    ''[/el oso]
    
    Call DoTileEvents(UserIndex, .Pos.map, .Pos.X, .Pos.Y)
    
    Call CheckUserLevel(UserIndex)
    Call WriteUpdateUserStats(UserIndex)
    
    Call WriteUpdateHungerAndThirst(UserIndex)
    Call WriteUpdateStrenghtAndDexterity(UserIndex)
    
    If haciendoBK Then
        Call WritePauseToggle(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Servidor> Por favor espera algunos segundos, el WorldSave está ejecutándose.", FontTypeNames.FONTTYPE_SERVER)
    End If
    
    If EnPausa Then
        Call WritePauseToggle(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER)
    End If
    
    If EnTesting And .Stats.ELV >= 18 Then
        Call WriteErrorMsg(UserIndex, "Servidor en Testing por unos minutos, conectese con PJs de nivel menor a 18. No se conecte con Pjs que puedan resultar importantes por ahora pues pueden arruinarse.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    If EsGM(UserIndex) Then
        DoAdminInvisible UserIndex ' @@ GM Invisible al conectarse
    End If
    
    NumUsers = NumUsers + 1
    .flags.UserLogged = True
    
    'usado para borrar Pjs
    Call WriteVar(CharPath & .name & ".chr", "INIT", "Logged", "1")
    
    Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    
    MapInfo(.Pos.map).NumUsers = MapInfo(.Pos.map).NumUsers + 1
    
    If .Stats.SkillPts > 0 Then
        Call WriteSendSkills(UserIndex)
        Call WriteLevelUp(UserIndex, .Stats.SkillPts)
    End If
    
    If NumUsers > recordusuarios Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("RECORD de usuarios conectados simultaneamente." & "Hay " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_INFO))
        recordusuarios = NumUsers
        Call WriteVar(IniPath & "Server.ini", "INIT", "RECORD", Str(recordusuarios))
        
        Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)
    End If
    
    If .NroMascotas > 0 And MapInfo(.Pos.map).Pk Then
        Dim i As Integer
        For i = 1 To MAXMASCOTAS
            If .MascotasType(i) > 0 Then
                .MascotasIndex(i) = SpawnNpc(.MascotasType(i), .Pos, True, True)
                
                If .MascotasIndex(i) > 0 Then
                    Npclist(.MascotasIndex(i)).MaestroUser = UserIndex
                    Call FollowAmo(.MascotasIndex(i))
                Else
                    .MascotasIndex(i) = 0
                End If
            End If
        Next i
    End If
    
    If .flags.Navegando = 1 Then
        Call WriteNavigateToggle(UserIndex)
    End If
    
     If .flags.Montando = 1 Then
       Call WriteMontateToggle(UserIndex)
    End If
    
    If .ACT Then
    WriteMultiMessage UserIndex, eMessages.DragOnn
    .ACT = True
    Else
    WriteMultiMessage UserIndex, eMessages.DragOff
    .ACT = False
    End If

    If criminal(UserIndex) Then
        Call WriteMultiMessage(UserIndex, eMessages.SafeModeOff) 'Call WriteSafeModeOff(UserIndex)
        .flags.Seguro = False
    Else
        .flags.Seguro = True
        Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn) 'Call WriteSafeModeOn(UserIndex)
    End If
    
    If .GuildIndex > 0 Then
        'welcome to the show baby...
        If Not modGuilds.m_ConectarMiembroAClan(UserIndex, .GuildIndex) Then
            Call WriteConsoleMsg(UserIndex, "Tu estado no te permite entrar al clan.", FontTypeNames.FONTTYPE_GUILD)
        End If
    End If
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
    
    Call WriteLoggedMessage(UserIndex)
    
    Call modGuilds.SendGuildNews(UserIndex)
    
    ' Esta protegido del ataque de npcs por 5 segundos, si no realiza ninguna accion
    Call IntervaloPermiteSerAtacado(UserIndex, True)
    
    If Lloviendo Then
        Call WriteRainToggle(UserIndex)
    End If
    
    tStr = modGuilds.a_ObtenerRechazoDeChar(.name)
    
    If LenB(tStr) <> 0 Then
        Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
    End If
    
    CheckRankingUser UserIndex, TopFrags
    CheckRankingUser UserIndex, TopLevel
    CheckRankingUser UserIndex, TopOro
    CheckRankingUser UserIndex, TopRetos
    CheckRankingUser UserIndex, TopTorneos
    CheckRankingUser UserIndex, TopClanes
    
    'Load the user statistics
    Call Statistics.UserConnected(UserIndex)
       
    
    Call MostrarNumUsers
    ActualizarAuras UserIndex
    
    
    #If SeguridadAlkon Then
        Call Security.UserConnected(UserIndex)
    #End If

    N = FreeFile
    Open App.Path & "\logs\numusers.log" For Output As N
    Print #N, NumUsers
    Close #N
    
    N = FreeFile
    'Log
    Open App.Path & "\logs\Connect.log" For Append Shared As #N
    Print #N, .name & " ha entrado al juego. UserIndex:" & UserIndex & " " & time & " " & Date
    Close #N

     If .flags.ModoCombate = False Then
        WriteConsoleMsg UserIndex, "No estás en modo combate, para activarlo presiona la tecla C.", FontTypeNames.FONTTYPE_INFO
    End If
    
    'Le informo que tiene el pj en mercado
    Dim CharFile$, GG As String
    CharFile = CharPath & ".chr"
    
    GG = val(GetVar(CharFile, "VENTA", "iVenta"))
    
    If .flags.EstaEnMercado = True And GG = "1" Then
        WriteConsoleMsg UserIndex, "Recuerda que si tu personaje está en el mercado, para retirarlo teclea /QUITARPJ, para mayor información ve a opciones", FontTypeNames.fonttype_dios
    End If
     
End With
End Sub
Sub ResetFaccionCaos(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'*************************************************
    With UserList(UserIndex).Faccion
        .ArmadaReal = 0
        .FuerzasCaos = 0
        .FechaIngreso = "No ingresó a ninguna Facción"
        .CriminalesMatados = 0
        .RecibioArmaduraCaos = 0
        .RecibioArmaduraReal = 0
        .RecibioExpInicialCaos = 0
        .RecibioExpInicialReal = 0
        .RecompensasCaos = 0
        .RecompensasReal = 0
        .Reenlistadas = 0
        .NivelIngreso = 0
        .MatadosIngreso = 0
        .NextRecompensa = 0
    End With
End Sub
Sub ResetFaccionReal(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'*************************************************
    With UserList(UserIndex).Faccion
        .ArmadaReal = 0
        .CiudadanosMatados = 0
        .FuerzasCaos = 0
        .FechaIngreso = "No ingresó a ninguna Facción"
        .RecibioArmaduraCaos = 0
        .RecibioArmaduraReal = 0
        .RecibioExpInicialCaos = 0
        .RecibioExpInicialReal = 0
        .RecompensasCaos = 0
        .RecompensasReal = 0
        .Reenlistadas = 0
        .NivelIngreso = 0
        .MatadosIngreso = 0
        .NextRecompensa = 0
    End With
End Sub

Sub ResetContadores(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'05/20/2007 Integer - Agregue todas las variables que faltaban.
'*************************************************
    With UserList(UserIndex).Counters
        .AGUACounter = 0
        .AttackCounter = 0
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Paralisis = 0
        .Pena = 0
        .PiqueteC = 0
        .STACounter = 0
        .Veneno = 0
        .Trabajando = 0
        .Ocultando = 0
        '.bPuedeMeditar = False
        .Lava = 0
        .Mimetismo = 0
        .Saliendo = False
        .Salir = 0
        .TiempoOculto = 0
        .TimerMagiaGolpe = 0
        .TimerGolpeMagia = 0
        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeUsarArco = 0
        .TimerPuedeTrabajar = 0
        .TimerUsar = 0
        .TimerUsarClick = 0
        .failedUsageAttempts = 0
        .goHome = 0
        .AsignedSkills = 0
    End With
End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex).Char
        .body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .loops = 0
        .Heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex)
        .name = vbNullString
        .desc = vbNullString
        .DescRM = vbNullString
        .Pos.map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .ip = vbNullString
        .clase = 0
        .Email = vbNullString
        .Genero = 0
        .Hogar = 0
        .Raza = 0
        
        .PartyIndex = 0
        .PartySolicitud = 0
        
        With .Stats
            .Banco = 0
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .def = 0
            '.CriminalesMatados = 0
            .NPCsMuertos = 0
            .UsuariosMatados = 0
            .SkillPts = 0
            .GLD = 0
            .UserAtributos(1) = 0
            .UserAtributos(2) = 0
            .UserAtributos(3) = 0
            .UserAtributos(4) = 0
            .UserAtributos(5) = 0
            .UserAtributosBackUP(1) = 0
            .UserAtributosBackUP(2) = 0
            .UserAtributosBackUP(3) = 0
            .UserAtributosBackUP(4) = 0
            .UserAtributosBackUP(5) = 0
        End With
        
    End With
End Sub

Sub ResetReputacion(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex).Reputacion
        .AsesinoRep = 0
        .BandidoRep = 0
        .BurguesRep = 0
        .LadronesRep = 0
        .NobleRep = 0
        .PlebeRep = 0
        .NobleRep = 0
        .Promedio = 0
    End With
End Sub

Sub ResetGuildInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If UserList(UserIndex).EscucheClan > 0 Then
        Call modGuilds.GMDejaDeEscucharClan(UserIndex, UserList(UserIndex).EscucheClan)
        UserList(UserIndex).EscucheClan = 0
    End If
    If UserList(UserIndex).GuildIndex > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(UserIndex, UserList(UserIndex).GuildIndex)
    End If
    UserList(UserIndex).GuildIndex = 0
End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 06/28/2008
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'03/29/2006 Maraxus - Reseteo el CentinelaOK también.
'06/28/2008 NicoNZ - Agrego el flag Inmovilizado
'*************************************
'************
    
    With UserList(UserIndex).flags
    .ElPedidorSeguimiento = 0
    .indexmercado = 0
    UserList(UserIndex).elpedidor = 0
            .MenuCliente = 0 'Esto de MenuCliente es de coso, Hispano i guess.
            .RandomKey = 0
        .LastSlotClient = 0
        .Siguiendo = 0
        .EnEvento = 0
.Compañero = 0
        .Comerciando = False
        .Ban = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .NpcInv = 0
        .StatsChanged = 0
        .TargetNPC = 0
        .TargetNpcTipo = eNPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Descuento = vbNullString
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .ModoCombate = False
        .Vuela = 0
        .Navegando = 0
        .Montando = 0
        .Oculto = 0
        .Envenenado = 0
        .invisible = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = 0
        .Privilegios = 0
        .PrivEspecial = False
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .ValCoDe = 0
        .Hechizo = 0
        .TimesWalk = 0
        .StartWalk = 0
        .CountSH = 0
        .Silenciado = 0
        .CentinelaOK = False
        .AdminPerseguible = False
        .lastMap = 0
        .Traveling = 0
        .AtacablePor = 0
        .AtacadoPorNpc = 0
        .AtacadoPorUser = 0
        .NoPuedeSerAtacado = False
        .OwnedNpc = 0
        .ShareNpcWith = 0
        .EnConsulta = False
        .Ignorado = False
    End With
End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim LoopC As Long
    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(UserIndex).Stats.UserHechizos(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserPets(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim LoopC As Long
    
    UserList(UserIndex).NroMascotas = 0
        
    For LoopC = 1 To MAXMASCOTAS
        UserList(UserIndex).MascotasIndex(LoopC) = 0
        UserList(UserIndex).MascotasType(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim LoopC As Long
    
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
          UserList(UserIndex).BancoInvent.Object(LoopC).Amount = 0
          UserList(UserIndex).BancoInvent.Object(LoopC).Equipped = 0
          UserList(UserIndex).BancoInvent.Object(LoopC).objindex = 0
    Next LoopC
    
    UserList(UserIndex).BancoInvent.NroItems = 0
End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex).ComUsu
        If .DestUsu > 0 Then
            Call FinComerciarUsu(.DestUsu)
            Call FinComerciarUsu(UserIndex)
        End If
    End With
End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim i As Long

UserList(UserIndex).ConnIDValida = False
UserList(UserIndex).ConnID = -1

Call LimpiarComercioSeguro(UserIndex)
Call ResetFaccionCaos(UserIndex)
Call ResetFaccionReal(UserIndex)
Call ResetContadores(UserIndex)
Call ResetGuildInfo(UserIndex)
Call ResetCharInfo(UserIndex)
Call ResetBasicUserInfo(UserIndex)
Call ResetReputacion(UserIndex)
Call ResetUserFlags(UserIndex)
Call ResetFlagsMercado(UserIndex)
Call LimpiarInventario(UserIndex)
Call ResetUserSpells(UserIndex)
Call ResetUserPets(UserIndex)
Call ResetUserBanco(UserIndex)
Call ResetQuestStats(UserIndex)
With UserList(UserIndex).ComUsu
    .Acepto = False
    
    For i = 1 To MAX_OFFER_SLOTS
        .cant(i) = 0
        .Objeto(i) = 0
    Next i
    
    .goldAmount = 0
    .DestNick = vbNullString
    .DestUsu = 0
End With
 
End Sub

Sub CloseUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

Dim N As Integer
Dim map As Integer
Dim name As String
Dim i As Integer

Dim aN As Integer

With UserList(UserIndex)
    aN = .flags.AtacadoPorNpc
    If aN > 0 Then
          Npclist(aN).Movement = Npclist(aN).flags.OldMovement
          Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
          Npclist(aN).flags.AttackedBy = vbNullString
    End If
    
    aN = .flags.NPCAtacado
    If aN > 0 Then
        If Npclist(aN).flags.AttackedFirstBy = .name Then
            Npclist(aN).flags.AttackedFirstBy = vbNullString
        End If
    End If
    .flags.AtacadoPorNpc = 0
    .flags.NPCAtacado = 0
    
    map = .Pos.map
    name = UCase$(.name)
    
    .Char.FX = 0
    .Char.loops = 0
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
    
    .flags.UserLogged = False
    .Counters.Saliendo = False
    
    'Le devolvemos el body y head originales
    If .flags.AdminInvisible = 1 Then
        .Char.body = .flags.OldBody
        .Char.Head = .flags.OldHead
        .flags.AdminInvisible = 0
    End If
    
    'si esta en party le devolvemos la experiencia
    If .PartyIndex > 0 Then Call mdParty.SalirDeParty(UserIndex)
    
    'Save statistics
    Call Statistics.UserDisconnected(UserIndex)
    
    ' Grabamos el personaje del usuario
    Call SaveUser(UserIndex, CharPath & name & ".chr")
    
    'usado para borrar Pjs
    Call WriteVar(CharPath & .name & ".chr", "INIT", "Logged", "0")

    
    'Quitar el dialogo
    'If MapInfo(Map).NumUsers > 0 Then
    '    Call SendToUserArea(UserIndex, "QDL" & .Char.charindex)
    'End If
    
    If MapInfo(map).NumUsers > 0 Then
        Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
    End If
    
    'Borrar el personaje
    If .Char.CharIndex > 0 Then
        Call EraseUserChar(UserIndex, .flags.AdminInvisible = 1)
    End If
    
    'Borrar mascotas
    For i = 1 To MAXMASCOTAS
        If .MascotasIndex(i) > 0 Then
            If Npclist(.MascotasIndex(i)).flags.NPCActive Then _
                Call QuitarNPC(.MascotasIndex(i))
        End If
    Next i
    
    'Update Map Users
    MapInfo(map).NumUsers = MapInfo(map).NumUsers - 1
    
    If MapInfo(map).NumUsers < 0 Then
        MapInfo(map).NumUsers = 0
    End If
    
    ' Si el usuario habia dejado un msg en la gm's queue lo borramos
    If Ayuda.Existe(.name) Then Call Ayuda.Quitar(.name)
    
    Call ResetUserSlot(UserIndex)
    
    Call MostrarNumUsers
    
    N = FreeFile(1)
    Open App.Path & "\logs\Connect.log" For Append Shared As #N
        Print #N, name & " ha dejado el juego. " & "User Index:" & UserIndex & " " & time & " " & Date
    Close #N
End With

Exit Sub

errhandler:
Call LogError("Error en CloseUser. Número " & Err.Number & " Descripción: " & Err.description)

End Sub

Sub ReloadSokcet()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler
#If UsarQueSocket = 1 Then

    Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
    If NumUsers <= 0 Then
        Call WSApiReiniciarSockets
    Else
'       Call apiclosesocket(SockListen)
'       SockListen = ListenForConnect(Puerto, hWndMsg, "")
    End If

#ElseIf UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    Call ConfigListeningSocket(frmMain.Socket1, Puerto)
    
#ElseIf UsarQueSocket = 2 Then

    

#End If

Exit Sub
errhandler:
    Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.description)

End Sub

Public Sub EnviarNoche(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Call WriteSendNight(UserIndex, IIf(DeNoche And (MapInfo(UserList(UserIndex).Pos.map).Zona = Campo Or MapInfo(UserList(UserIndex).Pos.map).Zona = Ciudad), True, False))
    Call WriteSendNight(UserIndex, IIf(DeNoche, True, False))
End Sub

Public Sub EcharPjsNoPrivilegiados()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim LoopC As Long
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
            If UserList(LoopC).flags.Privilegios And PlayerType.User Then
                Call CloseSocket(LoopC)
            End If
        End If
    Next LoopC

End Sub
Sub ResetFlagsMercado(ByVal UserIndex As Integer)

    ' @ Reseteamos el mercado interno del char.
    Dim i As Integer
    
    For i = 1 To 10
        UserList(UserIndex).Ofertas(i).OfertasHechas = vbNullString
        UserList(UserIndex).Ofertas(i).OfertasRecibidas = vbNullString
    Next i
End Sub
