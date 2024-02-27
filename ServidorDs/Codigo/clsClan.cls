VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsClan.cls
'
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

Option Explicit

''
' clase clan
'
' Es el "ADO" de los clanes. La interfaz entre el disco y
' el juego. Los datos no se guardan en memoria
' para evitar problemas de sincronizacion, y considerando
' que la performance de estas rutinas NO es critica.
' by el oso :p

Private p_GuildName                 As String
Private p_Alineacion                As ALINEACION_GUILD
Private p_OnlineMembers             As Collection   'Array de UserIndexes!
Private p_GMsOnline                 As Collection
Private p_PropuestasDePaz           As Collection
Private p_PropuestasDeAlianza       As Collection
Private p_IteradorRelaciones        As Integer
Private p_IteradorOnlineMembers     As Integer
Private p_IteradorPropuesta         As Integer
Private p_IteradorOnlineGMs         As Integer
Private p_GuildNumber               As Integer      'Numero de guild en el mundo
Private p_Relaciones()              As RELACIONES_GUILD 'array de relaciones con los otros clanes
Private GUILDINFOFILE               As String
Private GUILDPATH                   As String       'aca pq me es mas comodo setearlo y pq en ningun disenio
Private MEMBERSFILE                 As String       'decente la capa de arriba se entera donde estan
Private SOLICITUDESFILE             As String       'los datos fisicamente
Private PROPUESTASFILE              As String
Private RELACIONESFILE              As String
Private VOTACIONESFILE              As String

Private Const NEWSLENGTH = 1024
Private Const DESCLENGTH = 256
Private Const CODEXLENGTH = 256

Public Property Get GuildName() As String
10        GuildName = p_GuildName
End Property


'
'ALINEACION Y ANTIFACCION
'

Public Property Get Alineacion() As ALINEACION_GUILD
10        Alineacion = p_Alineacion
End Property

Public Property Get PuntosAntifaccion() As Integer
10        PuntosAntifaccion = val(GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Antifaccion"))
End Property

Public Property Let PuntosAntifaccion(ByVal p As Integer)
10        Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Antifaccion", CStr(p))
End Property

Public Function CambiarAlineacion(ByVal NuevaAlineacion As ALINEACION_GUILD) As Boolean
10        p_Alineacion = NuevaAlineacion
20        Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Alineacion", Alineacion2String(p_Alineacion))
          
30        If p_Alineacion = ALINEACION_NEUTRO Then CambiarAlineacion = True
End Function

'
'INICIALIZADORES
'

Private Sub Class_Initialize()
10        GUILDPATH = App.Path & "\GUILDS\"
20        GUILDINFOFILE = GUILDPATH & "guildsinfo.inf"
End Sub

Private Sub Class_Terminate()
10        Set p_OnlineMembers = Nothing
20        Set p_GMsOnline = Nothing
30        Set p_PropuestasDePaz = Nothing
40        Set p_PropuestasDeAlianza = Nothing
End Sub



Public Sub Inicializar(ByVal GuildName As String, ByVal GuildNumber As Integer, ByVal Alineacion As ALINEACION_GUILD)
      Dim i As Integer

10        p_GuildName = GuildName
20        p_GuildNumber = GuildNumber
30        p_Alineacion = Alineacion
40        Set p_OnlineMembers = New Collection
50        Set p_GMsOnline = New Collection
60        Set p_PropuestasDePaz = New Collection
70        Set p_PropuestasDeAlianza = New Collection
          'ALLIESFILE = GUILDPATH & p_GuildName & "-Allied.all"
          'ENEMIESFILE = GUILDPATH & p_GuildName & "-enemys.ene"
80        RELACIONESFILE = GUILDPATH & p_GuildName & "-relaciones.rel"
90        MEMBERSFILE = GUILDPATH & p_GuildName & "-members.mem"
100       PROPUESTASFILE = GUILDPATH & p_GuildName & "-propositions.pro"
110       SOLICITUDESFILE = GUILDPATH & p_GuildName & "-solicitudes.sol"
120       VOTACIONESFILE = GUILDPATH & p_GuildName & "-votaciones.vot"
130       p_IteradorOnlineMembers = 0
140       p_IteradorPropuesta = 0
150       p_IteradorOnlineGMs = 0
160       p_IteradorRelaciones = 0
170       ReDim Preserve p_Relaciones(1 To CANTIDADDECLANES) As RELACIONES_GUILD
180       For i = 1 To CANTIDADDECLANES
190           p_Relaciones(i) = String2Relacion(GetVar(RELACIONESFILE, "RELACIONES", CStr(i)))
200       Next i
210       For i = 1 To CANTIDADDECLANES
220           If Trim$(GetVar(PROPUESTASFILE, CStr(i), "Pendiente")) = "1" Then
230               Select Case String2Relacion(Trim$(GetVar(PROPUESTASFILE, CStr(i), "Tipo")))
                      Case RELACIONES_GUILD.ALIADOS
240                       p_PropuestasDeAlianza.Add i
250                   Case RELACIONES_GUILD.PAZ
260                       p_PropuestasDePaz.Add i
270               End Select
280           End If
290       Next i
End Sub

''
' esta TIENE QUE LLAMARSE LUEGO DE INICIALIZAR()
'
' @param Fundador Nombre del fundador del clan
'
Public Sub InicializarNuevoClan(ByRef Fundador As String)
      Dim OldQ    As String  'string pq al comienzo quizas no hay archivo guildinfo.ini y oldq es ""
      Dim NewQ    As Integer
          'para que genere los archivos
10        Call WriteVar(MEMBERSFILE, "INIT", "NroMembers", "0")
20        Call WriteVar(SOLICITUDESFILE, "INIT", "CantSolicitudes", "0")


30        OldQ = GetVar(GUILDINFOFILE, "INIT", "nroguilds")
40        If IsNumeric(OldQ) Then
50            NewQ = CInt(Trim$(OldQ)) + 1
60        Else
70            NewQ = 1
80        End If

90        Call WriteVar(GUILDINFOFILE, "INIT", "NroGuilds", NewQ)

100       Call WriteVar(GUILDINFOFILE, "GUILD" & NewQ, "Founder", Fundador)
110       Call WriteVar(GUILDINFOFILE, "GUILD" & NewQ, "GuildName", p_GuildName)
120       Call WriteVar(GUILDINFOFILE, "GUILD" & NewQ, "Date", Date)
130       Call WriteVar(GUILDINFOFILE, "GUILD" & NewQ, "Antifaccion", "0")
140       Call WriteVar(GUILDINFOFILE, "GUILD" & NewQ, "Alineacion", Alineacion2String(p_Alineacion))

End Sub

Public Sub ProcesarFundacionDeOtroClan()
10        ReDim Preserve p_Relaciones(1 To CANTIDADDECLANES) As RELACIONES_GUILD
20        p_Relaciones(CANTIDADDECLANES) = PAZ
End Sub


'
'MEMBRESIAS
'

Public Property Get Fundador() As String
10        Fundador = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Founder")
End Property

'Public Property Get JugadoresOnline() As String
'Dim i As Integer
'    'leve violacion de capas x aqui, je
'    For i = 1 To p_OnlineMembers.Count
'        JugadoresOnline = UserList(p_OnlineMembers.Item(i)).Name & "," & JugadoresOnline
'    Next i
'End Property

Public Property Get CantidadDeMiembros() As Integer
      Dim OldQ As String
10        OldQ = GetVar(MEMBERSFILE, "INIT", "NroMembers")
20        CantidadDeMiembros = IIf(IsNumeric(OldQ), CInt(OldQ), 0)
End Property

Public Sub SetLeader(ByRef leader As String)
10        Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Leader", leader)
End Sub

Public Function GetLeader() As String
10        GetLeader = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Leader")
End Function

Public Function GetMemberList() As String()
          Dim OldQ    As Integer
          Dim list() As String
          Dim i As Long
          
10        OldQ = Me.CantidadDeMiembros
          
20        If OldQ = 0 Then Exit Function
30        ReDim list(OldQ - 1) As String
          
40        For i = 1 To OldQ
50            list(i - 1) = UCase$(GetVar(MEMBERSFILE, "Members", "Member" & i))
60        Next i
          
70        GetMemberList = list
End Function

Public Sub ConectarMiembro(ByVal Userindex As Integer)
10        p_OnlineMembers.Add Userindex

20        With UserList(Userindex)
30            Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & " se ha conectado."))
40        End With
End Sub

Public Sub DesConectarMiembro(ByVal Userindex As Integer)
      Dim i   As Integer
10        For i = 1 To p_OnlineMembers.Count
20            If p_OnlineMembers.Item(i) = Userindex Then
30                p_OnlineMembers.Remove i
                  
40                With UserList(Userindex)
50                    Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & " se ha desconectado."))
60                End With

70                Exit Sub
80            End If
90        Next i
          

End Sub

Public Sub AceptarNuevoMiembro(ByRef Nombre As String)
      Dim OldQ    As Integer
      Dim OldQs   As String
      Dim ruta    As String

10    ruta = CharPath & Nombre & ".chr"
20        If FileExist(ruta) Then
30            Call WriteVar(ruta, "GUILD", "GUILDINDEX", p_GuildNumber)
40            Call WriteVar(ruta, "GUILD", "AspiranteA", "0")
              'CantPs = GetVar(CharPath & Nombre & ".chr", "GUILD", "ClanesParticipo")
              'If IsNumeric(CantPs) Then
              '    CantP = CInt(CantPs)
              'Else
              '    CantP = 0
              'End If
              'Call WriteVar(CharPath & Nombre & ".chr", "GUILD", "ClanesParticipo", CantP + 1)
50            OldQs = GetVar(MEMBERSFILE, "INIT", "NroMembers")
60            If IsNumeric(OldQs) Then
70                OldQ = CInt(OldQs)
80            Else
90                OldQ = 0
100           End If
110           Call WriteVar(MEMBERSFILE, "INIT", "NroMembers", OldQ + 1)
120           Call WriteVar(MEMBERSFILE, "Members", "Member" & OldQ + 1, Nombre)
130       End If

End Sub

Public Sub ExpulsarMiembro(ByRef Nombre As String)
      Dim OldQ        As Integer
      Dim Temps       As String
      Dim i           As Integer
      Dim EsMiembro   As Boolean
      Dim MiembroDe   As String

10        If LenB(dir$(CharPath & Nombre & ".chr")) <> 0 Then
20            OldQ = CInt(GetVar(MEMBERSFILE, "INIT", "NroMembers"))
30            i = 1
40            Nombre = UCase$(Nombre)
50            While i <= OldQ And UCase$(Trim$(GetVar(MEMBERSFILE, "Members", "Member" & i))) <> Nombre
60                i = i + 1
70            Wend
80            EsMiembro = i <= OldQ
              
90            If EsMiembro Then
100               Call WriteVar(CharPath & Nombre & ".chr", "GUILD", "GuildIndex", vbNullString)
110               While i < OldQ
120                   Temps = GetVar(MEMBERSFILE, "Members", "Member" & i + 1)
130                   Call WriteVar(MEMBERSFILE, "Members", "Member" & i, Temps)
140                   i = i + 1
150               Wend
160               Call WriteVar(MEMBERSFILE, "Members", "Member" & OldQ, vbNullString)
                  'seteo la cantidad de miembros nueva
170               Call WriteVar(MEMBERSFILE, "INIT", "NroMembers", OldQ - 1)
                  'lo echo a el
180               MiembroDe = GetVar(CharPath & Nombre & ".chr", "GUILD", "Miembro")
190               If Not InStr(1, MiembroDe, p_GuildName, vbTextCompare) > 0 Then
200                   If LenB(MiembroDe) <> 0 Then
210                       MiembroDe = MiembroDe & ","
220                   End If
230                   MiembroDe = MiembroDe & p_GuildName
240                   Call WriteVar(CharPath & Nombre & ".chr", "GUILD", "Miembro", MiembroDe)
250               End If
260           End If

270       End If

End Sub

'
'ASPIRANTES
'

Public Function GetAspirantes() As String()
          Dim OldQ    As Integer
          Dim list() As String
          Dim i As Long
          
10        OldQ = Me.CantidadAspirantes()
          
20        If OldQ > 1 Then
30            ReDim list(OldQ - 1) As String
40        Else
50            ReDim list(0) As String
60        End If
          
70        For i = 1 To OldQ
80            list(i - 1) = GetVar(SOLICITUDESFILE, "SOLICITUD" & i, "Nombre")
90        Next i
          
100       GetAspirantes = list
End Function

Public Function CantidadAspirantes() As Integer
      Dim Temps As String

10        CantidadAspirantes = 0
20        Temps = GetVar(SOLICITUDESFILE, "INIT", "CantSolicitudes")
30        If Not IsNumeric(Temps) Then
40            Exit Function
50        End If
60        CantidadAspirantes = CInt(Temps)

End Function

Public Function DetallesSolicitudAspirante(ByVal NroAspirante As Integer) As String
10        DetallesSolicitudAspirante = GetVar(SOLICITUDESFILE, "SOLICITUD" & NroAspirante, "Detalle")
End Function

Public Function NumeroDeAspirante(ByRef Nombre As String) As Integer
      Dim i       As Integer
          
10        NumeroDeAspirante = 0
          
20        For i = 1 To MAXASPIRANTES
30            If UCase$(Trim$(GetVar(SOLICITUDESFILE, "SOLICITUD" & i, "Nombre"))) = UCase$(Nombre) Then
40                NumeroDeAspirante = i
50                Exit Function
60            End If
70        Next i
End Function

Public Sub NuevoAspirante(ByRef Nombre As String, ByRef Peticion As String)
      Dim i       As Integer
      Dim OldQ    As String
      Dim OldQI   As Integer

10        OldQ = GetVar(SOLICITUDESFILE, "INIT", "CantSolicitudes")
20        If IsNumeric(OldQ) Then
30            OldQI = CInt(OldQ)
40        Else
50            OldQI = 0
60        End If
70        For i = 1 To MAXASPIRANTES
80            If GetVar(SOLICITUDESFILE, "SOLICITUD" & i, "Nombre") = vbNullString Then
90                Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & i, "Nombre", Nombre)
100               Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & i, "Detalle", IIf(Trim$(Peticion) = vbNullString, "Peticion vacia", Peticion))
110               Call WriteVar(SOLICITUDESFILE, "INIT", "CantSolicitudes", OldQI + 1)
120               Call WriteVar(CharPath & Nombre & ".chr", "GUILD", "ASPIRANTEA", p_GuildNumber)
130               Exit Sub
140           End If
150       Next i
End Sub

Public Sub RetirarAspirante(ByRef Nombre As String, ByRef NroAspirante As Integer)
      Dim OldQ    As String
      Dim OldQI   As String
      Dim Pedidos As String
      Dim i       As Integer

10        OldQ = GetVar(SOLICITUDESFILE, "INIT", "CantSolicitudes")
20        If IsNumeric(OldQ) Then
30            OldQI = CInt(OldQ)
40        Else
50            OldQI = 1
60        End If
          'Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & NroAspirante, "Nombre", vbNullString)
          'Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & NroAspirante, "Detalle", vbNullString)
70        Call WriteVar(CharPath & Nombre & ".chr", "GUILD", "ASPIRANTEA", "0")
80        Pedidos = GetVar(CharPath & Nombre & ".chr", "GUILD", "Pedidos")
90        If Not InStr(1, Pedidos, p_GuildName, vbTextCompare) > 0 Then
100           If LenB(Pedidos) <> 0 Then
110               Pedidos = Pedidos & ","
120           End If
130           Pedidos = Pedidos & p_GuildName
140           Call WriteVar(CharPath & Nombre & ".chr", "GUILD", "Pedidos", Pedidos)
150       End If
          
160       Call WriteVar(SOLICITUDESFILE, "INIT", "CantSolicitudes", OldQI - 1)
170       For i = NroAspirante To modGuilds.MAXASPIRANTES - 1
180           Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & i, "Nombre", GetVar(SOLICITUDESFILE, "SOLICITUD" & (i + 1), "Nombre"))
190           Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & i, "Detalle", GetVar(SOLICITUDESFILE, "SOLICITUD" & (i + 1), "Detalle"))
200       Next i
          
210       Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & modGuilds.MAXASPIRANTES, "Nombre", vbNullString)
220       Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & modGuilds.MAXASPIRANTES, "Detalle", vbNullString)
          
End Sub

Public Sub InformarRechazoEnChar(ByRef Nombre As String, ByRef Detalles As String)
10        Call WriteVar(CharPath & Nombre & ".chr", "GUILD", "MotivoRechazo", Detalles)
End Sub

'
'DEFINICION DEL CLAN (CODEX Y NOTICIAS)
'

Public Function GetFechaFundacion() As String
10        GetFechaFundacion = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Date")
End Function

Public Sub SetCodex(ByVal CodexNumber As Integer, ByRef codex As String)
10        Call ReplaceInvalidChars(codex)
20        codex = Left$(codex, CODEXLENGTH)
30        Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Codex" & CodexNumber, codex)
End Sub

Public Function GetCodex(ByVal CodexNumber As Integer) As String
10        GetCodex = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Codex" & CodexNumber)
End Function


Public Sub SetURL(ByRef URL As String)
10        Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "URL", Left$(URL, 40))
End Sub

Public Function GetURL() As String
10        GetURL = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "URL")
End Function

Public Sub SetGuildNews(ByRef News As String)
10        Call ReplaceInvalidChars(News)
          
20        News = Left$(News, NEWSLENGTH)
          
30        Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "GuildNews", News)
End Sub

Public Function GetGuildNews() As String
10        GetGuildNews = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "GuildNews")
End Function

Public Sub SetDesc(ByRef desc As String)
10        Call ReplaceInvalidChars(desc)
20        desc = Left$(desc, DESCLENGTH)
          
30        Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Desc", desc)
End Sub

Public Function GetDesc() As String
10        GetDesc = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Desc")
End Function

'
'
'ELECCIONES
'
'

Public Function EleccionesAbiertas() As Boolean
      Dim ee As String
10        ee = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesAbiertas")
20        EleccionesAbiertas = (ee = "1")     'cualquier otra cosa da falso
End Function

Public Sub AbrirElecciones()
10        Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesAbiertas", "1")
20        Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesFinalizan", DateAdd("d", 1, Now))
30        Call WriteVar(VOTACIONESFILE, "INIT", "NumVotos", "0")
End Sub

Private Sub CerrarElecciones()  'solo pueden cerrarse mediante recuento de votos
10        Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesAbiertas", "0")
20        Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesFinalizan", vbNullString)
30        Call Kill(VOTACIONESFILE)   'borramos toda la evidencia ;-)
End Sub

Public Sub ContabilizarVoto(ByRef Votante As String, ByRef Votado As String)
      Dim q       As Integer
      Dim Temps   As String

10        Temps = GetVar(VOTACIONESFILE, "INIT", "NumVotos")
20        q = IIf(IsNumeric(Temps), CInt(Temps), 0)
30        Call WriteVar(VOTACIONESFILE, "VOTOS", Votante, Votado)
40        Call WriteVar(VOTACIONESFILE, "INIT", "NumVotos", CStr(q + 1))
End Sub

Public Function YaVoto(ByRef Votante) As Boolean
10        YaVoto = ((LenB(Trim$(GetVar(VOTACIONESFILE, "VOTOS", Votante)))) <> 0)
End Function

Private Function ContarVotos(ByRef CantGanadores As Integer) As String
      Dim q           As Integer
      Dim i           As Integer
      Dim Temps       As String
      Dim tempV       As String
      Dim d           As diccionario

10    On Error GoTo errh
20        ContarVotos = vbNullString
30        CantGanadores = 0
40        Temps = GetVar(MEMBERSFILE, "INIT", "NroMembers")
50        q = IIf(IsNumeric(Temps), CInt(Temps), 0)
60        If q > 0 Then
              'el diccionario tiene clave el elegido y valor la #votos
70            Set d = New diccionario
              
80            For i = 1 To q
                  'miembro del clan
90                Temps = GetVar(MEMBERSFILE, "MEMBERS", "Member" & i)
                  
                  'a quienvoto
100               tempV = GetVar(VOTACIONESFILE, "VOTOS", Temps)
                  
                  'si voto a alguien contabilizamos el voto
110               If LenB(tempV) <> 0 Then
120                   If Not IsNull(d.At(tempV)) Then  'cuantos votos tiene?
130                       Call d.AtPut(tempV, CInt(d.At(tempV)) + 1)
140                   Else
150                       Call d.AtPut(tempV, 1)
160                   End If
170               End If
180           Next i
          
              'quien quedo con mas votos, y cuantos tuvieron esos votos?
190           ContarVotos = d.MayorValor(CantGanadores)
          
200           Set d = Nothing
210       End If
          
220   Exit Function
errh:
230       LogError ("clsClan.Contarvotos: " & Err.Description)
240       If Not d Is Nothing Then Set d = Nothing
250       ContarVotos = vbNullString
End Function

Public Function RevisarElecciones() As Boolean
      Dim FechaSufragio   As Date
      Dim Temps           As String
      Dim Ganador         As String
      Dim CantGanadores   As Integer
      Dim list()          As String
      Dim i               As Long

10        RevisarElecciones = False
20        Temps = Trim$(GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesFinalizan"))
          
30        If Temps = vbNullString Then Exit Function
          
40        If IsDate(Temps) Then
50            FechaSufragio = CDate(Temps)
60            If FechaSufragio < Now Then     'toca!
70                Ganador = ContarVotos(CantGanadores)

80                If CantGanadores > 1 Then
                      'empate en la votacion
90                    Call SetGuildNews("*Empate en la votación. " & Ganador & " con " & CantGanadores & " votos ganaron las elecciones del clan.")
100               ElseIf CantGanadores = 1 Then
110                   list = Me.GetMemberList()
                      
120                   For i = 0 To UBound(list())
130                       If Ganador = list(i) Then Exit For
140                   Next i
                      
150                   If i <= UBound(list()) Then
160                       Call SetGuildNews("*" & Ganador & " ganó la elección del clan*")
170                       Call Me.SetLeader(Ganador)
180                       RevisarElecciones = True
190                   Else
200                       Call SetGuildNews("*" & Ganador & " ganó la elección del clan pero abandonó las filas por lo que la votación queda desierta*")
210                   End If
220               Else
230                   Call SetGuildNews("*El período de votación se cerró sin votos*")
240               End If
                  
250               Call CerrarElecciones
                  
260           End If
270       Else
280           Call LogError("clsClan.RevisarElecciones: tempS is not Date")
290       End If

End Function

'/VOTACIONES


'
'RELACIONES
'

Public Property Get CantidadPropuestas(ByVal Tipo As RELACIONES_GUILD) As Integer
10        Select Case Tipo
              Case RELACIONES_GUILD.ALIADOS
20                CantidadPropuestas = p_PropuestasDeAlianza.Count
30            Case RELACIONES_GUILD.GUERRA
              
40            Case RELACIONES_GUILD.PAZ
50                CantidadPropuestas = p_PropuestasDePaz.Count
60        End Select
End Property

Public Property Get CantidadEnemys() As Integer
      Dim i As Integer
10        For i = 1 To CANTIDADDECLANES
20            CantidadEnemys = CantidadEnemys + IIf(p_Relaciones(i) = GUERRA, 1, 0)
30        Next i
End Property

Public Property Get CantidadAllies() As Integer
      Dim i As Integer
10        For i = 1 To CANTIDADDECLANES
20            CantidadAllies = CantidadAllies + IIf(p_Relaciones(i) = ALIADOS, 1, 0)
30        Next i
End Property

Public Function GetRelacion(ByVal OtroGuild As Integer) As RELACIONES_GUILD
10        GetRelacion = p_Relaciones(OtroGuild)
End Function

Public Sub SetRelacion(ByVal GuildIndex As Integer, ByVal Relacion As RELACIONES_GUILD)
10        p_Relaciones(GuildIndex) = Relacion
20        Call WriteVar(RELACIONESFILE, "RELACIONES", GuildIndex, Relacion2String(Relacion))
End Sub

Public Sub SetPropuesta(ByVal Tipo As RELACIONES_GUILD, ByVal OtroGuild As Integer, ByRef Detalle As String)
10        Call WriteVar(PROPUESTASFILE, CStr(OtroGuild), "Detalle", Detalle)
20        Call WriteVar(PROPUESTASFILE, CStr(OtroGuild), "Tipo", Relacion2String(Tipo))
30        Call WriteVar(PROPUESTASFILE, CStr(OtroGuild), "Pendiente", "1")
40        Select Case Tipo
              Case RELACIONES_GUILD.ALIADOS
50                p_PropuestasDeAlianza.Add OtroGuild
60            Case RELACIONES_GUILD.PAZ
70                p_PropuestasDePaz.Add OtroGuild
80        End Select
End Sub

Public Sub AnularPropuestas(ByVal OtroGuild As Integer)
      Dim i As Integer

10        Call WriteVar(PROPUESTASFILE, CStr(OtroGuild), "Detalle", vbNullString)
20        Call WriteVar(PROPUESTASFILE, CStr(OtroGuild), "Pendiente", "0")
30        For i = 1 To p_PropuestasDePaz.Count
40            If p_PropuestasDePaz.Item(i) = OtroGuild Then p_PropuestasDePaz.Remove (i)
50            Exit Sub
60        Next i
70        For i = 1 To p_PropuestasDeAlianza.Count
80            If p_PropuestasDeAlianza.Item(i) = OtroGuild Then p_PropuestasDeAlianza.Remove (i)
90            Exit Sub
100       Next i
          
End Sub

Public Function GetPropuesta(ByVal OtroGuild As Integer, ByRef Tipo As RELACIONES_GUILD) As String
      'trae la solicitd que haya, no valida si es actual o de que tipo es
10        GetPropuesta = GetVar(PROPUESTASFILE, CStr(OtroGuild), "Detalle")
20        Tipo = String2Relacion(GetVar(PROPUESTASFILE, CStr(OtroGuild), "Tipo"))
End Function

Public Function HayPropuesta(ByVal OtroGuild As Integer, ByRef Tipo As RELACIONES_GUILD) As Boolean
      Dim i As Integer

10        HayPropuesta = False
20        Select Case Tipo
              Case RELACIONES_GUILD.ALIADOS
30                For i = 1 To p_PropuestasDeAlianza.Count
40                    If p_PropuestasDeAlianza.Item(i) = OtroGuild Then
50                        HayPropuesta = True
60                    End If
70                Next i
80            Case RELACIONES_GUILD.PAZ
90                For i = 1 To p_PropuestasDePaz.Count
100                   If p_PropuestasDePaz.Item(i) = OtroGuild Then
110                       HayPropuesta = True
120                   End If
130               Next i
140           Case RELACIONES_GUILD.GUERRA

150       End Select

End Function

'Public Function GetEnemy(ByVal EnemyIndex As Integer) As String
'    GetEnemy = GetVar(ENEMIESFILE, "ENEMYS", "ENEMY" & EnemyIndex)
'End Function

'Public Function GetAllie(ByVal AllieIndex As Integer) As String
'    GetAllie = GetVar(ALLIESFILE, "ALLIES", "ALLIE" & AllieIndex)
'End Function


'
'ITERADORES
'

Public Function Iterador_ProximaPropuesta(ByVal Tipo As RELACIONES_GUILD) As Integer
          
10        Iterador_ProximaPropuesta = 0
20        Select Case Tipo
              Case RELACIONES_GUILD.ALIADOS
30                If p_IteradorPropuesta < p_PropuestasDeAlianza.Count Then
40                    p_IteradorPropuesta = p_IteradorPropuesta + 1
50                    Iterador_ProximaPropuesta = p_PropuestasDeAlianza.Item(p_IteradorPropuesta)
60                End If
                  
70                If p_IteradorPropuesta >= p_PropuestasDeAlianza.Count Then
80                    p_IteradorPropuesta = 0
90                End If
100           Case RELACIONES_GUILD.PAZ
110               If p_IteradorPropuesta < p_PropuestasDePaz.Count Then
120                   p_IteradorPropuesta = p_IteradorPropuesta + 1
130                   Iterador_ProximaPropuesta = p_PropuestasDePaz.Item(p_IteradorPropuesta)
140               End If
                  
150               If p_IteradorPropuesta >= p_PropuestasDePaz.Count Then
160                   p_IteradorPropuesta = 0
170               End If
180       End Select
          
End Function

Public Function m_Iterador_ProximoUserIndex() As Integer
          
10        If p_IteradorOnlineMembers < p_OnlineMembers.Count Then
20            p_IteradorOnlineMembers = p_IteradorOnlineMembers + 1
30            m_Iterador_ProximoUserIndex = p_OnlineMembers.Item(p_IteradorOnlineMembers)
40        Else
50            p_IteradorOnlineMembers = 0
60            m_Iterador_ProximoUserIndex = 0
70        End If
End Function

Public Function Iterador_ProximoGM() As Integer
          
10        If p_IteradorOnlineGMs < p_GMsOnline.Count Then
20            p_IteradorOnlineGMs = p_IteradorOnlineGMs + 1
30            Iterador_ProximoGM = p_GMsOnline.Item(p_IteradorOnlineGMs)
40        Else
50            p_IteradorOnlineGMs = 0
60            Iterador_ProximoGM = 0
70        End If
End Function

Public Function Iterador_ProximaRelacion(ByVal r As RELACIONES_GUILD) As Integer
          
10        While p_IteradorRelaciones < UBound(p_Relaciones)
              
20            p_IteradorRelaciones = p_IteradorRelaciones + 1
30            If p_Relaciones(p_IteradorRelaciones) = r Then
40                Iterador_ProximaRelacion = p_IteradorRelaciones
50                Exit Function
60            End If
70        Wend

80        If p_IteradorRelaciones >= UBound(p_Relaciones) Then
90            p_IteradorRelaciones = 0
100       End If
End Function
'
'
'


'
'ADMINISTRATIVAS
'

Public Sub ConectarGM(ByVal Userindex As Integer)
10        p_GMsOnline.Add Userindex
End Sub

Public Sub DesconectarGM(ByVal Userindex As Integer)
      Dim i As Integer
10        For i = 1 To p_GMsOnline.Count
20            If p_GMsOnline.Item(i) = Userindex Then
30                p_GMsOnline.Remove (i)
40            End If
50        Next i
End Sub



'
'VARIAS, EXTRAS Y DEMASES
'

Private Sub ReplaceInvalidChars(ByRef S As String)
10        If InStrB(S, Chr$(13)) <> 0 Then
20            S = Replace(S, Chr$(13), vbNullString)
30        End If
40        If InStrB(S, Chr$(10)) <> 0 Then
50            S = Replace(S, Chr$(10), vbNullString)
60        End If
70        If InStrB(S, "¬") <> 0 Then
80            S = Replace(S, "¬", vbNullString)   'morgo usaba esto como "separador"
90        End If
End Sub

