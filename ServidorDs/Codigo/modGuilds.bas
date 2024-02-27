Attribute VB_Name = "modGuilds"
'**************************************************************
' modGuilds.bas - Module to allow the usage of areas instead of maps.
' Saves a lot of bandwidth.
'
' Implemented by Mariano Barrou (El Oso)
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

'guilds nueva version. Hecho por el oso, eliminando los problemas
'de sincronizacion con los datos en el HD... entre varios otros
'º¬

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DECLARACIOENS PUBLICAS CONCERNIENTES AL JUEGO
'Y CONFIGURACION DEL SISTEMA DE CLANES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private GUILDINFOFILE   As String
'archivo .\guilds\guildinfo.ini o similar

Private Const MAX_GUILDS As Integer = 1000
'cantidad maxima de guilds en el servidor

Public CANTIDADDECLANES As Integer
'cantidad actual de clanes en el servidor

Public guilds(1 To MAX_GUILDS) As clsClan
'array global de guilds, se indexa por userlist().guildindex

Private Const CANTIDADMAXIMACODEX As Byte = 8
'cantidad maxima de codecs que se pueden definir

Public Const MAXASPIRANTES As Byte = 10
'cantidad maxima de aspirantes que puede tener un clan acumulados a la vez

Private Const MAXANTIFACCION As Byte = 100
'puntos maximos de antifaccion que un clan tolera antes de ser cambiada su alineacion

Public Enum ALINEACION_GUILD
    ALINEACION_LEGION = 1
    ALINEACION_CRIMINAL = 2
    ALINEACION_NEUTRO = 3
    ALINEACION_CIUDA = 4
    ALINEACION_ARMADA = 5
    ALINEACION_MASTER = 6
End Enum
'alineaciones permitidas

Public Enum SONIDOS_GUILD
    SND_CREACIONCLAN = 44
    SND_ACEPTADOCLAN = 43
    SND_DECLAREWAR = 45
End Enum
'numero de .wav del cliente

Public Enum RELACIONES_GUILD
    GUERRA = -1
    PAZ = 0
    ALIADOS = 1
End Enum
'estado entre clanes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub LoadGuildsDB()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim CantClanes  As String
      Dim i           As Integer
      Dim tempStr     As String
      Dim Alin        As ALINEACION_GUILD
          
10        GUILDINFOFILE = App.Path & "\guilds\guildsinfo.inf"

20        CantClanes = GetVar(GUILDINFOFILE, "INIT", "nroGuilds")
          
30        If IsNumeric(CantClanes) Then
40            CANTIDADDECLANES = CInt(CantClanes)
50        Else
60            CANTIDADDECLANES = 0
70        End If
          
80        For i = 1 To CANTIDADDECLANES
90            Set guilds(i) = New clsClan
100           tempStr = GetVar(GUILDINFOFILE, "GUILD" & i, "GUILDNAME")
110           Alin = String2Alineacion(GetVar(GUILDINFOFILE, "GUILD" & i, "Alineacion"))
120           Call guilds(i).Inicializar(tempStr, i, Alin)
130       Next i
          
End Sub

Public Function m_ConectarMiembroAClan(ByVal Userindex As Integer, ByVal GuildIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************


      Dim NuevaA  As Boolean
      Dim News    As String

10        If GuildIndex > CANTIDADDECLANES Or GuildIndex <= 0 Then Exit Function 'x las dudas...
20        If m_EstadoPermiteEntrar(Userindex, GuildIndex) Then
30            Call guilds(GuildIndex).ConectarMiembro(Userindex)
40            UserList(Userindex).GuildIndex = GuildIndex
50            m_ConectarMiembroAClan = True
60        Else
70            m_ConectarMiembroAClan = m_ValidarPermanencia(Userindex, True, NuevaA)
80            If NuevaA Then News = News & "El clan tiene nueva alineación."
              'If NuevoL Or NuevaA Then Call guilds(GuildIndex).SetGuildNews(News)
90        End If

End Function


Public Function m_ValidarPermanencia(ByVal Userindex As Integer, ByVal SumaAntifaccion As Boolean, _
                            ByRef CambioAlineacion As Boolean) As Boolean
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 14/12/2009
      '25/03/2009: ZaMa - Desequipo los items faccionarios que tenga el funda al abandonar la faccion
      '14/12/2009: ZaMa - La alineacion del clan depende del lider
      '14/02/2010: ZaMa - Ya no es necesario saber si el lider cambia, ya que no puede cambiar.
      '***************************************************

      Dim GuildIndex  As Integer

10        m_ValidarPermanencia = True
          
20        GuildIndex = UserList(Userindex).GuildIndex
30        If GuildIndex > CANTIDADDECLANES And GuildIndex <= 0 Then Exit Function
          
40        If Not m_EstadoPermiteEntrar(Userindex, GuildIndex) Then
              
              ' Es el lider, bajamos 1 rango de alineacion
50            If GuildLeader(GuildIndex) = UserList(Userindex).Name Then
60                Call LogClanes(UserList(Userindex).Name & ", líder de " & guilds(GuildIndex).GuildName & " hizo bajar la alienación de su clan.")
              
70                CambioAlineacion = True
                  
                  ' Por si paso de ser armada/legion a pk/ciuda, chequeo de nuevo
80                Do
90                    Call UpdateGuildMembers(GuildIndex)
100               Loop Until m_EstadoPermiteEntrar(Userindex, GuildIndex)
110           Else
120               Call LogClanes(UserList(Userindex).Name & " de " & guilds(GuildIndex).GuildName & " es expulsado en validar permanencia.")
              
130               m_ValidarPermanencia = False
140               If SumaAntifaccion Then guilds(GuildIndex).PuntosAntifaccion = guilds(GuildIndex).PuntosAntifaccion + 1
                  
150               CambioAlineacion = guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION
                  
160               Call LogClanes(UserList(Userindex).Name & " de " & guilds(GuildIndex).GuildName & _
                      IIf(CambioAlineacion, " SI ", " NO ") & "provoca cambio de alineación. MAXANT:" & CambioAlineacion)
                  
170               Call m_EcharMiembroDeClan(-1, UserList(Userindex).Name)
                  
                  ' Llegamos a la maxima cantidad de antifacciones permitidas, bajamos un grado de alineación
180               If CambioAlineacion Then
190                   Call UpdateGuildMembers(GuildIndex)
200               End If
210           End If
220       End If
End Function

Private Sub UpdateGuildMembers(ByVal GuildIndex As Integer)
      '***************************************************
      'Autor: ZaMa
      'Last Modification: 14/01/2010 (ZaMa)
      '14/01/2010: ZaMa - Pulo detalles en el funcionamiento general.
      '***************************************************
          Dim GuildMembers() As String
          Dim TotalMembers As Integer
          Dim MemberIndex As Long
          Dim Sale As Boolean
          Dim MemberName As String
          Dim Userindex As Integer
          Dim Reenlistadas As Integer
          
          ' Si devuelve true, cambio a neutro y echamos a todos los que estén de mas, sino no echamos a nadie
10        If guilds(GuildIndex).CambiarAlineacion(BajarGrado(GuildIndex)) Then 'ALINEACION_NEUTRO)
              
              'uso GetMemberList y no los iteradores pq voy a rajar gente y puedo alterar
              'internamente al iterador en el proceso
20            GuildMembers = guilds(GuildIndex).GetMemberList()
30            TotalMembers = UBound(GuildMembers)
              
40            For MemberIndex = 0 To TotalMembers
50                MemberName = GuildMembers(MemberIndex)
                  
                  'vamos a violar un poco de capas..
60                Userindex = NameIndex(MemberName)
70                If Userindex > 0 Then
80                    Sale = Not m_EstadoPermiteEntrar(Userindex, GuildIndex)
90                Else
100                   Sale = Not m_EstadoPermiteEntrarChar(MemberName, GuildIndex)
110               End If

120               If Sale Then
130                   If m_EsGuildLeader(MemberName, GuildIndex) Then  'hay que sacarlo de las facciones
                       
140                       If Userindex > 0 Then
150                           If UserList(Userindex).Faccion.ArmadaReal <> 0 Then
160                               Call ExpulsarFaccionReal(Userindex)
                                  ' No cuenta como reenlistada :p.
170                               UserList(Userindex).Faccion.Reenlistadas = UserList(Userindex).Faccion.Reenlistadas - 1
180                           ElseIf UserList(Userindex).Faccion.FuerzasCaos <> 0 Then
190                               Call ExpulsarFaccionCaos(Userindex)
                                  ' No cuenta como reenlistada :p.
200                               UserList(Userindex).Faccion.Reenlistadas = UserList(Userindex).Faccion.Reenlistadas - 1
210                           End If
220                       Else
230                           If FileExist(CharPath & MemberName & ".chr") Then
240                               Call WriteVar(CharPath & MemberName & ".chr", "FACCIONES", "EjercitoCaos", 0)
250                               Call WriteVar(CharPath & MemberName & ".chr", "FACCIONES", "EjercitoReal", 0)
260                               Reenlistadas = GetVar(CharPath & MemberName & ".chr", "FACCIONES", "Reenlistadas")
270                               Call WriteVar(CharPath & MemberName & ".chr", "FACCIONES", "Reenlistadas", _
                                          IIf(Reenlistadas > 1, Reenlistadas - 1, Reenlistadas))
280                           End If
290                       End If
300                   Else    'sale si no es guildLeader
310                       Call m_EcharMiembroDeClan(-1, MemberName)
320                   End If
330               End If
340           Next MemberIndex
350       Else
              ' Resetea los puntos de antifacción
360           guilds(GuildIndex).PuntosAntifaccion = 0
370       End If
End Sub

Private Function BajarGrado(ByVal GuildIndex As Integer) As ALINEACION_GUILD
      '***************************************************
      'Autor: ZaMa
      'Last Modification: 27/11/2009
      'Reduce el grado de la alineacion a partir de la alineacion dada
      '***************************************************

10    Select Case guilds(GuildIndex).Alineacion
          Case ALINEACION_ARMADA
20            BajarGrado = ALINEACION_CIUDA
30        Case ALINEACION_LEGION
40            BajarGrado = ALINEACION_CRIMINAL
50        Case Else
60            BajarGrado = ALINEACION_NEUTRO
70    End Select

End Function

Public Sub m_DesconectarMiembroDelClan(ByVal Userindex As Integer, ByVal GuildIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    If UserList(Userindex).GuildIndex > CANTIDADDECLANES Then Exit Sub
20        Call guilds(GuildIndex).DesConectarMiembro(Userindex)
End Sub

Private Function m_EsGuildLeader(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        m_EsGuildLeader = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).GetLeader)))
End Function

Private Function m_EsGuildFounder(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        m_EsGuildFounder = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).Fundador)))
End Function

Public Function m_EcharMiembroDeClan(ByVal Expulsador As Integer, ByVal Expulsado As String, Optional ByVal Disolution As Boolean = False) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      'UI echa a Expulsado del clan de Expulsado
      Dim Userindex   As Integer
      Dim GI          As Integer
          
10        m_EcharMiembroDeClan = 0

20        Userindex = NameIndex(Expulsado)
30        If Userindex > 0 Then
              'pj online
40            GI = UserList(Userindex).GuildIndex
50            If GI > 0 Then
60                If m_PuedeSalirDeClan(Expulsado, GI, Expulsador, Disolution) Then
70                    Call guilds(GI).DesConectarMiembro(Userindex)
80                    Call guilds(GI).ExpulsarMiembro(Expulsado)
90                    Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
100                   UserList(Userindex).GuildIndex = 0
110                   Call RefreshCharStatus(Userindex)
120                   m_EcharMiembroDeClan = GI
130               Else
140                   m_EcharMiembroDeClan = 0
150               End If
160           Else
170               m_EcharMiembroDeClan = 0
180           End If
190       Else
              'pj offline
200           GI = GetGuildIndexFromChar(Expulsado)
210           If GI > 0 Then
220               If m_PuedeSalirDeClan(Expulsado, GI, Expulsador, Disolution) Then
230                   Call guilds(GI).ExpulsarMiembro(Expulsado)
240                   Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
250                   m_EcharMiembroDeClan = GI
260               Else
270                   m_EcharMiembroDeClan = 0
280               End If
290           Else
300               m_EcharMiembroDeClan = 0
310           End If
320       End If

End Function

Public Sub ActualizarWebSite(ByVal Userindex As Integer, ByRef Web As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim GI As Integer

10        GI = UserList(Userindex).GuildIndex
20        If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
          
30        If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then Exit Sub
          
40        Call guilds(GI).SetURL(Web)
          
End Sub


Public Sub ChangeCodexAndDesc(ByRef desc As String, ByRef codex() As String, ByVal GuildIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim i As Long
          
10        If GuildIndex < 1 Or GuildIndex > CANTIDADDECLANES Then Exit Sub
          
20        With guilds(GuildIndex)
30            Call .SetDesc(desc)
              
40            For i = 0 To UBound(codex())
50                Call .SetCodex(i, codex(i))
60            Next i
              
70            For i = i To CANTIDADMAXIMACODEX
80                Call .SetCodex(i, vbNullString)
90            Next i
100       End With
End Sub

Public Sub ActualizarNoticias(ByVal Userindex As Integer, ByRef Datos As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: 21/02/2010
      '21/02/2010: ZaMa - Ahora le avisa a los miembros que cambio el guildnews.
      '***************************************************

          Dim GI As Integer

10        With UserList(Userindex)
20            GI = .GuildIndex
              
30            If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
              
40            If Not m_EsGuildLeader(.Name, GI) Then Exit Sub
              
50            Call guilds(GI).SetGuildNews(Datos)
              
60            Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & " ha actualizado las noticias del clan!"))
70        End With
End Sub

Public Function CrearNuevoClan(ByVal FundadorIndex As Integer, ByRef desc As String, ByRef GuildName As String, ByRef URL As String, ByRef codex() As String, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim CantCodex       As Integer
      Dim i               As Integer
      Dim dummyString     As String

10        CrearNuevoClan = False
20        If Not PuedeFundarUnClan(FundadorIndex, Alineacion, dummyString) Then
30            refError = dummyString
40            Exit Function
50        End If

60        If GuildName = vbNullString Or Not GuildNameValido(GuildName) Then
70            refError = "Nombre de clan inválido."
80            Exit Function
90        End If
          
100       If YaExiste(GuildName) Then
110           refError = "Ya existe un clan con ese nombre."
120           Exit Function
130       End If

140       CantCodex = UBound(codex()) + 1

          'tenemos todo para fundar ya
150       If CANTIDADDECLANES < UBound(guilds) Then
160           CANTIDADDECLANES = CANTIDADDECLANES + 1
              'ReDim Preserve Guilds(1 To CANTIDADDECLANES) As clsClan

              'constructor custom de la clase clan
170           Set guilds(CANTIDADDECLANES) = New clsClan
              
180           With guilds(CANTIDADDECLANES)
190               Call .Inicializar(GuildName, CANTIDADDECLANES, Alineacion)
                  
                  'Damos de alta al clan como nuevo inicializando sus archivos
200               Call .InicializarNuevoClan(UserList(FundadorIndex).Name)
                  
                  'seteamos codex y descripcion
210               For i = 1 To CantCodex
220                   Call .SetCodex(i, codex(i - 1))
230               Next i
240               Call .SetDesc(desc)
250               Call .SetGuildNews("Clan creado con alineación: " & Alineacion2String(Alineacion))
260               Call .SetLeader(UserList(FundadorIndex).Name)
270               Call .SetURL(URL)
                  
                  '"conectamos" al nuevo miembro a la lista de la clase
280               Call .AceptarNuevoMiembro(UserList(FundadorIndex).Name)
290               Call .ConectarMiembro(FundadorIndex)
300           End With
              
310           UserList(FundadorIndex).GuildIndex = CANTIDADDECLANES
320           Call RefreshCharStatus(FundadorIndex)
              
330           For i = 1 To CANTIDADDECLANES - 1
340               Call guilds(i).ProcesarFundacionDeOtroClan
350           Next i
360       Else
370           refError = "No hay más slots para fundar clanes. Consulte a un administrador."
380           Exit Function
390       End If
          
400        Call QuitarObjetos(886, 1, FundadorIndex)
410        Call QuitarObjetos(887, 1, FundadorIndex)
420        Call QuitarObjetos(888, 1, FundadorIndex)
           
          
430       CrearNuevoClan = True
          
          
End Function

Public Sub SendGuildNews(ByVal Userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim GuildIndex  As Integer
      Dim i               As Integer
      Dim go As Integer

10        GuildIndex = UserList(Userindex).GuildIndex
20        If GuildIndex = 0 Then Exit Sub

          Dim enemies() As String
          
30        Exit Sub
40        With guilds(GuildIndex)
50            If .CantidadEnemys Then
60                ReDim enemies(0 To .CantidadEnemys - 1) As String
70            Else
80                ReDim enemies(0)
90            End If
              
              Dim allies() As String
              
100           If .CantidadAllies Then
110               ReDim allies(0 To .CantidadAllies - 1) As String
120           Else
130               ReDim allies(0)
140           End If
              
150           i = .Iterador_ProximaRelacion(RELACIONES_GUILD.GUERRA)
160           go = 0
              
170           While i > 0
180               enemies(go) = guilds(i).GuildName
190               i = .Iterador_ProximaRelacion(RELACIONES_GUILD.GUERRA)
200               go = go + 1
210           Wend
              
220           i = .Iterador_ProximaRelacion(RELACIONES_GUILD.ALIADOS)
230           go = 0
              
240           While i > 0
250               allies(go) = guilds(i).GuildName
260               i = .Iterador_ProximaRelacion(RELACIONES_GUILD.ALIADOS)
270           Wend
          
280           Call WriteGuildNews(Userindex, .GetGuildNews, enemies, allies)
          
290           If .EleccionesAbiertas Then
300               Call WriteConsoleMsg(Userindex, "Hoy es la votación para elegir un nuevo líder para el clan.", FontTypeNames.FONTTYPE_GUILD)
310               Call WriteConsoleMsg(Userindex, "La elección durará 24 horas, se puede votar a cualquier miembro del clan.", FontTypeNames.FONTTYPE_GUILD)
320               Call WriteConsoleMsg(Userindex, "Para votar escribe /VOTO NICKNAME.", FontTypeNames.FONTTYPE_GUILD)
330               Call WriteConsoleMsg(Userindex, "Sólo se computará un voto por miembro. Tu voto no puede ser cambiado.", FontTypeNames.FONTTYPE_GUILD)
340           End If
350       End With

End Sub

Public Function m_PuedeSalirDeClan(ByRef Nombre As String, ByVal GuildIndex As Integer, ByVal QuienLoEchaUI As Integer, Optional ByVal Disolution As Boolean = False) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      'sale solo si no es fundador del clan.

10        m_PuedeSalirDeClan = False
20        If GuildIndex = 0 Then Exit Function
          
          'esto es un parche, si viene en -1 es porque la invoca la rutina de expulsion automatica de clanes x antifacciones
30        If QuienLoEchaUI = -1 Then
40            m_PuedeSalirDeClan = True
50            Exit Function
60        End If

          'cuando UI no puede echar a nombre?
          'si no es gm Y no es lider del clan del pj Y no es el mismo que se va voluntariamente
70        If UserList(QuienLoEchaUI).flags.Privilegios And PlayerType.User Then
80            If Not m_EsGuildLeader(UCase$(UserList(QuienLoEchaUI).Name), GuildIndex) Then
90                If UCase$(UserList(QuienLoEchaUI).Name) <> UCase$(Nombre) Then      'si no sale voluntariamente...
100                   Exit Function
110               End If
120           End If
130       End If
          
          ' Ahora el lider es el unico que no puede salir del clan
140       If Disolution = False Then
150           m_PuedeSalirDeClan = UCase$(guilds(GuildIndex).GetLeader) <> UCase$(Nombre)
160       Else
170           m_PuedeSalirDeClan = True
180       End If

End Function

Public Function PuedeFundarUnClan(ByVal Userindex As Integer, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean
      '***************************************************
      'Autor: Unknown
      'Last Modification: 27/11/2009
      'Returns true if can Found a guild
      '27/11/2009: ZaMa - Ahora valida si ya fundo clan o no.
      '***************************************************
          
10        If UserList(Userindex).GuildIndex > 0 Then
20            refError = "Ya perteneces a un clan, no puedes fundar otro"
30            Exit Function
40        End If
          
50         If (UserList(Userindex).Stats.Gld < 25000000) _
       Or TieneObjetos(886, 1, Userindex) = False _
       Or TieneObjetos(887, 1, Userindex) = False _
        Or TieneObjetos(888, 1, Userindex) = False _
          Or (UserList(Userindex).Stats.UserSkills(eSkill.Liderazgo) < 100) Then

60            refError = "Para fundar clan necesitas encontrar los tres amuletos de Lider, disponer de 25.000.000 monedas de oro en tu billetera y tener 100 puntos en liderazgo. "
70            Exit Function
80        End If
          
90        Select Case Alineacion
              Case ALINEACION_GUILD.ALINEACION_ARMADA
100               If UserList(Userindex).Faccion.ArmadaReal <> 1 Then
110                   refError = "Para fundar un clan real debes ser miembro del ejército real."
120                   Exit Function
130               End If
140           Case ALINEACION_GUILD.ALINEACION_CIUDA
150               If criminal(Userindex) Then
160                   refError = "Para fundar un clan de ciudadanos no debes ser criminal."
170                   Exit Function
180               End If
190           Case ALINEACION_GUILD.ALINEACION_CRIMINAL
200               If Not criminal(Userindex) Then
210                   refError = "Para fundar un clan de criminales no debes ser ciudadano."
220                   Exit Function
230               End If
240           Case ALINEACION_GUILD.ALINEACION_LEGION
250               If UserList(Userindex).Faccion.FuerzasCaos <> 1 Then
260                   refError = "Para fundar un clan del mal debes pertenecer a la legión oscura."
270                   Exit Function
280               End If
290           Case ALINEACION_GUILD.ALINEACION_MASTER
300               If UserList(Userindex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
310                   refError = "Para fundar un clan sin alineación debes ser un dios."
320                   Exit Function
330               End If
340           Case ALINEACION_GUILD.ALINEACION_NEUTRO
350               If UserList(Userindex).Faccion.ArmadaReal <> 0 Or UserList(Userindex).Faccion.FuerzasCaos <> 0 Then
360                   refError = "Para fundar un clan neutro no debes pertenecer a ninguna facción."
370                   Exit Function
380               End If
390       End Select
          
          
400       PuedeFundarUnClan = True
          
End Function

Private Function m_EstadoPermiteEntrarChar(ByRef Personaje As String, ByVal GuildIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim Promedio    As Long
      Dim ELV         As Integer
      Dim f           As Byte

10        m_EstadoPermiteEntrarChar = False
          
20        If InStrB(Personaje, "\") <> 0 Then
30            Personaje = Replace(Personaje, "\", vbNullString)
40        End If
50        If InStrB(Personaje, "/") <> 0 Then
60            Personaje = Replace(Personaje, "/", vbNullString)
70        End If
80        If InStrB(Personaje, ".") <> 0 Then
90            Personaje = Replace(Personaje, ".", vbNullString)
100       End If
          
110       If FileExist(CharPath & Personaje & ".chr") Then
120           Promedio = CLng(GetVar(CharPath & Personaje & ".chr", "REP", "Promedio"))
130           Select Case guilds(GuildIndex).Alineacion
                  Case ALINEACION_GUILD.ALINEACION_ARMADA
140                   If Promedio >= 0 Then
150                       ELV = CInt(GetVar(CharPath & Personaje & ".chr", "Stats", "ELV"))
160                       If ELV >= 25 Then
170                           f = CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "EjercitoReal"))
180                       End If
190                       m_EstadoPermiteEntrarChar = IIf(ELV >= 25, f <> 0, True)
200                   End If
210               Case ALINEACION_GUILD.ALINEACION_CIUDA
220                   m_EstadoPermiteEntrarChar = Promedio >= 0
230               Case ALINEACION_GUILD.ALINEACION_CRIMINAL
240                   m_EstadoPermiteEntrarChar = Promedio < 0
250               Case ALINEACION_GUILD.ALINEACION_NEUTRO
260                   m_EstadoPermiteEntrarChar = CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "EjercitoReal")) = 0
270                   m_EstadoPermiteEntrarChar = m_EstadoPermiteEntrarChar And (CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "EjercitoCaos")) = 0)
280               Case ALINEACION_GUILD.ALINEACION_LEGION
290                   If Promedio < 0 Then
300                       ELV = CInt(GetVar(CharPath & Personaje & ".chr", "Stats", "ELV"))
310                       If ELV >= 25 Then
320                           f = CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "EjercitoCaos"))
330                       End If
340                       m_EstadoPermiteEntrarChar = IIf(ELV >= 25, f <> 0, True)
350                   End If
360               Case Else
370                   m_EstadoPermiteEntrarChar = True
380           End Select
390       End If
End Function

Private Function m_EstadoPermiteEntrar(ByVal Userindex As Integer, ByVal GuildIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        Select Case guilds(GuildIndex).Alineacion
              Case ALINEACION_GUILD.ALINEACION_ARMADA
20                m_EstadoPermiteEntrar = Not criminal(Userindex) And _
                          IIf(UserList(Userindex).Stats.ELV >= 25, UserList(Userindex).Faccion.ArmadaReal <> 0, True)
              
30            Case ALINEACION_GUILD.ALINEACION_LEGION
40                m_EstadoPermiteEntrar = criminal(Userindex) And _
                          IIf(UserList(Userindex).Stats.ELV >= 25, UserList(Userindex).Faccion.FuerzasCaos <> 0, True)
              
50            Case ALINEACION_GUILD.ALINEACION_NEUTRO
60                m_EstadoPermiteEntrar = UserList(Userindex).Faccion.ArmadaReal = 0 And UserList(Userindex).Faccion.FuerzasCaos = 0
              
70            Case ALINEACION_GUILD.ALINEACION_CIUDA
80                m_EstadoPermiteEntrar = Not criminal(Userindex)
              
90            Case ALINEACION_GUILD.ALINEACION_CRIMINAL
100               m_EstadoPermiteEntrar = criminal(Userindex)
              
110           Case Else   'game masters
120               m_EstadoPermiteEntrar = True
130       End Select
End Function

Public Function String2Alineacion(ByRef S As String) As ALINEACION_GUILD
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        Select Case S
              Case "Neutral"
20                String2Alineacion = ALINEACION_NEUTRO
30            Case "Del Mal"
40                String2Alineacion = ALINEACION_LEGION
50            Case "Real"
60                String2Alineacion = ALINEACION_ARMADA
70            Case "Game Masters"
80                String2Alineacion = ALINEACION_MASTER
90            Case "Legal"
100               String2Alineacion = ALINEACION_CIUDA
110           Case "Criminal"
120               String2Alineacion = ALINEACION_CRIMINAL
130       End Select
End Function

Public Function Alineacion2String(ByVal Alineacion As ALINEACION_GUILD) As String
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        Select Case Alineacion
              Case ALINEACION_GUILD.ALINEACION_NEUTRO
20                Alineacion2String = "Neutral"
30            Case ALINEACION_GUILD.ALINEACION_LEGION
40                Alineacion2String = "Del Mal"
50            Case ALINEACION_GUILD.ALINEACION_ARMADA
60                Alineacion2String = "Real"
70            Case ALINEACION_GUILD.ALINEACION_MASTER
80                Alineacion2String = "Game Masters"
90            Case ALINEACION_GUILD.ALINEACION_CIUDA
100               Alineacion2String = "Legal"
110           Case ALINEACION_GUILD.ALINEACION_CRIMINAL
120               Alineacion2String = "Criminal"
130       End Select
End Function

Public Function Relacion2String(ByVal Relacion As RELACIONES_GUILD) As String
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        Select Case Relacion
              Case RELACIONES_GUILD.ALIADOS
20                Relacion2String = "A"
30            Case RELACIONES_GUILD.GUERRA
40                Relacion2String = "G"
50            Case RELACIONES_GUILD.PAZ
60                Relacion2String = "P"
70            Case RELACIONES_GUILD.ALIADOS
80                Relacion2String = "?"
90        End Select
End Function

Public Function String2Relacion(ByVal S As String) As RELACIONES_GUILD
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        Select Case UCase$(Trim$(S))
              Case vbNullString, "P"
20                String2Relacion = RELACIONES_GUILD.PAZ
30            Case "G"
40                String2Relacion = RELACIONES_GUILD.GUERRA
50            Case "A"
60                String2Relacion = RELACIONES_GUILD.ALIADOS
70            Case Else
80                String2Relacion = RELACIONES_GUILD.PAZ
90        End Select
End Function

Private Function GuildNameValido(ByVal cad As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim Car     As Byte
      Dim i       As Integer

      'old function by morgo

10    cad = LCase$(cad)

20    For i = 1 To Len(cad)
30        Car = Asc(mid$(cad, i, 1))

40        If (Car < 97 Or Car > 122) And (Car <> 255) And (Car <> 32) Then
50            GuildNameValido = False
60            Exit Function
70        End If
          
80    Next i

90    GuildNameValido = True

End Function

Private Function YaExiste(ByVal GuildName As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim i   As Integer

10    YaExiste = False
20    GuildName = UCase$(GuildName)

30    For i = 1 To CANTIDADDECLANES
40        YaExiste = (UCase$(guilds(i).GuildName) = GuildName)
50        If YaExiste Then Exit Function
60    Next i

End Function

Public Function HasFound(ByRef UserName As String) As Boolean
      '***************************************************
      'Autor: ZaMa
      'Last Modification: 27/11/2009
      'Returns true if it's already the founder of other guild
      '***************************************************
      Dim i As Long
      Dim Name As String

10    Name = UCase$(UserName)

20    For i = 1 To CANTIDADDECLANES
30        HasFound = (UCase$(guilds(i).Fundador) = Name)
40        If HasFound Then Exit Function
50    Next i

End Function

Public Function v_AbrirElecciones(ByVal Userindex As Integer, ByRef refError As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim GuildIndex      As Integer

10        v_AbrirElecciones = False
20        GuildIndex = UserList(Userindex).GuildIndex
          
30        If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
40            refError = "Tú no perteneces a ningún clan."
50            Exit Function
60        End If
          
70        If Not m_EsGuildLeader(UserList(Userindex).Name, GuildIndex) Then
80            refError = "No eres el líder de tu clan"
90            Exit Function
100       End If
          
110       If guilds(GuildIndex).EleccionesAbiertas Then
120           refError = "Las elecciones ya están abiertas."
130           Exit Function
140       End If
          
150       v_AbrirElecciones = True
160       Call guilds(GuildIndex).AbrirElecciones
          
End Function

Public Function v_UsuarioVota(ByVal Userindex As Integer, ByRef Votado As String, ByRef refError As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim GuildIndex      As Integer
      Dim list()          As String
      Dim i As Long

10        v_UsuarioVota = False
20        GuildIndex = UserList(Userindex).GuildIndex
          
30        If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
40            refError = "Tú no perteneces a ningún clan."
50            Exit Function
60        End If

70        With guilds(GuildIndex)
80            If Not .EleccionesAbiertas Then
90                refError = "No hay elecciones abiertas en tu clan."
100               Exit Function
110           End If
              
              
120           list = .GetMemberList()
130           For i = 0 To UBound(list())
140               If UCase$(Votado) = list(i) Then Exit For
150           Next i
              
160           If i > UBound(list()) Then
170               refError = Votado & " no pertenece al clan."
180               Exit Function
190           End If
              
              
200           If .YaVoto(UserList(Userindex).Name) Then
210               refError = "Ya has votado, no puedes cambiar tu voto."
220               Exit Function
230           End If
              
240           Call .ContabilizarVoto(UserList(Userindex).Name, Votado)
250           v_UsuarioVota = True
260       End With

End Function

Public Sub v_RutinaElecciones()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim i       As Integer

10    On Error GoTo errh
20        For i = 1 To CANTIDADDECLANES
30            If Not guilds(i) Is Nothing Then
40                If guilds(i).RevisarElecciones Then
50                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & guilds(i).GetLeader & " es el nuevo líder de " & guilds(i).GuildName & ".", FontTypeNames.FONTTYPE_SERVER))
60                End If
70            End If
proximo:
80        Next i
90    Exit Sub
errh:
100       Call LogError("modGuilds.v_RutinaElecciones():" & Err.Description)
110       Resume proximo
End Sub

Private Function GetGuildIndexFromChar(ByRef PlayerName As String) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      'aca si que vamos a violar las capas deliveradamente ya que
      'visual basic no permite declarar metodos de clase
      Dim Temps   As String
10        If InStrB(PlayerName, "\") <> 0 Then
20            PlayerName = Replace(PlayerName, "\", vbNullString)
30        End If
40        If InStrB(PlayerName, "/") <> 0 Then
50            PlayerName = Replace(PlayerName, "/", vbNullString)
60        End If
70        If InStrB(PlayerName, ".") <> 0 Then
80            PlayerName = Replace(PlayerName, ".", vbNullString)
90        End If
100       Temps = GetVar(CharPath & PlayerName & ".chr", "GUILD", "GUILDINDEX")
110       If IsNumeric(Temps) Then
120           GetGuildIndexFromChar = CInt(Temps)
130       Else
140           GetGuildIndexFromChar = 0
150       End If
End Function

Public Function GuildIndex(ByRef GuildName As String) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      'me da el indice del guildname
      Dim i As Integer

10        GuildIndex = 0
20        GuildName = UCase$(GuildName)
30        For i = 1 To CANTIDADDECLANES
40            If UCase$(guilds(i).GuildName) = GuildName Then
50                GuildIndex = i
60                Exit Function
70            End If
80        Next i
End Function

Public Function m_ListaDeMiembrosOnline(ByVal Userindex As Integer, ByVal GuildIndex As Integer, Optional ByVal InCVC As Boolean = False) As String
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim i As Integer

          
10        If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
20            i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
30            While i > 0
                  'No mostramos dioses y admins
40                If Not InCVC Then
50                    If i <> Userindex And ((UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or (UserList(Userindex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0)) Then _
                          m_ListaDeMiembrosOnline = m_ListaDeMiembrosOnline & UserList(i).Name & ","
60                    i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
70                Else
80                    m_ListaDeMiembrosOnline = m_ListaDeMiembrosOnline & UserList(i).Name & ","
90                    i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
100               End If

110           Wend
120       End If
130       If Len(m_ListaDeMiembrosOnline) > 0 Then
140           m_ListaDeMiembrosOnline = Left$(m_ListaDeMiembrosOnline, Len(m_ListaDeMiembrosOnline) - 1)
150       End If
End Function

Public Function PrepareGuildsList() As String()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim tStr() As String
          Dim i As Long
          
10        If CANTIDADDECLANES = 0 Then
20            ReDim tStr(0) As String
30        Else
40            ReDim tStr(CANTIDADDECLANES - 1) As String
              
50            For i = 1 To CANTIDADDECLANES
60                tStr(i - 1) = guilds(i).GuildName
70            Next i
80        End If
          
90        PrepareGuildsList = tStr
End Function

Public Sub SendGuildDetails(ByVal Userindex As Integer, ByRef GuildName As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim codex(CANTIDADMAXIMACODEX - 1)  As String
          Dim GI      As Integer
          Dim i       As Long

10        GI = GuildIndex(GuildName)
20        If GI = 0 Then Exit Sub
          
30        With guilds(GI)
40            For i = 1 To CANTIDADMAXIMACODEX
50                codex(i - 1) = .GetCodex(i)
60            Next i
              
70            Call Protocol.WriteGuildDetails(Userindex, GuildName, .Fundador, .GetFechaFundacion, .GetLeader, _
                                          .GetURL, .CantidadDeMiembros, .EleccionesAbiertas, Alineacion2String(.Alineacion), _
                                          .CantidadEnemys, .CantidadAllies, .PuntosAntifaccion & "/" & CStr(MAXANTIFACCION), _
                                          codex, .GetDesc)
80        End With
End Sub

Public Sub SendGuildLeaderInfo(ByVal Userindex As Integer)
      '***************************************************
      'Autor: Mariano Barrou (El Oso)
      'Last Modification: 12/10/06
      'Las Modified By: Juan Martín Sotuyo Dodero (Maraxus)
      '***************************************************
          Dim GI      As Integer
          Dim guildList() As String
          Dim MemberList() As String
          Dim aspirantsList() As String

10        With UserList(Userindex)
20            GI = .GuildIndex
              
30            guildList = PrepareGuildsList()
              
40            If GI <= 0 Or GI > CANTIDADDECLANES Then
                  'Send the guild list instead
50                Call WriteGuildList(Userindex, guildList)
60                Exit Sub
70            End If
              
80            MemberList = guilds(GI).GetMemberList()
              
90            If Not m_EsGuildLeader(.Name, GI) Then
                  'Send the guild list instead
100               Call WriteGuildMemberInfo(Userindex, guildList, MemberList)
110               Exit Sub
120           End If
              
130           aspirantsList = guilds(GI).GetAspirantes()
              
140           Call WriteGuildLeaderInfo(Userindex, guildList, MemberList, guilds(GI).GetGuildNews(), aspirantsList)
150       End With
End Sub


Public Function m_Iterador_ProximoUserIndex(ByVal GuildIndex As Integer) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          'itera sobre los onlinemembers
10        m_Iterador_ProximoUserIndex = 0
20        If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
30            m_Iterador_ProximoUserIndex = guilds(GuildIndex).m_Iterador_ProximoUserIndex()
40        End If
End Function

Public Function Iterador_ProximoGM(ByVal GuildIndex As Integer) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          'itera sobre los gms escuchando este clan
10        Iterador_ProximoGM = 0
20        If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
30            Iterador_ProximoGM = guilds(GuildIndex).Iterador_ProximoGM()
40        End If
End Function

Public Function r_Iterador_ProximaPropuesta(ByVal GuildIndex As Integer, ByVal Tipo As RELACIONES_GUILD) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          'itera sobre las propuestas
10        r_Iterador_ProximaPropuesta = 0
20        If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
30            r_Iterador_ProximaPropuesta = guilds(GuildIndex).Iterador_ProximaPropuesta(Tipo)
40        End If
End Function

Public Function GMEscuchaClan(ByVal Userindex As Integer, ByVal GuildName As String) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim GI As Integer

          'listen to no guild at all
10        If LenB(GuildName) = 0 And UserList(Userindex).EscucheClan <> 0 Then
              'Quit listening to previous guild!!
20            Call WriteConsoleMsg(Userindex, "Dejas de escuchar a : " & guilds(UserList(Userindex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
30            guilds(UserList(Userindex).EscucheClan).DesconectarGM (Userindex)
40            Exit Function
50        End If
          
      'devuelve el guildindex
60        GI = GuildIndex(GuildName)
70        If GI > 0 Then
80            If UserList(Userindex).EscucheClan <> 0 Then
90                If UserList(Userindex).EscucheClan = GI Then
                      'Already listening to them...
100                   Call WriteConsoleMsg(Userindex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
110                   GMEscuchaClan = GI
120                   Exit Function
130               Else
                      'Quit listening to previous guild!!
140                   Call WriteConsoleMsg(Userindex, "Dejas de escuchar a : " & guilds(UserList(Userindex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
150                   guilds(UserList(Userindex).EscucheClan).DesconectarGM (Userindex)
160               End If
170           End If
              
180           Call guilds(GI).ConectarGM(Userindex)
190           Call WriteConsoleMsg(Userindex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
200           GMEscuchaClan = GI
210           UserList(Userindex).EscucheClan = GI
220       Else
230           Call WriteConsoleMsg(Userindex, "Error, el clan no existe.", FontTypeNames.FONTTYPE_GUILD)
240           GMEscuchaClan = 0
250       End If
          
End Function

Public Sub GMDejaDeEscucharClan(ByVal Userindex As Integer, ByVal GuildIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      'el index lo tengo que tener de cuando me puse a escuchar
10        UserList(Userindex).EscucheClan = 0
20        Call guilds(GuildIndex).DesconectarGM(Userindex)
End Sub
Public Function r_DeclararGuerra(ByVal Userindex As Integer, ByRef GuildGuerra As String, ByRef refError As String) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim GI  As Integer
      Dim GIG As Integer

10        r_DeclararGuerra = 0
20        GI = UserList(Userindex).GuildIndex
          
30                            If Not UserList(Userindex).Pos.map = 205 Then
40        WriteConsoleMsg Userindex, "Sistema de guerras de clan deshabilitado momentáneamente.", FontTypeNames.FONTTYPE_INFO
50        Exit Function
60         End If
          
70        If GI <= 0 Or GI > CANTIDADDECLANES Then
80            refError = "No eres miembro de ningún clan."
90            Exit Function
100       End If
          
110       If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
120           refError = "No eres el líder de tu clan."
130           Exit Function
140       End If
          
150       If Trim$(GuildGuerra) = vbNullString Then
160           refError = "No has seleccionado ningún clan."
170           Exit Function
180       End If
          
190       GIG = GuildIndex(GuildGuerra)
200       If guilds(GI).GetRelacion(GIG) = GUERRA Then
210           refError = "Tu clan ya está en guerra con " & GuildGuerra & "."
220           Exit Function
230       End If
              
240       If GI = GIG Then
250           refError = "No puedes declarar la guerra a tu mismo clan."
260           Exit Function
270       End If

280       If GIG < 1 Or GIG > CANTIDADDECLANES Then
290           Call LogError("ModGuilds.r_DeclararGuerra: " & GI & " declara a " & GuildGuerra)
300           refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
310           Exit Function
320       End If

330       Call guilds(GI).AnularPropuestas(GIG)
340       Call guilds(GIG).AnularPropuestas(GI)
350       Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.GUERRA)
360       Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.GUERRA)
          
370       r_DeclararGuerra = GIG

End Function


Public Function r_AceptarPropuestaDePaz(ByVal Userindex As Integer, ByRef GuildPaz As String, ByRef refError As String) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
      Dim GI      As Integer
      Dim GIG     As Integer

10                            If Not UserList(Userindex).Pos.map = 205 Then
20        WriteConsoleMsg Userindex, "Sistema de guerras de clan deshabilitado momentáneamente.", FontTypeNames.FONTTYPE_INFO
30        Exit Function
40         End If

50        GI = UserList(Userindex).GuildIndex
60        If GI <= 0 Or GI > CANTIDADDECLANES Then
70            refError = "No eres miembro de ningún clan."
80            Exit Function
90        End If
          
100       If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
110           refError = "No eres el líder de tu clan."
120           Exit Function
130       End If
          
140       If Trim$(GuildPaz) = vbNullString Then
150           refError = "No has seleccionado ningún clan."
160           Exit Function
170       End If

180       GIG = GuildIndex(GuildPaz)
          
190       If GIG < 1 Or GIG > CANTIDADDECLANES Then
200           Call LogError("ModGuilds.r_AceptarPropuestaDePaz: " & GI & " acepta de " & GuildPaz)
210           refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
220           Exit Function
230       End If

240       If guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.GUERRA Then
250           refError = "No estás en guerra con ese clan."
260           Exit Function
270       End If
          
280       If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
290           refError = "No hay ninguna propuesta de paz para aceptar."
300           Exit Function
310       End If

320       Call guilds(GI).AnularPropuestas(GIG)
330       Call guilds(GIG).AnularPropuestas(GI)
340       Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.PAZ)
350       Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.PAZ)
          
360       r_AceptarPropuestaDePaz = GIG
End Function

Public Function r_RechazarPropuestaDeAlianza(ByVal Userindex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      'devuelve el index al clan guildPro
      Dim GI      As Integer
      Dim GIG     As Integer


10                            If Not UserList(Userindex).Pos.map = 205 Then
20        WriteConsoleMsg Userindex, "Sistema de guerras de clan deshabilitado momentáneamente.", FontTypeNames.FONTTYPE_INFO
30        Exit Function
40         End If

50        r_RechazarPropuestaDeAlianza = 0
60        GI = UserList(Userindex).GuildIndex
          
70        If GI <= 0 Or GI > CANTIDADDECLANES Then
80            refError = "No eres miembro de ningún clan."
90            Exit Function
100       End If
          
110       If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
120           refError = "No eres el líder de tu clan."
130           Exit Function
140       End If
          
150       If Trim$(GuildPro) = vbNullString Then
160           refError = "No has seleccionado ningún clan."
170           Exit Function
180       End If

190       GIG = GuildIndex(GuildPro)
          
200       If GIG < 1 Or GIG > CANTIDADDECLANES Then
210           Call LogError("ModGuilds.r_RechazarPropuestaDeAlianza: " & GI & " acepta de " & GuildPro)
220           refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
230           Exit Function
240       End If
          
250       If Not guilds(GI).HayPropuesta(GIG, ALIADOS) Then
260           refError = "No hay propuesta de alianza del clan " & GuildPro
270           Exit Function
280       End If
          
290       Call guilds(GI).AnularPropuestas(GIG)
          'avisamos al otro clan
300       Call guilds(GIG).SetGuildNews(guilds(GI).GuildName & " ha rechazado nuestra propuesta de alianza. " & guilds(GIG).GetGuildNews())
310       r_RechazarPropuestaDeAlianza = GIG

End Function


Public Function r_RechazarPropuestaDePaz(ByVal Userindex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      'devuelve el index al clan guildPro
      Dim GI      As Integer
      Dim GIG     As Integer

10                            If Not UserList(Userindex).Pos.map = 205 Then
20        WriteConsoleMsg Userindex, "Sistema de guerras de clan deshabilitado momentáneamente.", FontTypeNames.FONTTYPE_INFO
30        Exit Function
40         End If

50        r_RechazarPropuestaDePaz = 0
60        GI = UserList(Userindex).GuildIndex
          
70        If GI <= 0 Or GI > CANTIDADDECLANES Then
80            refError = "No eres miembro de ningún clan."
90            Exit Function
100       End If
          
110       If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
120           refError = "No eres el líder de tu clan."
130           Exit Function
140       End If
          
150       If Trim$(GuildPro) = vbNullString Then
160           refError = "No has seleccionado ningún clan."
170           Exit Function
180       End If

190       GIG = GuildIndex(GuildPro)
          
200       If GIG < 1 Or GIG > CANTIDADDECLANES Then
210           Call LogError("ModGuilds.r_RechazarPropuestaDePaz: " & GI & " acepta de " & GuildPro)
220           refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
230           Exit Function
240       End If
          
250       If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
260           refError = "No hay propuesta de paz del clan " & GuildPro
270           Exit Function
280       End If
          
290       Call guilds(GI).AnularPropuestas(GIG)
          'avisamos al otro clan
300       Call guilds(GIG).SetGuildNews(guilds(GI).GuildName & " ha rechazado nuestra propuesta de paz. " & guilds(GIG).GetGuildNews())
310       r_RechazarPropuestaDePaz = GIG

End Function

Public Function r_AceptarPropuestaDeAlianza(ByVal Userindex As Integer, ByRef GuildAllie As String, ByRef refError As String) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
      Dim GI      As Integer
      Dim GIG     As Integer


10                            If Not UserList(Userindex).Pos.map = 205 Then
20        WriteConsoleMsg Userindex, "Sistema de guerras de clan deshabilitado momentáneamente.", FontTypeNames.FONTTYPE_INFO
30        Exit Function
40         End If

50        r_AceptarPropuestaDeAlianza = 0
60        GI = UserList(Userindex).GuildIndex
70        If GI <= 0 Or GI > CANTIDADDECLANES Then
80            refError = "No eres miembro de ningún clan."
90            Exit Function
100       End If
          
110       If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
120           refError = "No eres el líder de tu clan."
130           Exit Function
140       End If
          
150       If Trim$(GuildAllie) = vbNullString Then
160           refError = "No has seleccionado ningún clan."
170           Exit Function
180       End If

190       GIG = GuildIndex(GuildAllie)
          
200       If GIG < 1 Or GIG > CANTIDADDECLANES Then
210           Call LogError("ModGuilds.r_AceptarPropuestaDeAlianza: " & GI & " acepta de " & GuildAllie)
220           refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
230           Exit Function
240       End If

250       If guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.PAZ Then
260           refError = "No estás en paz con el clan, solo puedes aceptar propuesas de alianzas con alguien que estes en paz."
270           Exit Function
280       End If
          
290       If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.ALIADOS) Then
300           refError = "No hay ninguna propuesta de alianza para aceptar."
310           Exit Function
320       End If

330       Call guilds(GI).AnularPropuestas(GIG)
340       Call guilds(GIG).AnularPropuestas(GI)
350       Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.ALIADOS)
360       Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.ALIADOS)
          
370       r_AceptarPropuestaDeAlianza = GIG

End Function


Public Function r_ClanGeneraPropuesta(ByVal Userindex As Integer, ByRef OtroClan As String, ByVal Tipo As RELACIONES_GUILD, ByRef Detalle As String, ByRef refError As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim OtroClanGI      As Integer
      Dim GI              As Integer


10                            If Not UserList(Userindex).Pos.map = 205 Then
20        WriteConsoleMsg Userindex, "Sistema de guerras de clan deshabilitado momentáneamente.", FontTypeNames.FONTTYPE_INFO
30        Exit Function
40         End If

50        r_ClanGeneraPropuesta = False
          
60        GI = UserList(Userindex).GuildIndex
70        If GI <= 0 Or GI > CANTIDADDECLANES Then
80            refError = "No eres miembro de ningún clan."
90            Exit Function
100       End If
          
110       OtroClanGI = GuildIndex(OtroClan)
          
120       If OtroClanGI = GI Then
130           refError = "No puedes declarar relaciones con tu propio clan."
140           Exit Function
150       End If
          
160       If OtroClanGI <= 0 Or OtroClanGI > CANTIDADDECLANES Then
170           refError = "El sistema de clanes esta inconsistente, el otro clan no existe."
180           Exit Function
190       End If
          
200       If guilds(OtroClanGI).HayPropuesta(GI, Tipo) Then
210           refError = "Ya hay propuesta de " & Relacion2String(Tipo) & " con " & OtroClan
220           Exit Function
230       End If
          
240       If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
250           refError = "No eres el líder de tu clan."
260           Exit Function
270       End If
          
          'de acuerdo al tipo procedemos validando las transiciones
280       If Tipo = RELACIONES_GUILD.PAZ Then
290           If guilds(GI).GetRelacion(OtroClanGI) <> RELACIONES_GUILD.GUERRA Then
300               refError = "No estás en guerra con " & OtroClan
310               Exit Function
320           End If
330       ElseIf Tipo = RELACIONES_GUILD.GUERRA Then
              'por ahora no hay propuestas de guerra
340       ElseIf Tipo = RELACIONES_GUILD.ALIADOS Then
350           If guilds(GI).GetRelacion(OtroClanGI) <> RELACIONES_GUILD.PAZ Then
360               refError = "Para solicitar alianza no debes estar ni aliado ni en guerra con " & OtroClan
370               Exit Function
380           End If
390       End If
          
400       Call guilds(OtroClanGI).SetPropuesta(Tipo, GI, Detalle)
410       r_ClanGeneraPropuesta = True

End Function

Public Function r_VerPropuesta(ByVal Userindex As Integer, ByRef OtroGuild As String, ByVal Tipo As RELACIONES_GUILD, ByRef refError As String) As String
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim OtroClanGI      As Integer
      Dim GI              As Integer
          
10        r_VerPropuesta = vbNullString
20        refError = vbNullString
          
30        GI = UserList(Userindex).GuildIndex
40        If GI <= 0 Or GI > CANTIDADDECLANES Then
50            refError = "No eres miembro de ningún clan."
60            Exit Function
70        End If
          
80        If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
90            refError = "No eres el líder de tu clan."
100           Exit Function
110       End If
          
120       OtroClanGI = GuildIndex(OtroGuild)
          
130       If Not guilds(GI).HayPropuesta(OtroClanGI, Tipo) Then
140           refError = "No existe la propuesta solicitada."
150           Exit Function
160       End If
          
170       r_VerPropuesta = guilds(GI).GetPropuesta(OtroClanGI, Tipo)
          
End Function

Public Function r_ListaDePropuestas(ByVal Userindex As Integer, ByVal Tipo As RELACIONES_GUILD) As String()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim GI  As Integer
          Dim i   As Integer
          Dim proposalCount As Integer
          Dim proposals() As String
          
10        GI = UserList(Userindex).GuildIndex
          
20        If GI > 0 And GI <= CANTIDADDECLANES Then
30            With guilds(GI)
40                proposalCount = .CantidadPropuestas(Tipo)
                  
                  'Resize array to contain all proposals
50                If proposalCount > 0 Then
60                    ReDim proposals(proposalCount - 1) As String
70                Else
80                    ReDim proposals(0) As String
90                End If
                  
                  'Store each guild name
100               For i = 0 To proposalCount - 1
110                   proposals(i) = guilds(.Iterador_ProximaPropuesta(Tipo)).GuildName
120               Next i
130           End With
140       End If
          
150       r_ListaDePropuestas = proposals
End Function

Public Sub a_RechazarAspiranteChar(ByRef Aspirante As String, ByVal guild As Integer, ByRef Detalles As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If InStrB(Aspirante, "\") <> 0 Then
20            Aspirante = Replace(Aspirante, "\", "")
30        End If
40        If InStrB(Aspirante, "/") <> 0 Then
50            Aspirante = Replace(Aspirante, "/", "")
60        End If
70        If InStrB(Aspirante, ".") <> 0 Then
80            Aspirante = Replace(Aspirante, ".", "")
90        End If
          
100       Call guilds(guild).InformarRechazoEnChar(Aspirante, Detalles)
End Sub

Public Function a_ObtenerRechazoDeChar(ByRef Aspirante As String) As String
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If InStrB(Aspirante, "\") <> 0 Then
20            Aspirante = Replace(Aspirante, "\", "")
30        End If
40        If InStrB(Aspirante, "/") <> 0 Then
50            Aspirante = Replace(Aspirante, "/", "")
60        End If
70        If InStrB(Aspirante, ".") <> 0 Then
80            Aspirante = Replace(Aspirante, ".", "")
90        End If
100       a_ObtenerRechazoDeChar = GetVar(CharPath & Aspirante & ".chr", "GUILD", "MotivoRechazo")
110       Call WriteVar(CharPath & Aspirante & ".chr", "GUILD", "MotivoRechazo", vbNullString)
End Function

Public Function a_RechazarAspirante(ByVal Userindex As Integer, ByRef Nombre As String, ByRef refError As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim GI              As Integer
      Dim NroAspirante    As Integer

10        a_RechazarAspirante = False
20        GI = UserList(Userindex).GuildIndex
30        If GI <= 0 Or GI > CANTIDADDECLANES Then
40            refError = "No perteneces a ningún clan"
50            Exit Function
60        End If

70        NroAspirante = guilds(GI).NumeroDeAspirante(Nombre)

80        If NroAspirante = 0 Then
90            refError = Nombre & " no es aspirante a tu clan."
100           Exit Function
110       End If

120       Call guilds(GI).RetirarAspirante(Nombre, NroAspirante)
130       refError = "Fue rechazada tu solicitud de ingreso a " & guilds(GI).GuildName
140       a_RechazarAspirante = True

End Function

Public Function a_DetallesAspirante(ByVal Userindex As Integer, ByRef Nombre As String) As String
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim GI              As Integer
          Dim NroAspirante    As Integer

10        GI = UserList(Userindex).GuildIndex
20        If GI <= 0 Or GI > CANTIDADDECLANES Then
30            Exit Function
40        End If
          
50        If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
60            Exit Function
70        End If
          
80        NroAspirante = guilds(GI).NumeroDeAspirante(Nombre)
90        If NroAspirante > 0 Then
100           a_DetallesAspirante = guilds(GI).DetallesSolicitudAspirante(NroAspirante)
110       End If
          
End Function

Public Sub SendDetallesPersonaje(ByVal Userindex As Integer, ByVal Personaje As String)
       '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim GI          As Integer
          Dim NroAsp      As Integer
          Dim GuildName   As String
          Dim UserFile    As clsIniManager
          Dim Miembro     As String
          Dim GuildActual As Integer
          Dim list()      As String
          Dim i           As Long
          
10        On Error GoTo error
20        GI = UserList(Userindex).GuildIndex
          
30        Personaje = UCase$(Personaje)
          
40        If GI <= 0 Or GI > CANTIDADDECLANES Then
50            Call Protocol.WriteConsoleMsg(Userindex, "No perteneces a ningún clan.", FontTypeNames.FONTTYPE_INFO)
60            Exit Sub
70        End If
          
80        If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
90            Call Protocol.WriteConsoleMsg(Userindex, "No eres el líder de tu clan.", FontTypeNames.FONTTYPE_INFO)
100           Exit Sub
110       End If
          
120       If InStrB(Personaje, "\") <> 0 Then
130           Personaje = Replace$(Personaje, "\", vbNullString)
140       End If
150       If InStrB(Personaje, "/") <> 0 Then
160           Personaje = Replace$(Personaje, "/", vbNullString)
170       End If
180       If InStrB(Personaje, ".") <> 0 Then
190           Personaje = Replace$(Personaje, ".", vbNullString)
200       End If
          
210       NroAsp = guilds(GI).NumeroDeAspirante(Personaje)
          
220       If NroAsp = 0 Then
230           list = guilds(GI).GetMemberList()
              
240           For i = 0 To UBound(list())
250               If Personaje = list(i) Then Exit For
260           Next i
              
270           If i > UBound(list()) Then
280               Call Protocol.WriteConsoleMsg(Userindex, "El personaje no es ni aspirante ni miembro del clan.", FontTypeNames.FONTTYPE_INFO)
290               Exit Sub
300           End If
310       End If
          
320       If Not FileExist(CharPath & Personaje & ".CHR", vbArchive) Then
330           Call guilds(GI).RetirarAspirante(Personaje, NroAsp)
340       End If
          
          'ahora traemos la info
          
350       Set UserFile = New clsIniManager
          
360       With UserFile
370           .Initialize (CharPath & Personaje & ".chr")
              
              ' Get the character's current guild
380           GuildActual = val(.GetValue("GUILD", "GuildIndex"))
390           If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
400               GuildName = "<" & guilds(GuildActual).GuildName & ">"
410           Else
420               GuildName = "Ninguno"
430           End If
              
              'Get previous guilds
440           Miembro = .GetValue("GUILD", "Miembro")
450           If Len(Miembro) > 400 Then
460               Miembro = ".." & Right$(Miembro, 400)
470           End If
              
480           Call Protocol.WriteCharacterInfo(Userindex, Personaje, .GetValue("INIT", "Raza"), .GetValue("INIT", "Clase"), _
                                      .GetValue("INIT", "Genero"), .GetValue("STATS", "ELV"), .GetValue("STATS", "GLD"), _
                                      .GetValue("STATS", "Banco"), .GetValue("REP", "Promedio"), .GetValue("GUILD", "Pedidos"), _
                                      GuildName, Miembro, .GetValue("FACCIONES", "EjercitoReal"), .GetValue("FACCIONES", "EjercitoCaos"), _
                                      .GetValue("FACCIONES", "CiudMatados"), .GetValue("FACCIONES", "CrimMatados"))
490       End With
          
500       Set UserFile = Nothing
          
510       Exit Sub
error:
520       Set UserFile = Nothing
530       If Not (FileExist(CharPath & Personaje & ".chr", vbArchive)) Then
540           Call LogError("El usuario " & UserList(Userindex).Name & " (" & Userindex & _
                          " ) ha pedido los detalles del personaje " & Personaje & " que no se encuentra.")
550       Else
560           Call LogError("[" & Err.Number & "] " & Err.Description & " En la rutina SendDetallesPersonaje, por el usuario " & _
                          UserList(Userindex).Name & " (" & Userindex & " ), pidiendo información sobre el personaje " & Personaje)
570       End If
End Sub

Public Function a_NuevoAspirante(ByVal Userindex As Integer, ByRef clan As String, ByRef Solicitud As String, ByRef refError As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim ViejoSolicitado     As String
      Dim ViejoGuildINdex     As Integer
      Dim ViejoNroAspirante   As Integer
      Dim NuevoGuildIndex     As Integer

10        a_NuevoAspirante = False

20        If UserList(Userindex).GuildIndex > 0 Then
30            refError = "Ya perteneces a un clan, debes salir del mismo antes de solicitar ingresar a otro."
40            Exit Function
50        End If
          
60        If EsNewbie(Userindex) Then
70            refError = "Los newbies no tienen derecho a entrar a un clan."
80            Exit Function
90        End If

100       NuevoGuildIndex = GuildIndex(clan)
110       If NuevoGuildIndex = 0 Then
120           refError = "Ese clan no existe, avise a un administrador."
130           Exit Function
140       End If
          
150       If Not m_EstadoPermiteEntrar(Userindex, NuevoGuildIndex) Then
160           refError = "Tú no puedes entrar a un clan de alineación " & Alineacion2String(guilds(NuevoGuildIndex).Alineacion)
170           Exit Function
180       End If

190       If guilds(NuevoGuildIndex).CantidadAspirantes >= MAXASPIRANTES Then
200           refError = "El clan tiene demasiados aspirantes. Contáctate con un miembro para que procese las solicitudes."
210           Exit Function
220       End If

230       ViejoSolicitado = GetVar(CharPath & UserList(Userindex).Name & ".chr", "GUILD", "ASPIRANTEA")

240       If LenB(ViejoSolicitado) <> 0 Then
              'borramos la vieja solicitud
250           ViejoGuildINdex = CInt(ViejoSolicitado)
260           If ViejoGuildINdex <> 0 Then
270               ViejoNroAspirante = guilds(ViejoGuildINdex).NumeroDeAspirante(UserList(Userindex).Name)
280               If ViejoNroAspirante > 0 Then
290                   Call guilds(ViejoGuildINdex).RetirarAspirante(UserList(Userindex).Name, ViejoNroAspirante)
300               End If
310           Else
                  'RefError = "Inconsistencia en los clanes, avise a un administrador"
                  'Exit Function
320           End If
330       End If
          
340       Call guilds(NuevoGuildIndex).NuevoAspirante(UserList(Userindex).Name, Solicitud)
350       a_NuevoAspirante = True
End Function

Public Function a_AceptarAspirante(ByVal Userindex As Integer, ByRef Aspirante As String, ByRef refError As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim GI              As Integer
      Dim NroAspirante    As Integer
      Dim AspiranteUI     As Integer

          'un pj ingresa al clan :D

10        a_AceptarAspirante = False
          
20        GI = UserList(Userindex).GuildIndex
30        If GI <= 0 Or GI > CANTIDADDECLANES Then
40            refError = "No perteneces a ningún clan"
50            Exit Function
60        End If
          
70        If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
80            refError = "No eres el líder de tu clan"
90            Exit Function
100       End If
          
110       NroAspirante = guilds(GI).NumeroDeAspirante(Aspirante)
          
120       If NroAspirante = 0 Then
130           refError = "El Pj no es aspirante al clan."
140           Exit Function
150       End If
          
160       AspiranteUI = NameIndex(Aspirante)
170       If AspiranteUI > 0 Then
              'pj Online
180           If Not m_EstadoPermiteEntrar(AspiranteUI, GI) Then
190               refError = Aspirante & " no puede entrar a un clan de alineación " & Alineacion2String(guilds(GI).Alineacion)
200               Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
210               Exit Function
220           ElseIf Not UserList(AspiranteUI).GuildIndex = 0 Then
230               refError = Aspirante & " ya es parte de otro clan."
240               Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
250               Exit Function
260           End If
270       Else
280           If Not m_EstadoPermiteEntrarChar(Aspirante, GI) Then
290               refError = Aspirante & " no puede entrar a un clan de alineación " & Alineacion2String(guilds(GI).Alineacion)
300               Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
310               Exit Function
320           ElseIf GetGuildIndexFromChar(Aspirante) Then
330               refError = Aspirante & " ya es parte de otro clan."
340               Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
350               Exit Function
360           End If
370       End If
          'el pj es aspirante al clan y puede entrar
          
380       Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
390       Call guilds(GI).AceptarNuevoMiembro(Aspirante)
          
          ' If player is online, update tag
400       If AspiranteUI > 0 Then
410           Call RefreshCharStatus(AspiranteUI)
420       End If
          
430       a_AceptarAspirante = True
End Function

Public Function GuildName(ByVal GuildIndex As Integer) As String
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
              Exit Function
          
20        GuildName = guilds(GuildIndex).GuildName
End Function

Public Function GuildLeader(ByVal GuildIndex As Integer) As String
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
              Exit Function
          
20        GuildLeader = guilds(GuildIndex).GetLeader
End Function

Public Function GuildAlignment(ByVal GuildIndex As Integer) As String
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
              Exit Function
          
20        GuildAlignment = Alineacion2String(guilds(GuildIndex).Alineacion)
End Function

Public Function GuildFounder(ByVal GuildIndex As Integer) As String
      '***************************************************
      'Autor: ZaMa
      'Returns the guild founder's name
      'Last Modification: 25/03/2009
      '***************************************************
10        If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
              Exit Function
          
20        GuildFounder = guilds(GuildIndex).Fundador
End Function



' Disolución de clanes
Private Function FreeSlotDisolved(ByVal Userindex As Integer) As Byte

          Dim LoopC As Integer
          
10        With UserList(Userindex)
20            For LoopC = 1 To MAX_GUILDS_DISOLVED
30                If .GuildDisolved(LoopC) = 0 Then
40                    FreeSlotDisolved = LoopC
50                    Exit For
60                End If
70            Next LoopC
80        End With
End Function
Public Sub DisolverGuildIndex(ByVal Userindex As Integer)

          Dim MemberList() As String
          Dim LoopC As Integer
          Dim FreeSlot As Byte
          Dim GuildIndex As Integer
          Dim codex() As String
          
10        With UserList(Userindex)
20            If .GuildIndex = 0 Then
30                WriteConsoleMsg Userindex, "¡Tu no tienes clan para disolver!", FontTypeNames.FONTTYPE_INFO
40                Exit Sub
50            End If
              
60            If Not m_EsGuildLeader(UCase$(.Name), .GuildIndex) Then
70                WriteConsoleMsg Userindex, "No eres el lider del clan " & guilds(.GuildIndex).GuildName & ".", FontTypeNames.FONTTYPE_INFO
80                Exit Sub
90            End If
              
100           FreeSlot = FreeSlotDisolved(Userindex)
              
110           If FreeSlot = 0 Then
120               WriteConsoleMsg Userindex, "¡Pero cuantos clanes quieres disolver! Has alcanzado el máximo que permite el servidor", FontTypeNames.FONTTYPE_INFO
130               Exit Sub
140           End If
              
              ' Comienza la disolución del clan
              
150           .GuildDisolved(FreeSlot) = .GuildIndex
              
160           ReDim codex(0 To CANTIDADMAXIMACODEX) As String
              
170           For LoopC = 0 To CANTIDADMAXIMACODEX
180               codex(LoopC) = "Clan disuelto por el lider."
190           Next LoopC
              
200           Call modGuilds.ChangeCodexAndDesc("Clan disuelto por el lider.", codex, .GuildIndex)
              
210           MemberList = guilds(.GuildIndex).GetMemberList()
              
220           For LoopC = LBound(MemberList()) To UBound(MemberList())
230               GuildIndex = modGuilds.m_EcharMiembroDeClan(Userindex, MemberList(LoopC), True)
                  
240               If GuildIndex <> 0 Then
250                   WriteConsoleMsg Userindex, "Personaje expulsado por disolución: " & MemberList(LoopC), FontTypeNames.FONTTYPE_INFO
260               Else
270                   WriteConsoleMsg Userindex, "El personaje " & MemberList(LoopC) & " no ha podido ser expulsado", FontTypeNames.FONTTYPE_INFO
280               End If
290           Next LoopC
                  

300           .GuildIndex = 0
310           WriteVar App.Path & "\CHARFILE\" & UCase$(.Name) & ".chr", "GUILD", "GUILDINDEX", "0"
320           WriteConsoleMsg Userindex, "Has disuelto el clan podrás reanudarlo cuando gustes con el comando /REANUDARCLAN y el TAG del mismo. Recuerda que la reanudación constante de un clan es penada. Hazlo cuando lo necesites.", FontTypeNames.FONTTYPE_INFO
330       End With
End Sub
Private Function SlotGuildDisolved(ByVal Userindex As Integer, ByVal GuildName As String) As Byte
          Dim LoopC As Integer
          
10        With UserList(Userindex)
20            For LoopC = 1 To MAX_GUILDS_DISOLVED
30                If .GuildDisolved(LoopC) > 0 Then
40                    If StrComp(UCase$(guilds(.GuildDisolved(LoopC)).GuildName), GuildName) = 0 Then
50                        SlotGuildDisolved = LoopC
60                        Exit For
70                    End If
80                End If
90            Next LoopC
100       End With
End Function
Public Sub ReanudarGuildIndex(ByVal Userindex As Integer, ByVal GuildName As String)

          Dim GuildIndex As Integer
          Dim LoopC As Integer
          Dim codex() As String
          Dim ErrorMsg As String
          
10        With UserList(Userindex)
20            If .GuildIndex <> 0 Then
30                WriteConsoleMsg Userindex, "¡Para reanudar un clan debes salir del que estás!", FontTypeNames.FONTTYPE_INFO
40                Exit Sub
50            End If
              
60            GuildIndex = SlotGuildDisolved(Userindex, UCase$(GuildName))
              
70            If GuildIndex = 0 Then
80                WriteConsoleMsg Userindex, "No se ha encontrado que hayas disuelto el clan " & GuildName, FontTypeNames.FONTTYPE_INFO
90                Exit Sub
100           End If
              
110           If Not m_EsGuildLeader(UCase$(.Name), .GuildDisolved(GuildIndex)) Then
120               WriteConsoleMsg Userindex, "No eres el lider del clan " & guilds(.GuildDisolved(GuildIndex)).GuildName & ".", FontTypeNames.FONTTYPE_INFO
130               Exit Sub
140           End If
              
150           If .Stats.Gld < 500000 Then
160               WriteConsoleMsg Userindex, "No tienes el dinero suficiente para reanudar tu alianza. 500.000 monedas de oro son las que necesitas", FontTypeNames.FONTTYPE_INFO
170               Exit Sub
180           End If

190           ReDim codex(0 To CANTIDADMAXIMACODEX) As String
200           For LoopC = 0 To CANTIDADMAXIMACODEX
210               codex(LoopC) = "Clan reanudado por el lider."
220           Next LoopC
              
230           Call modGuilds.ChangeCodexAndDesc("Clan reanudado por el lider.", codex, .GuildDisolved(GuildIndex))
              
240           .GuildIndex = .GuildDisolved(GuildIndex)
250           Call guilds(.GuildDisolved(GuildIndex)).AceptarNuevoMiembro(.Name)
260           Call RefreshCharStatus(Userindex)
270           .GuildDisolved(GuildIndex) = 0
280           .Stats.Gld = .Stats.Gld - 500000
290           WriteUpdateGold Userindex
              
300       End With
End Sub
