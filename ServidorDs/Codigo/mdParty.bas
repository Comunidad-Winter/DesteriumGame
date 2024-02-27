Attribute VB_Name = "mdParty"
'**************************************************************
' mdParty.bas - Library of functions to manipulate parties.
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
' SOPORTES PARA LAS PARTIES
' (Ver este modulo como una clase abstracta "PartyManager")
'


''
'cantidad maxima de parties en el servidor
Public Const MAX_PARTIES As Integer = 300

''
'nivel minimo para crear party
Public Const MINPARTYLEVEL As Byte = 15

''
'Cantidad maxima de gente en la party
Public Const PARTY_MAXMEMBERS As Byte = 5

''
'Si esto esta en True, la exp sale por cada golpe que le da
'Si no, la exp la recibe al salirse de la party (pq las partys, floodean)
Public Const PARTY_EXPERIENCIAPORGOLPE As Boolean = False

''
'maxima diferencia de niveles permitida en una party
Public Const MAXPARTYDELTALEVEL As Byte = 7

''
'distancia al leader para que este acepte el ingreso
Public Const MAXDISTANCIAINGRESOPARTY As Byte = 2

''
'maxima distancia a un exito para obtener su experiencia
Public Const PARTY_MAXDISTANCIA As Byte = 18

''
'restan las muertes de los miembros?
Public Const CASTIGOS As Boolean = False

''
'Numero al que elevamos el nivel de cada miembro de la party
'Esto es usado para calcular la distribución de la experiencia entre los miembros
'Se lee del archivo de balance
Public ExponenteNivelParty As Single

''
'tPartyMember
'
' @param UserIndex UserIndex
' @param Experiencia Experiencia
'
Public Type tPartyMember
    UserIndex   As Integer
    Experiencia As Double
    bPorcentaje As Byte
End Type


Public Function NextParty() As Integer
      Dim i As Integer
10    NextParty = -1
20    For i = 1 To MAX_PARTIES
30        If Parties(i) Is Nothing Then
40            NextParty = i
50            Exit Function
60        End If
70    Next i
End Function

Public Function PuedeCrearParty(ByVal UserIndex As Integer) As Boolean
10        PuedeCrearParty = True
      '    If UserList(UserIndex).Stats.ELV < MINPARTYLEVEL Then
          
20        If CInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma)) * UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) < 90 Then
30            Call WriteConsoleMsg(UserIndex, "Tu carisma y liderazgo no son suficientes para liderar una party.", FontTypeNames.FONTTYPE_PARTY)
40            PuedeCrearParty = False
50        ElseIf UserList(UserIndex).flags.Muerto = 1 Then
60            Call WriteConsoleMsg(UserIndex, "Estás muerto!", FontTypeNames.FONTTYPE_PARTY)
70            PuedeCrearParty = False
80        End If
End Function

Public Sub CrearParty(ByVal UserIndex As Integer)
      Dim tInt As Integer
10    If UserList(UserIndex).PartyIndex = 0 Then
20        If UserList(UserIndex).flags.Muerto = 0 Then
30            If UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) >= 5 Then
40                tInt = mdParty.NextParty
50                If tInt = -1 Then
60                    Call WriteConsoleMsg(UserIndex, "Por el momento no se pueden crear mas parties", FontTypeNames.FONTTYPE_PARTY)
70                    Exit Sub
80                Else
90                    Set Parties(tInt) = New clsParty
100                   If Not Parties(tInt).NuevoMiembro(UserIndex) Then
110                       Call WriteConsoleMsg(UserIndex, "La party está llena, no puedes entrar", FontTypeNames.FONTTYPE_PARTY)
120                       Set Parties(tInt) = Nothing
130                       Exit Sub
140                   Else
150                       Call WriteConsoleMsg(UserIndex, "¡Has formado una party!", FontTypeNames.FONTTYPE_PARTY)
160                       UserList(UserIndex).PartyIndex = tInt
170                       UserList(UserIndex).PartySolicitud = 0
180                       If Not Parties(tInt).HacerLeader(UserIndex) Then
190                           Call WriteConsoleMsg(UserIndex, "No puedes hacerte líder.", FontTypeNames.FONTTYPE_PARTY)
200                       Else
210                           Call WriteConsoleMsg(UserIndex, "¡ Te has convertido en líder de la party !", FontTypeNames.FONTTYPE_PARTY)
220                       End If
230                   End If
240               End If
250           Else
260               Call WriteConsoleMsg(UserIndex, " No tienes suficientes puntos de liderazgo para liderar una party.", FontTypeNames.FONTTYPE_PARTY)
270           End If
280       Else
290           Call WriteConsoleMsg(UserIndex, "Estás muerto!", FontTypeNames.FONTTYPE_PARTY)
300       End If
310   Else
320       Call WriteConsoleMsg(UserIndex, " Ya perteneces a una party.", FontTypeNames.FONTTYPE_PARTY)
330   End If
End Sub

Public Sub SolicitarIngresoAParty(ByVal UserIndex As Integer)
      'ESTO ES enviado por el PJ para solicitar el ingreso a la party
      Dim tInt As Integer

10        If UserList(UserIndex).PartyIndex > 0 Then
              'si ya esta en una party
20            Call WriteConsoleMsg(UserIndex, "Ya perteneces a una party, escribe /SALIRPARTY para abandonarla", FontTypeNames.FONTTYPE_PARTY)
30            UserList(UserIndex).PartySolicitud = 0
40            Exit Sub
50        End If
60        If UserList(UserIndex).flags.Muerto = 1 Then
70            Call WriteConsoleMsg(UserIndex, "¡Estás muerto!", FontTypeNames.FONTTYPE_INFO)
80            UserList(UserIndex).PartySolicitud = 0
90            Exit Sub
100       End If

110       tInt = UserList(UserIndex).flags.TargetUser
120       If tInt > 0 Then
130           If UserList(tInt).PartyIndex > 0 Then
140               UserList(UserIndex).PartySolicitud = UserList(tInt).PartyIndex
150               Call WriteConsoleMsg(UserIndex, "El fundador decidirá si te acepta en la party", FontTypeNames.FONTTYPE_PARTY)
160               WriteConsoleMsg tInt, "El personaje " & UserList(UserIndex).Name & " quiere ingresar en la party.", FontTypeNames.FONTTYPE_PARTY
170           Else
180               Call WriteConsoleMsg(UserIndex, UserList(tInt).Name & " no es fundador de ninguna party.", FontTypeNames.FONTTYPE_INFO)
190               UserList(UserIndex).PartySolicitud = 0
200               Exit Sub
210           End If
220       Else
230           Call WriteConsoleMsg(UserIndex, "Para ingresar a una party debes hacer click sobre el fundador y apretar F3.", FontTypeNames.FONTTYPE_PARTY)
240           UserList(UserIndex).PartySolicitud = 0
250       End If
End Sub

Public Sub SalirDeParty(ByVal UserIndex As Integer)
      Dim PI As Integer
10    PI = UserList(UserIndex).PartyIndex
20    If PI > 0 Then
30        If Parties(PI).SaleMiembro(UserIndex) Then
              'sale el leader
40            Set Parties(PI) = Nothing
50        Else
60            UserList(UserIndex).PartyIndex = 0
70        End If
80    Else
90        Call WriteConsoleMsg(UserIndex, "No eres miembro de ninguna party.", FontTypeNames.FONTTYPE_INFO)
100   End If

End Sub

Public Sub ExpulsarDeParty(ByVal leader As Integer, ByVal OldMember As Integer)
      Dim PI As Integer
10    PI = UserList(leader).PartyIndex

20    If PI = UserList(OldMember).PartyIndex Then
30        If Parties(PI).SaleMiembro(OldMember) Then
              'si la funcion me da true, entonces la party se disolvio
              'y los partyindex fueron reseteados a 0
40            Set Parties(PI) = Nothing
50        Else
60            UserList(OldMember).PartyIndex = 0
70        End If
80    Else
90        Call WriteConsoleMsg(leader, LCase(UserList(OldMember).Name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
100   End If

End Sub

Public Function puedeCambiarPorcentajes(ByVal user_Index As Integer, ByRef errorString As String) As Boolean

          '
          ' @ maTih.-
          
10        puedeCambiarPorcentajes = False
          
20        With UserList(user_Index)
          
30             If (.PartyIndex = 0) Then
40                 errorString = "No eres miembro de ninguna party"
50                 Exit Function
60             End If
               
70             If (Parties(.PartyIndex).EsPartyLeader(user_Index) = False) Then
80                errorString = "No eres el lider de tu party."
90                Exit Function
100            End If
               
110            If (Parties(.PartyIndex).CantMiembros = 1) Then
120               errorString = "Estás solo en la party."
130               Exit Function
140            End If
               
150            puedeCambiarPorcentajes = True
               
160       End With
          
End Function

Public Function getPartyString(ByVal rUserIndex As Integer)

          '
          ' @ maTih.-
          
          Dim pIndex As Integer
          
10        pIndex = UserList(rUserIndex).PartyIndex
          
20        If Parties(pIndex).EsPartyLeader(rUserIndex) = False Then
30           getPartyString = "nada"
40           Exit Function
50        End If
          
60        getPartyString = Parties(pIndex).preparePorcentajeString()

End Function

Public Function validarNuevosPorcentajes(ByVal leaderIndex As Integer, ByRef bArray() As Byte, ByRef errorStr As String) As Boolean

          '
          ' @ maTih.-
          
10        validarNuevosPorcentajes = False
          
20        With UserList(leaderIndex)

               Dim j As Long
               Dim t As Long
               Dim m As Integer
               
30             m = .Stats.UserSkills(eSkill.Liderazgo)
               
40             If m > 90 Then m = 90
               
50             For j = 1 To UBound(bArray())
60                 If (bArray(j) > 0) Then
70                     t = t + bArray(j)
                       
80                     If (bArray(j) > m) Then
90                        errorStr = "No tienes tantos skills en liderazgo."
100                       Exit Function
110                    End If
                       
120                    If (t > 100) Then
130                       errorStr = "La suma de los porcentajes exede el máximo (100)"
140                       Exit Function
150                    End If
                       
160                End If
170            Next j
               
180            validarNuevosPorcentajes = (t = 100)
               
               
190       End With
          
End Function

''
' Determines if a user can use party commands like /acceptparty or not.
'
' @param User Specifies reference to user
' @return  True if the user can use party commands, false if not.
Public Function UserPuedeEjecutarComandos(ByVal User As Integer) As Boolean
      '*************************************************
      'Author: Marco Vanotti(Marco)
      'Last modified: 05/05/09
      '
      '*************************************************
          Dim PI As Integer
          
10        PI = UserList(User).PartyIndex
          
20        If PI > 0 Then
30            If Parties(PI).EsPartyLeader(User) Then
40                UserPuedeEjecutarComandos = True
50            Else
60                Call WriteConsoleMsg(User, "¡No eres el líder de tu Party!", FontTypeNames.FONTTYPE_PARTY)
70                Exit Function
80            End If
90        Else
100           Call WriteConsoleMsg(User, "No eres miembro de ninguna party.", FontTypeNames.FONTTYPE_INFO)
110           Exit Function
120       End If
End Function

Public Sub AprobarIngresoAParty(ByVal leader As Integer, ByVal NewMember As Integer)
      'el UI es el leader
      Dim PI As Integer
      Dim razon As String

10    PI = UserList(leader).PartyIndex

20    If UserList(NewMember).PartySolicitud = PI Then
30        If Not UserList(NewMember).flags.Muerto = 1 Then
40            If UserList(NewMember).PartyIndex = 0 Then
50            EnvioNewMember = UserList(NewMember).PartyIndex
60                If Parties(PI).PuedeEntrar(NewMember, razon) Then
70                    If Parties(PI).NuevoMiembro(NewMember) Then
80                        Call Parties(PI).MandarMensajeAConsola(UserList(leader).Name & " ha aceptado a " & UserList(NewMember).Name & " en la party.", "Servidor")
90                        UserList(NewMember).PartyIndex = PI
100                       UserList(NewMember).PartySolicitud = 0
110                   Else
                          'no pudo entrar
                          'ACA UNO PUEDE CODIFICAR OTRO TIPO DE ERRORES...
120                       Call SendData(SendTarget.ToAdmins, leader, PrepareMessageConsoleMsg(" Servidor> CATASTROFE EN PARTIES, NUEVOMIEMBRO DIO FALSE! :S ", FontTypeNames.FONTTYPE_PARTY))
130                       End If
140                   Else
                      'no debe entrar
150                   Call WriteConsoleMsg(leader, razon, FontTypeNames.FONTTYPE_PARTY)
160               End If
170           Else
180               Call WriteConsoleMsg(leader, UserList(NewMember).Name & " ya es miembro de otra party.", FontTypeNames.FONTTYPE_PARTY)
190               Exit Sub
200           End If
210       Else
220           Call WriteConsoleMsg(leader, "¡Está muerto, no puedes aceptar miembros en ese estado!", FontTypeNames.FONTTYPE_PARTY)
230           Exit Sub
240       End If
250   Else
260       Call WriteConsoleMsg(leader, LCase(UserList(NewMember).Name) & " no ha solicitado ingresar a tu party.", FontTypeNames.FONTTYPE_PARTY)
270       Exit Sub
280   End If

End Sub

Public Sub BroadCastParty(ByVal UserIndex As Integer, ByRef texto As String)
      Dim PI As Integer
          
10        PI = UserList(UserIndex).PartyIndex
          
20        If PI > 0 Then
30            Call Parties(PI).MandarMensajeAConsola(texto, UserList(UserIndex).Name)
40        End If

End Sub

Public Sub OnlineParty(ByVal UserIndex As Integer)
      '*************************************************
      'Author: Unknown
      'Last modified: 11/27/09 (Budi)
      'Adapte la función a los nuevos métodos de clsParty
      '*************************************************
      Dim i As Integer
      Dim PI As Integer
      Dim Text As String
      Dim MembersOnline(1 To PARTY_MAXMEMBERS) As Integer
10        PI = UserList(UserIndex).PartyIndex
          
20        If PI > 0 Then
30            Call Parties(PI).ObtenerMiembrosOnline(MembersOnline)
40            Text = "Nombre(Exp): "
50            For i = 1 To PARTY_MAXMEMBERS
60                If MembersOnline(i) > 0 Then
70                    Text = Text & " - " & UserList(MembersOnline(i)).Name & " (" & Fix(Parties(PI).MiExperiencia(MembersOnline(i))) & ")"
80                End If
90            Next i
100           Text = Text & ". Experiencia total: " & Parties(PI).ObtenerExperienciaTotal
110           Call WriteConsoleMsg(UserIndex, Text, FontTypeNames.FONTTYPE_PARTY)
120       End If
          
End Sub


Public Sub TransformarEnLider(ByVal OldLeader As Integer, ByVal NewLeader As Integer)
      Dim PI As Integer

10    If OldLeader = NewLeader Then Exit Sub

20    PI = UserList(OldLeader).PartyIndex

30    If PI = UserList(NewLeader).PartyIndex Then
40        If UserList(NewLeader).flags.Muerto = 0 Then
50            If Parties(PI).HacerLeader(NewLeader) Then
60                Call Parties(PI).MandarMensajeAConsola("El nuevo líder de la party es " & UserList(NewLeader).Name, UserList(OldLeader).Name)
70            Else
80                Call WriteConsoleMsg(OldLeader, "¡No se ha hecho el cambio de mando!", FontTypeNames.FONTTYPE_PARTY)
90            End If
100       Else
110           Call WriteConsoleMsg(OldLeader, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
120       End If
130   Else
140       Call WriteConsoleMsg(OldLeader, LCase(UserList(NewLeader).Name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
150   End If

End Sub


Public Sub ActualizaExperiencias()
      'esta funcion se invoca antes de worlsaves, y apagar servidores
      'en caso que la experiencia sea acumulada y no por golpe
      'para que grabe los datos en los charfiles
      Dim i As Integer

10    If Not PARTY_EXPERIENCIAPORGOLPE Then
          
20        haciendoBK = True
30        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
          
40        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Distribuyendo experiencia en parties.", FontTypeNames.FONTTYPE_SERVER))
50        For i = 1 To MAX_PARTIES
60            If Not Parties(i) Is Nothing Then
70                Call Parties(i).FlushExperiencia
80            End If
90        Next i
100       Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Experiencia distribuida.", FontTypeNames.FONTTYPE_SERVER))
110       Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
120       haciendoBK = False

130   End If

End Sub

Public Sub ObtenerExito(ByVal UserIndex As Integer, ByVal Exp As Long, Mapa As Integer, X As Integer, Y As Integer)
10        If Exp <= 0 Then
20            If Not CASTIGOS Then Exit Sub
30        End If
          
40        Call Parties(UserList(UserIndex).PartyIndex).ObtenerExito(Exp, Mapa, X, Y)

End Sub

Public Function CantMiembros(ByVal UserIndex As Integer) As Integer
10    CantMiembros = 0
20    If UserList(UserIndex).PartyIndex > 0 Then
30        CantMiembros = Parties(UserList(UserIndex).PartyIndex).CantMiembros
40    End If

End Function

''
' Sets the new p_sumaniveleselevados to the party.
'
' @param UserInidex Specifies reference to user
' @remarks When a user level up and he is in a party, we call this sub to don't desestabilice the party exp formula
Public Sub ActualizarSumaNivelesElevados(ByVal UserIndex As Integer)
      '*************************************************
      'Author: Marco Vanotti (MarKoxX)
      'Last modified: 28/10/08
      '
      '*************************************************
10        If UserList(UserIndex).PartyIndex > 0 Then
20            Call Parties(UserList(UserIndex).PartyIndex).UpdateSumaNivelesElevados(UserList(UserIndex).Stats.ELV)
30        End If
End Sub




