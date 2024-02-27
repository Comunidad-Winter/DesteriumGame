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
    Userindex   As Integer
    Experiencia As Double
    bPorcentaje As Byte
End Type


Public Function NextParty() As Integer
Dim i As Integer
NextParty = -1
For i = 1 To MAX_PARTIES
    If Parties(i) Is Nothing Then
        NextParty = i
        Exit Function
    End If
Next i
End Function

Public Function PuedeCrearParty(ByVal Userindex As Integer) As Boolean
    PuedeCrearParty = True
'    If UserList(UserIndex).Stats.ELV < MINPARTYLEVEL Then
    
    If CInt(UserList(Userindex).Stats.UserAtributos(eAtributos.Carisma)) * UserList(Userindex).Stats.UserSkills(eSkill.Liderazgo) < 90 Then
        Call WriteConsoleMsg(Userindex, "Tu carisma y liderazgo no son suficientes para liderar una party.", FontTypeNames.FONTTYPE_PARTY)
        PuedeCrearParty = False
    ElseIf UserList(Userindex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(Userindex, "Estás muerto!", FontTypeNames.FONTTYPE_PARTY)
        PuedeCrearParty = False
    End If
End Function

Public Sub CrearParty(ByVal Userindex As Integer)
Dim tInt As Integer
If UserList(Userindex).PartyIndex = 0 Then
    If UserList(Userindex).flags.Muerto = 0 Then
        If UserList(Userindex).Stats.UserSkills(eSkill.Liderazgo) >= 5 Then
            tInt = mdParty.NextParty
            If tInt = -1 Then
                Call WriteConsoleMsg(Userindex, "Por el momento no se pueden crear mas parties", FontTypeNames.FONTTYPE_PARTY)
                Exit Sub
            Else
                Set Parties(tInt) = New clsParty
                If Not Parties(tInt).NuevoMiembro(Userindex) Then
                    Call WriteConsoleMsg(Userindex, "La party está llena, no puedes entrar", FontTypeNames.FONTTYPE_PARTY)
                    Set Parties(tInt) = Nothing
                    Exit Sub
                Else
                    Call WriteConsoleMsg(Userindex, "¡Has formado una party!", FontTypeNames.FONTTYPE_PARTY)
                    UserList(Userindex).PartyIndex = tInt
                    UserList(Userindex).PartySolicitud = 0
                    If Not Parties(tInt).HacerLeader(Userindex) Then
                        Call WriteConsoleMsg(Userindex, "No puedes hacerte líder.", FontTypeNames.FONTTYPE_PARTY)
                    Else
                        Call WriteConsoleMsg(Userindex, "¡ Te has convertido en líder de la party !", FontTypeNames.FONTTYPE_PARTY)
                    End If
                End If
            End If
        Else
            Call WriteConsoleMsg(Userindex, " No tienes suficientes puntos de liderazgo para liderar una party.", FontTypeNames.FONTTYPE_PARTY)
        End If
    Else
        Call WriteConsoleMsg(Userindex, "Estás muerto!", FontTypeNames.FONTTYPE_PARTY)
    End If
Else
    Call WriteConsoleMsg(Userindex, " Ya perteneces a una party.", FontTypeNames.FONTTYPE_PARTY)
End If
End Sub

Public Sub SolicitarIngresoAParty(ByVal Userindex As Integer)
'ESTO ES enviado por el PJ para solicitar el ingreso a la party
Dim tInt As Integer

    If UserList(Userindex).PartyIndex > 0 Then
        'si ya esta en una party
        Call WriteConsoleMsg(Userindex, "Ya perteneces a una party, escribe /SALIRPARTY para abandonarla", FontTypeNames.FONTTYPE_PARTY)
        UserList(Userindex).PartySolicitud = 0
        Exit Sub
    End If
    If UserList(Userindex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(Userindex, "¡Estás muerto!", FontTypeNames.FONTTYPE_INFO)
        UserList(Userindex).PartySolicitud = 0
        Exit Sub
    End If
    tInt = UserList(Userindex).flags.TargetUser
    If tInt > 0 Then
        If UserList(tInt).PartyIndex > 0 Then
            UserList(Userindex).PartySolicitud = UserList(tInt).PartyIndex
            Call WriteConsoleMsg(Userindex, "El fundador decidirá si te acepta en la party", FontTypeNames.FONTTYPE_PARTY)
            WriteConsoleMsg tInt, "El personaje " & UserList(Userindex).name & " quiere ingresar en la party.", FontTypeNames.FONTTYPE_PARTY
        Else
            Call WriteConsoleMsg(Userindex, UserList(tInt).name & " no es fundador de ninguna party.", FontTypeNames.FONTTYPE_INFO)
            UserList(Userindex).PartySolicitud = 0
            Exit Sub
        End If
    Else
        Call WriteConsoleMsg(Userindex, "Para ingresar a una party debes hacer click sobre el fundador y apretar F3.", FontTypeNames.FONTTYPE_PARTY)
        UserList(Userindex).PartySolicitud = 0
    End If
End Sub

Public Sub SalirDeParty(ByVal Userindex As Integer)
Dim PI As Integer
PI = UserList(Userindex).PartyIndex
If PI > 0 Then
    If Parties(PI).SaleMiembro(Userindex) Then
        'sale el leader
        Set Parties(PI) = Nothing
    Else
        UserList(Userindex).PartyIndex = 0
    End If
Else
    Call WriteConsoleMsg(Userindex, "No eres miembro de ninguna party.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Public Sub ExpulsarDeParty(ByVal leader As Integer, ByVal OldMember As Integer)
Dim PI As Integer
PI = UserList(leader).PartyIndex

If PI = UserList(OldMember).PartyIndex Then
    If Parties(PI).SaleMiembro(OldMember) Then
        'si la funcion me da true, entonces la party se disolvio
        'y los partyindex fueron reseteados a 0
        Set Parties(PI) = Nothing
    Else
        UserList(OldMember).PartyIndex = 0
    End If
Else
    Call WriteConsoleMsg(leader, LCase(UserList(OldMember).name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Public Function puedeCambiarPorcentajes(ByVal user_Index As Integer, ByRef errorString As String) As Boolean

    '
    ' @ maTih.-
    
    puedeCambiarPorcentajes = False
    
    With UserList(user_Index)
    
         If (.PartyIndex = 0) Then
             errorString = "No eres miembro de ninguna party"
             Exit Function
         End If
         
         If (Parties(.PartyIndex).EsPartyLeader(user_Index) = False) Then
            errorString = "No eres el lider de tu party."
            Exit Function
         End If
         
         If (Parties(.PartyIndex).CantMiembros = 1) Then
            errorString = "Estás solo en la party."
            Exit Function
         End If
         
         puedeCambiarPorcentajes = True
         
    End With
    
End Function

Public Function getPartyString(ByVal rUserIndex As Integer)

    '
    ' @ maTih.-
    
    Dim pIndex As Integer
    
    pIndex = UserList(rUserIndex).PartyIndex
    
    If Parties(pIndex).EsPartyLeader(rUserIndex) = False Then
       getPartyString = "nada"
       Exit Function
    End If
    
    getPartyString = Parties(pIndex).preparePorcentajeString()

End Function

Public Function validarNuevosPorcentajes(ByVal leaderIndex As Integer, ByRef bArray() As Byte, ByRef errorStr As String) As Boolean

    '
    ' @ maTih.-
    
    validarNuevosPorcentajes = False
    
    With UserList(leaderIndex)
2
         Dim j As Long
         Dim t As Long
         Dim m As Integer
         
         m = .Stats.UserSkills(eSkill.Liderazgo)
         
         For j = 1 To UBound(bArray())
             If (bArray(j) > 0) Then
                 t = t + bArray(j)
                 
                 If (bArray(j) > m) Then
                    errorStr = "No tienes tantos skills en liderazgo."
                    Exit Function
                 End If
                 
                 If (t > 100) Then
                    errorStr = "La suma de los porcentajes exede el máximo (100)"
                    Exit Function
                 End If
                 
             End If
         Next j
         
         validarNuevosPorcentajes = (t = 100)
         
         
    End With
    
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
    
    PI = UserList(User).PartyIndex
    
    If PI > 0 Then
        If Parties(PI).EsPartyLeader(User) Then
            UserPuedeEjecutarComandos = True
        Else
            Call WriteConsoleMsg(User, "¡No eres el líder de tu Party!", FontTypeNames.FONTTYPE_PARTY)
            Exit Function
        End If
    Else
        Call WriteConsoleMsg(User, "No eres miembro de ninguna party.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
End Function

Public Sub AprobarIngresoAParty(ByVal leader As Integer, ByVal NewMember As Integer)
'el UI es el leader
Dim PI As Integer
Dim razon As String

PI = UserList(leader).PartyIndex

If UserList(NewMember).PartySolicitud = PI Then
    If Not UserList(NewMember).flags.Muerto = 1 Then
        If UserList(NewMember).PartyIndex = 0 Then
        EnvioNewMember = UserList(NewMember).PartyIndex
            If Parties(PI).PuedeEntrar(NewMember, razon) Then
                If Parties(PI).NuevoMiembro(NewMember) Then
                    Call Parties(PI).MandarMensajeAConsola(UserList(leader).name & " ha aceptado a " & UserList(NewMember).name & " en la party.", "Servidor")
                    UserList(NewMember).PartyIndex = PI
                    UserList(NewMember).PartySolicitud = 0
                Else
                    'no pudo entrar
                    'ACA UNO PUEDE CODIFICAR OTRO TIPO DE ERRORES...
                    Call SendData(SendTarget.ToAdmins, leader, PrepareMessageConsoleMsg(" Servidor> CATASTROFE EN PARTIES, NUEVOMIEMBRO DIO FALSE! :S ", FontTypeNames.FONTTYPE_PARTY))
                    End If
                Else
                'no debe entrar
                Call WriteConsoleMsg(leader, razon, FontTypeNames.FONTTYPE_PARTY)
            End If
        Else
            Call WriteConsoleMsg(leader, UserList(NewMember).name & " ya es miembro de otra party.", FontTypeNames.FONTTYPE_PARTY)
            Exit Sub
        End If
    Else
        Call WriteConsoleMsg(leader, "¡Está muerto, no puedes aceptar miembros en ese estado!", FontTypeNames.FONTTYPE_PARTY)
        Exit Sub
    End If
Else
    Call WriteConsoleMsg(leader, LCase(UserList(NewMember).name) & " no ha solicitado ingresar a tu party.", FontTypeNames.FONTTYPE_PARTY)
    Exit Sub
End If

End Sub

Public Sub BroadCastParty(ByVal Userindex As Integer, ByRef texto As String)
Dim PI As Integer
    
    PI = UserList(Userindex).PartyIndex
    
    If PI > 0 Then
        Call Parties(PI).MandarMensajeAConsola(texto, UserList(Userindex).name)
    End If

End Sub

Public Sub OnlineParty(ByVal Userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 11/27/09 (Budi)
'Adapte la función a los nuevos métodos de clsParty
'*************************************************
Dim i As Integer
Dim PI As Integer
Dim Text As String
Dim MembersOnline(1 To PARTY_MAXMEMBERS) As Integer
    PI = UserList(Userindex).PartyIndex
    
    If PI > 0 Then
        Call Parties(PI).ObtenerMiembrosOnline(MembersOnline)
        Text = "Nombre(Exp): "
        For i = 1 To PARTY_MAXMEMBERS
            If MembersOnline(i) > 0 Then
                Text = Text & " - " & UserList(MembersOnline(i)).name & " (" & Fix(Parties(PI).MiExperiencia(MembersOnline(i))) & ")"
            End If
        Next i
        Text = Text & ". Experiencia total: " & Parties(PI).ObtenerExperienciaTotal
        Call WriteConsoleMsg(Userindex, Text, FontTypeNames.FONTTYPE_PARTY)
    End If
    
End Sub


Public Sub TransformarEnLider(ByVal OldLeader As Integer, ByVal NewLeader As Integer)
Dim PI As Integer

If OldLeader = NewLeader Then Exit Sub

PI = UserList(OldLeader).PartyIndex

If PI = UserList(NewLeader).PartyIndex Then
    If UserList(NewLeader).flags.Muerto = 0 Then
        If Parties(PI).HacerLeader(NewLeader) Then
            Call Parties(PI).MandarMensajeAConsola("El nuevo líder de la party es " & UserList(NewLeader).name, UserList(OldLeader).name)
        Else
            Call WriteConsoleMsg(OldLeader, "¡No se ha hecho el cambio de mando!", FontTypeNames.FONTTYPE_PARTY)
        End If
    Else
        Call WriteConsoleMsg(OldLeader, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(OldLeader, LCase(UserList(NewLeader).name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub


Public Sub ActualizaExperiencias()
'esta funcion se invoca antes de worlsaves, y apagar servidores
'en caso que la experiencia sea acumulada y no por golpe
'para que grabe los datos en los charfiles
Dim i As Integer

If Not PARTY_EXPERIENCIAPORGOLPE Then
    
    haciendoBK = True
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Distribuyendo experiencia en parties.", FontTypeNames.FONTTYPE_SERVER))
    For i = 1 To MAX_PARTIES
        If Not Parties(i) Is Nothing Then
            Call Parties(i).FlushExperiencia
        End If
    Next i
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Experiencia distribuida.", FontTypeNames.FONTTYPE_SERVER))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    haciendoBK = False

End If

End Sub

Public Sub ObtenerExito(ByVal Userindex As Integer, ByVal Exp As Long, mapa As Integer, X As Integer, Y As Integer)
    If Exp <= 0 Then
        If Not CASTIGOS Then Exit Sub
    End If
    
    Call Parties(UserList(Userindex).PartyIndex).ObtenerExito(Exp, mapa, X, Y)

End Sub

Public Function CantMiembros(ByVal Userindex As Integer) As Integer
CantMiembros = 0
If UserList(Userindex).PartyIndex > 0 Then
    CantMiembros = Parties(UserList(Userindex).PartyIndex).CantMiembros
End If

End Function

''
' Sets the new p_sumaniveleselevados to the party.
'
' @param UserInidex Specifies reference to user
' @remarks When a user level up and he is in a party, we call this sub to don't desestabilice the party exp formula
Public Sub ActualizarSumaNivelesElevados(ByVal Userindex As Integer)
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 28/10/08
'
'*************************************************
    If UserList(Userindex).PartyIndex > 0 Then
        Call Parties(UserList(Userindex).PartyIndex).UpdateSumaNivelesElevados(UserList(Userindex).Stats.ELV)
    End If
End Sub




