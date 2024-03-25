Attribute VB_Name = "Extra"
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

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo EsNewbie_Err

    '</EhHeader>

    EsNewbie = UserList(UserIndex).Stats.Elv <= LimiteNewbie
    '<EhFooter>
    Exit Function

EsNewbie_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.EsNewbie " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function esArmada(ByVal UserIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo esArmada_Err

    '</EhHeader>
    esArmada = (UserList(UserIndex).Faction.Status = r_Armada)
    '<EhFooter>
    Exit Function

esArmada_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.esArmada " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function esCaos(ByVal UserIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo esCaos_Err

    '</EhHeader>
    esCaos = (UserList(UserIndex).Faction.Status = r_Caos)
    '<EhFooter>
    Exit Function

esCaos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.esCaos " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function EsGm(ByVal UserIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo EsGm_Err

    '</EhHeader>

    EsGm = (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios))
    '<EhFooter>
    Exit Function

EsGm_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.EsGm " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function EsGmDios(ByVal UserIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo EsGmDios_Err

    '</EhHeader>
    
    EsGmDios = (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios))
    '<EhFooter>
    Exit Function

EsGmDios_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.EsGmDios " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function EsGmPriv(ByVal UserIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo EsGmPriv_Err

    '</EhHeader>
    
    EsGmPriv = CharIs_Admin(UCase$(UserList(UserIndex).Name))
    'EsGmPriv = (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin))
    '<EhFooter>
    Exit Function

EsGmPriv_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.EsGmPriv " & "at line " & Erl
        
    '</EhFooter>
End Function

Function GetDayString(AccessDay As Byte) As String

    Select Case AccessDay

        Case 1: GetDayString = "Lunes"

        Case 2: GetDayString = "Martes"

        Case 3: GetDayString = "Miércoles"

        Case 4: GetDayString = "Jueves"

        Case 5: GetDayString = "Viernes"

        Case 6: GetDayString = "Sábado"

        Case 7: GetDayString = "Domingo"

        Case Else: GetDayString = "día inválido"

    End Select

End Function

' # Chequea los online del servidor para determinar si esta disponible
Public Function CheckMap_Onlines(ByVal UserIndex As Integer, _
                                 ByRef DestPos As WorldPos) As Boolean

    If MapInfo(DestPos.Map).MinOns > 0 Then
        If NumUsers + UsersBot < MapInfo(DestPos.Map).MinOns Then
            Exit Function

        End If

    End If
    
    CheckMap_Onlines = True

End Function

' # Chequea el horario del mapa para poder ingresar. Puede habilitarse varios días en diferentes horarios.
Public Function CheckMap_HourDay(ByVal UserIndex As Integer, _
                                 ByRef DestPos As WorldPos) As Boolean

    On Error GoTo ErrHandler

    Dim currDay             As Integer

    Dim currTime            As Integer

    Dim i                   As Integer

    Dim availabilityMessage As String

    Dim accessInfo          As String
    
    If MapInfo(DestPos.Map).AccessDays(0) = 0 Then
        CheckMap_HourDay = True
        Exit Function

    End If
    
    currDay = Weekday(Date) ' 1=domingo, 2=lunes, 3 martes, 4 miercoles , 5 jueves, 6 viernes..., 7=sábado
    currTime = Format(Time, "HHMM")
    availabilityMessage = "El mapa se ha restringido. Disponible el día: " & vbCrLf
    
    For i = LBound(MapInfo(DestPos.Map).AccessDays) To UBound(MapInfo(DestPos.Map).AccessDays)
    
        accessInfo = GetDayString(MapInfo(DestPos.Map).AccessDays(i)) & " desde las " & Format(MapInfo(DestPos.Map).AccessTimeStarts(i), "00:00") & " hasta las " & Format(MapInfo(DestPos.Map).accessTimeEnds(i), "00:00") & ". "
        availabilityMessage = availabilityMessage & accessInfo & vbCrLf
        
        If MapInfo(DestPos.Map).AccessDays(i) = currDay Then
            If currTime >= MapInfo(DestPos.Map).AccessTimeStarts(i) And currTime <= MapInfo(DestPos.Map).accessTimeEnds(i) Then
                CheckMap_HourDay = True
                Exit Function

            End If

        End If

    Next i

    ' Si llegamos aquí, el acceso está denegado
    Call WriteConsoleMsg(UserIndex, availabilityMessage, FontTypeNames.FONTTYPE_INFO)
    CheckMap_HourDay = False
    
    Exit Function
ErrHandler:
    CheckMap_HourDay = False

End Function

Public Sub DoTileEvents(ByVal UserIndex As Integer, _
                        ByVal Map As Integer, _
                        ByVal X As Integer, _
                        ByVal Y As Integer)

    '***************************************************
    'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 06/03/2010
    'Handles the Map passage of Users. Allows the existance
    'of exclusive maps for Newbies, Royal Army and Caos Legion members
    'and enables GMs to enter every map without restriction.
    'Uses: Mapinfo(map).Restringir = "NEWBIE" (newbies), "ARMADA", "CAOS", "FACCION" or "NO".
    ' 06/03/2010 : Now we have 5 attemps to not fall into a map change or another teleport while going into a teleport. (Marco)
    '***************************************************

    Dim nPos       As WorldPos

    Dim FxFlag     As Boolean

    Dim TelepRadio As Integer

    Dim DestPos    As WorldPos

    'Controla las salidas
    If InMapBounds(Map, X, Y) Then

        With MapData(Map, X, Y)

            If .ObjInfo.ObjIndex > 0 Then
                FxFlag = ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport
                TelepRadio = ObjData(.ObjInfo.ObjIndex).Radio

            End If
            
            If .TileExit.Map > 0 And .TileExit.Map <= NumMaps Then
                
                ' Es un teleport, entra en una posicion random, acorde al radio (si es 0, es pos fija)
                ' We have 5 attempts to not falling into another teleport or a map exit.. If we get to the fifth attemp,
                ' the teleport will act as if its radius = 0.
                
                If FxFlag And TelepRadio > 0 Then

                    Dim attemps As Long

                    Dim exitMap As Boolean

                    Do
                        DestPos.X = .TileExit.X + RandomNumber(TelepRadio * (-1), TelepRadio)
                        DestPos.Y = .TileExit.Y + RandomNumber(TelepRadio * (-1), TelepRadio)
                        
                        attemps = attemps + 1
                        
                        exitMap = MapData(.TileExit.Map, DestPos.X, DestPos.Y).TileExit.Map > 0 And MapData(.TileExit.Map, DestPos.X, DestPos.Y).TileExit.Map <= NumMaps
                    Loop Until (attemps >= 5 Or exitMap = False)
                    
                    If attemps >= 5 Then
                        DestPos.X = .TileExit.X
                        DestPos.Y = .TileExit.Y

                    End If

                    ' Posicion fija
                Else
                    DestPos.X = .TileExit.X
                    DestPos.Y = .TileExit.Y

                End If
                
                DestPos.Map = .TileExit.Map
                
                If EsGm(UserIndex) Then
                    Call Logs_User(UserList(UserIndex).Name, eLog.eGm, eLogDescUser.eNone, "Utilizó un teleport hacia el mapa " & DestPos.Map & " (" & DestPos.X & "," & DestPos.Y & ")")

                End If
                    
                Dim TeleportIndex As Integer

                Dim CanTelep      As Boolean
                    
                TeleportIndex = MapData(Map, X, Y).TeleportIndex

                ' @ Es un teleport de usuario se fija si puede ingresar..
                If TeleportIndex > 0 Then

                    ' ¿ Es diferente al que invoco? Comprueba si tiene valida
                    If mTeleports.Teleports(TeleportIndex).UserIndex <> UserIndex Then
                        
                        ' El Teleport permite que entren compañeros del clansuli
                        If mTeleports.Teleports(TeleportIndex).CanGuild = True Then
                            If UserList(UserIndex).GuildIndex = UserList(mTeleports.Teleports(TeleportIndex).UserIndex).GuildIndex Then
                                CanTelep = True

                            End If

                        End If
                                               
                        If mTeleports.Teleports(TeleportIndex).CanParty Then
                            If UserList(UserIndex).GroupIndex = UserList(mTeleports.Teleports(TeleportIndex).UserIndex).GroupIndex Then
                                CanTelep = True

                            End If

                        End If

                        If Not CanTelep Then
                            Call WriteConsoleMsg(UserIndex, "¡No estás habilitado para usar el Portal!", FontTypeNames.FONTTYPE_INFO)
                            Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                            
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
    
                            End If
    
                            Exit Sub
                                
                        End If
                            
                    End If

                End If
                                
                ' @  Faccion
                If MapInfo(DestPos.Map).Faction <> 0 Then

                    If UserList(UserIndex).Faction.Status = 0 And MapInfo(DestPos.Map).Faction > eFaccion.fLegion Then
                        Call WriteConsoleMsg(UserIndex, "Debes pertenecer a alguna facción para ingresar al mapa", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                        End If
                        
                        Exit Sub

                    End If
                    
                    If UserList(UserIndex).Faction.Status <> MapInfo(DestPos.Map).Faction And MapInfo(DestPos.Map).Faction <= eFaccion.fLegion Then
                        Call WriteConsoleMsg(UserIndex, "Debes pertenecer a la facción " & InfoFaction(MapInfo(DestPos.Map).Faction).Name & " para entrar al mapa.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                        End If
                        
                        Exit Sub

                    End If

                    ' Si esta muerto no puede entrar
                    If UserList(UserIndex).Faction.Status <> MapInfo(DestPos.Map).Faction Then
                        Call WriteConsoleMsg(UserIndex, "Sólo guerreros sin maná podrán ingresar a este lugar.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                        End If
                        
                        Exit Sub

                    End If

                End If
                
                ' Si es un mapa que no admite usuarios con maná
                If MapInfo(DestPos.Map).NoMana <> 0 Then

                    ' Si esta muerto no puede entrar
                    If UserList(UserIndex).Stats.MaxMan > 0 Then
                        Call WriteConsoleMsg(UserIndex, "Sólo guerreros sin maná podrán ingresar a este lugar.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                        End If
                        
                        Exit Sub

                    End If

                End If

                ' Si es un mapa que no admite muertos
                If MapInfo(DestPos.Map).OnDeathGoTo.Map <> 0 Then

                    ' Si esta muerto no puede entrar
                    If UserList(UserIndex).flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "Sólo se permite entrar al mapa a los personajes vivos.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                        End If
                        
                        Exit Sub

                    End If

                End If
                
                '¿Es mapa requeridor de clan?
                If MapInfo(DestPos.Map).Guild > 0 Then
                    If UserList(UserIndex).GuildIndex = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Sólo se permite entrar si dispones de un clan.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                        End If
                        
                        Exit Sub
                        
                    End If

                End If
                
                '¿Es mapa requeridor de nivel mínimo?
                If MapInfo(DestPos.Map).LvlMin > 0 Then
                    If MapInfo(DestPos.Map).LvlMin > UserList(UserIndex).Stats.Elv Then
                        Call WriteConsoleMsg(UserIndex, "Sólo se permite entrar al mapa siendo nivel '" & MapInfo(DestPos.Map).LvlMin & "'", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                        End If
                        
                        Exit Sub
                        
                    End If

                End If
                
                '¿Es mapa que permite un nivel máximo?
                If MapInfo(DestPos.Map).LvlMax > 0 Then
                    If MapInfo(DestPos.Map).LvlMax < UserList(UserIndex).Stats.Elv Then
                        Call WriteConsoleMsg(UserIndex, "Sólo se permite entrar al mapa siendo nivel '" & MapInfo(DestPos.Map).LvlMax & "' o inferior.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                        End If
                        
                        Exit Sub
                        
                    End If

                End If
                
                '¿Requiere un horario y día especial?
                If Not CheckMap_HourDay(UserIndex, DestPos) Then
                    Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                    End If

                    Exit Sub

                End If
                
                '¿Requiere onlines especificos?
                If Not CheckMap_Onlines(UserIndex, DestPos) Then
                    Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                    End If
                    
                    Exit Sub

                End If
                
                '¿Es mapa que permite solo PREMIUMS?
                If MapInfo(DestPos.Map).Premium > 0 Then
                    If Not UserList(UserIndex).flags.Premium = 1 Then
                        Call WriteConsoleMsg(UserIndex, "Sólo se permite entrar al mapa siendo [PREMIUM].", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                        End If
                        
                        Exit Sub
                        
                    End If

                End If
                
                '¿Es mapa que permite solo BRONCE?
                If MapInfo(DestPos.Map).Bronce > 0 Then
                    If Not UserList(UserIndex).flags.Bronce = 1 Then
                        Call WriteConsoleMsg(UserIndex, "Sólo se permite entrar al mapa siendo [AVENTURERO].", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                        End If
                        
                        Exit Sub
                        
                    End If

                End If
                
                '¿Es mapa que permite solo PLATA?
                If MapInfo(DestPos.Map).Plata > 0 Then
                    If Not UserList(UserIndex).flags.Plata = 1 Then
                        Call WriteConsoleMsg(UserIndex, "Sólo se permite entrar al mapa siendo [PLATA].", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                        End If
                        
                        Exit Sub
                        
                    End If

                End If

                '¿Es mapa de newbies?
                If MapInfo(DestPos.Map).Restringir = eRestrict.restrict_newbie Then

                    '¿El usuario es un newbie?
                    If EsNewbie(UserIndex) Or EsGm(UserIndex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)

                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                            End If

                        End If

                    Else 'No es newbie
                        Call WriteConsoleMsg(UserIndex, "Mapa exclusivo para newbies.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, False)

                        End If

                    End If

                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_armada Then '¿Es mapa de Armadas?

                    '¿El usuario es Armada?
                    If esArmada(UserIndex) Or EsGm(UserIndex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)

                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                            End If

                        End If

                    Else 'No es armada
                        Call WriteConsoleMsg(UserIndex, "Mapa exclusivo para miembros del ejército real.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                        End If

                    End If

                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_caos Then '¿Es mapa de Caos?

                    '¿El usuario es Caos?
                    If esCaos(UserIndex) Or EsGm(UserIndex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)

                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                            End If

                        End If

                    Else 'No es caos
                        Call WriteConsoleMsg(UserIndex, "Mapa exclusivo para miembros de la legión oscura.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                        End If

                    End If

                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_faccion Then '¿Es mapa de faccionarios?

                    '¿El usuario es Armada o Caos?
                    If esArmada(UserIndex) Or esCaos(UserIndex) Or EsGm(UserIndex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)

                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                            End If

                        End If

                    Else 'No es Faccionario
                        Call WriteConsoleMsg(UserIndex, "Solo se permite entrar al mapa si eres miembro de alguna facción.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                        End If

                    End If

                Else 'No es un mapa de newbies, ni Armadas, ni Caos, ni faccionario.

                    If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(UserIndex), , True) Then
                        Call WarpUserChar(UserIndex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        
                    Else
                        
                        Call ClosestLegalPos(DestPos, nPos, , , True)

                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                        End If

                    End If

                End If

                If MapInfo(DestPos.Map).Pk Then

                    Dim Temp As Long

                    Temp = MapData(Map, X, Y).ObjInfo.ObjIndex
                    
                    If Temp <> 0 Then
                        If ObjData(Temp).OBJType = otTeleport Then
                            If Not UserList(UserIndex).Counters.ShieldBlocked > 0 Then
                                UserList(UserIndex).Counters.Shield = 3
                                UserList(UserIndex).Counters.ShieldBlocked = 5
                                Call RefreshCharStatus(UserIndex)

                            End If

                        End If

                    End If

                End If
                
                'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
                Dim aN As Integer
                
                aN = UserList(UserIndex).flags.AtacadoPorNpc

                If aN > 0 Then
                    Npclist(aN).Movement = Npclist(aN).flags.OldMovement
                    Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
                    Npclist(aN).flags.AttackedBy = vbNullString
                    Npclist(aN).flags.AttackedByInteger = 0
                    Npclist(aN).Target = 0

                End If
            
                aN = UserList(UserIndex).flags.NPCAtacado

                If aN > 0 Then
                    If Npclist(aN).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                        Npclist(aN).flags.AttackedFirstBy = vbNullString

                    End If

                End If

                UserList(UserIndex).flags.AtacadoPorNpc = 0
                UserList(UserIndex).flags.NPCAtacado = 0

            End If

        End With

    End If

End Sub

Function InRangoVision(ByVal UserIndex As Integer, _
                       ByVal X As Integer, _
                       ByVal Y As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo InRangoVision_Err

    '</EhHeader>

    If X > UserList(UserIndex).Pos.X - MinXBorder And X < UserList(UserIndex).Pos.X + MinXBorder Then
        If Y > UserList(UserIndex).Pos.Y - MinYBorder And Y < UserList(UserIndex).Pos.Y + MinYBorder Then
            InRangoVision = True

            Exit Function

        End If

    End If

    InRangoVision = False

    '<EhFooter>
    Exit Function

InRangoVision_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.InRangoVision " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function InVisionRangeAndMap(ByVal UserIndex As Integer, _
                                    ByRef OtherUserPos As WorldPos) As Boolean

    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo InVisionRangeAndMap_Err

    '</EhHeader>
    
    With UserList(UserIndex)
        
        ' Same map?
        If .Pos.Map <> OtherUserPos.Map Then Exit Function
    
        ' In x range?
        If OtherUserPos.X < .Pos.X - MinXBorder Or OtherUserPos.X > .Pos.X + MinXBorder Then Exit Function
        
        ' In y range?
        If OtherUserPos.Y < .Pos.Y - MinYBorder And OtherUserPos.Y > .Pos.Y + MinYBorder Then Exit Function

    End With

    InVisionRangeAndMap = True
    
    '<EhFooter>
    Exit Function

InVisionRangeAndMap_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.InVisionRangeAndMap " & "at line " & Erl
        
    '</EhFooter>
End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, _
                          X As Integer, _
                          Y As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo InRangoVisionNPC_Err

    '</EhHeader>

    If X > Npclist(NpcIndex).Pos.X - MinXBorder And X < Npclist(NpcIndex).Pos.X + MinXBorder Then
        If Y > Npclist(NpcIndex).Pos.Y - MinYBorder And Y < Npclist(NpcIndex).Pos.Y + MinYBorder Then
            InRangoVisionNPC = True

            Exit Function

        End If

    End If

    InRangoVisionNPC = False

    '<EhFooter>
    Exit Function

InRangoVisionNPC_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.InRangoVisionNPC " & "at line " & Erl
        
    '</EhFooter>
End Function

Function InMapBounds(ByVal Map As Integer, _
                     ByVal X As Integer, _
                     ByVal Y As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo InMapBounds_Err

    '</EhHeader>

    If (Map <= 0 Or Map > NumMaps) Or X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        InMapBounds = False
    Else
        InMapBounds = True

    End If
    
    '<EhFooter>
    Exit Function

InMapBounds_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.InMapBounds " & "at line " & Erl
        
    '</EhFooter>
End Function

Private Function RhombLegalPos(ByRef Pos As WorldPos, _
                               ByRef vX As Long, _
                               ByRef vY As Long, _
                               ByVal Distance As Long, _
                               Optional PuedeAgua As Boolean = False, _
                               Optional PuedeTierra As Boolean = True, _
                               Optional ByVal CheckExitTile As Boolean = False, _
                               Optional ByVal DifPos As Boolean = False) As Boolean

    '***************************************************
    'Author: Marco Vanotti (Marco)
    'Last Modification: -
    ' walks all the perimeter of a rhomb of side  "distance + 1",
    ' which starts at Pos.x - Distance and Pos.y
    '***************************************************
    '<EhHeader>
    On Error GoTo RhombLegalPos_Err

    '</EhHeader>

    Dim i As Long
    
    vX = Pos.X - Distance
    vY = Pos.Y
    
    For i = 0 To Distance - 1

        If (LegalPos(Pos.Map, vX + i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile, DifPos)) Then
            vX = vX + i
            vY = vY - i
            RhombLegalPos = True

            Exit Function

        End If

    Next
    
    vX = Pos.X
    vY = Pos.Y - Distance
    
    For i = 0 To Distance - 1

        If (LegalPos(Pos.Map, vX + i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile, DifPos)) Then
            vX = vX + i
            vY = vY + i
            RhombLegalPos = True

            Exit Function

        End If

    Next
    
    vX = Pos.X + Distance
    vY = Pos.Y
    
    For i = 0 To Distance - 1

        If (LegalPos(Pos.Map, vX - i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile, DifPos)) Then
            vX = vX - i
            vY = vY + i
            RhombLegalPos = True

            Exit Function

        End If

    Next
    
    vX = Pos.X
    vY = Pos.Y + Distance
    
    For i = 0 To Distance - 1

        If (LegalPos(Pos.Map, vX - i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile, DifPos)) Then
            vX = vX - i
            vY = vY - i
            RhombLegalPos = True

            Exit Function

        End If

    Next
    
    RhombLegalPos = False
    
    '<EhFooter>
    Exit Function

RhombLegalPos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.RhombLegalPos " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function RhombLegalTilePos(ByRef Pos As WorldPos, _
                                  ByRef vX As Long, _
                                  ByRef vY As Long, _
                                  ByVal Distance As Long, _
                                  ByVal ObjIndex As Integer, _
                                  ByVal ObjAmount As Long, _
                                  ByVal PuedeAgua As Boolean, _
                                  ByVal PuedeTierra As Boolean) As Boolean

    '<EhHeader>
    On Error GoTo RhombLegalTilePos_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: -
    ' walks all the perimeter of a rhomb of side  "distance + 1",
    ' which starts at Pos.x - Distance and Pos.y
    ' and searchs for a valid position to drop items
    '***************************************************

    Dim i           As Long

    Dim HayObj      As Boolean
    
    Dim X           As Integer

    Dim Y           As Integer

    Dim MapObjIndex As Integer
    
    vX = Pos.X - Distance
    vY = Pos.Y
    
    For i = 0 To Distance - 1
        
        X = vX + i
        Y = vY - i
        
        If (LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
            
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex, ObjAmount) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True

                Exit Function

            End If
            
        End If

    Next
    
    vX = Pos.X
    vY = Pos.Y - Distance
    
    For i = 0 To Distance - 1
        
        X = vX + i
        Y = vY + i
        
        If (LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
            
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex, ObjAmount) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True

                Exit Function

            End If

        End If

    Next
    
    vX = Pos.X + Distance
    vY = Pos.Y
    
    For i = 0 To Distance - 1
        
        X = vX - i
        Y = vY + i
    
        If (LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
        
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex, ObjAmount) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True

                Exit Function

            End If

        End If

    Next
    
    vX = Pos.X
    vY = Pos.Y + Distance
    
    For i = 0 To Distance - 1
        
        X = vX - i
        Y = vY - i
    
        If (LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then

            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex, ObjAmount) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True

                Exit Function

            End If

        End If

    Next
    
    RhombLegalTilePos = False

    '<EhFooter>
    Exit Function

RhombLegalTilePos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.RhombLegalTilePos " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function HayObjeto(ByVal mapa As Integer, _
                          ByVal X As Long, _
                          ByVal Y As Long, _
                          ByVal ObjIndex As Integer, _
                          ByVal ObjAmount As Long) As Boolean

    '<EhHeader>
    On Error GoTo HayObjeto_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: -
    'Checks if there's space in a tile to add an itemAmount
    '***************************************************
    Dim MapObjIndex As Integer

    MapObjIndex = MapData(mapa, X, Y).ObjInfo.ObjIndex
            
    ' Hay un objeto tirado?
    If MapObjIndex <> 0 Then

        ' Es el mismo objeto?
        If MapObjIndex = ObjIndex Then
            ' La suma es menor a 10k?
            HayObjeto = (MapData(mapa, X, Y).ObjInfo.Amount + ObjAmount > MAX_INVENTORY_OBJS)
        Else
            HayObjeto = True

        End If

    Else
        HayObjeto = False

    End If

    '<EhFooter>
    Exit Function

HayObjeto_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.HayObjeto " & "at line " & Erl
        
    '</EhFooter>
End Function

Sub ClosestLegalPos(Pos As WorldPos, _
                    ByRef nPos As WorldPos, _
                    Optional PuedeAgua As Boolean = False, _
                    Optional PuedeTierra As Boolean = True, _
                    Optional ByVal CheckExitTile As Boolean = False, _
                    Optional ByVal DifPos As Boolean = False)

    '*****************************************************************
    'Author: Unknown (original version)
    'Last Modification: 09/14/2010 (Marco)
    'History:
    ' - 01/24/2007 (ToxicWaste)
    'Encuentra la posicion legal mas cercana y la guarda en nPos
    '*****************************************************************
    '<EhHeader>
    On Error GoTo ClosestLegalPos_Err

    '</EhHeader>

    Dim Found As Boolean

    Dim LoopC As Integer

    Dim tX    As Long

    Dim tY    As Long
    
    nPos = Pos
    tX = Pos.X
    tY = Pos.Y
    
    LoopC = 1
    
    ' La primera posicion es valida?
    If LegalPos(Pos.Map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra, CheckExitTile) Then
        Found = True
    
        ' Busca en las demas posiciones, en forma de "rombo"
    Else

        While (Not Found) And LoopC <= 12

            If RhombLegalPos(Pos, tX, tY, LoopC, PuedeAgua, PuedeTierra, CheckExitTile, DifPos) Then
                nPos.X = tX
                nPos.Y = tY
                Found = True

            End If
        
            LoopC = LoopC + 1

        Wend
        
    End If
    
    If Not Found Then
        nPos.X = 0
        nPos.Y = 0

    End If

    '<EhFooter>
    Exit Sub

ClosestLegalPos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.ClosestLegalPos " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)

    '***************************************************
    'Author: Unknown
    'Last Modification: 09/14/2010
    'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
    '*****************************************************************
    '<EhHeader>
    On Error GoTo ClosestStablePos_Err

    '</EhHeader>
    Call ClosestLegalPos(Pos, nPos, , , True)
          
    '<EhFooter>
    Exit Sub

ClosestStablePos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.ClosestStablePos " & "at line " & Erl
        
    '</EhFooter>
End Sub

Function NameIndex(ByVal Name As String) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo NameIndex_Err

    '</EhHeader>

    Dim UserIndex As Long
    
    '¿Nombre valido?
    If LenB(Name) = 0 Then
        NameIndex = 0

        Exit Function

    End If
    
    If InStrB(Name, "+") <> 0 Then
        Name = UCase$(Replace(Name, "+", " "))

    End If
    
    UserIndex = 1

    Do Until UCase$(UserList(UserIndex).Name) = UCase$(Name)
        
        UserIndex = UserIndex + 1
        
        If UserIndex > MaxUsers Then
            NameIndex = 0

            Exit Function

        End If

    Loop
     
    NameIndex = UserIndex
    '<EhFooter>
    Exit Function

NameIndex_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.NameIndex " & "at line " & Erl
        
    '</EhFooter>
End Function

Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo CheckForSameIP_Err

    '</EhHeader>

    Dim LoopC  As Long

    Dim Amount As Integer
    
    For LoopC = 1 To LastUser

        If UserList(LoopC).flags.UserLogged = True Then
            If UserList(LoopC).IpAddress = UserIP And UserIndex <> LoopC Then
                Amount = Amount + 1

            End If

        End If

    Next LoopC
    
    CheckForSameIP = Amount

    '<EhFooter>
    Exit Function

CheckForSameIP_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.CheckForSameIP " & "at line " & Erl
        
    '</EhFooter>
End Function

Function CheckForSameName(ByVal Name As String) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo CheckForSameName_Err

    '</EhHeader>

    'Controlo que no existan usuarios con el mismo nombre
    Dim LoopC As Long
    
    For LoopC = 1 To LastUser

        If UserList(LoopC).flags.UserLogged Then
            If UCase$(UserList(LoopC).Name) = UCase$(Name) Then
                CheckForSameName = True

                Exit Function

            End If

        End If

    Next LoopC
    
    CheckForSameName = False
    '<EhFooter>
    Exit Function

CheckForSameName_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.CheckForSameName " & "at line " & Erl
        
    '</EhFooter>
End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'Toma una posicion y se mueve hacia donde esta perfilado
    '*****************************************************************
    '<EhHeader>
    On Error GoTo HeadtoPos_Err

    '</EhHeader>

    Select Case Head

        Case eHeading.NORTH
            Pos.Y = Pos.Y - 1
        
        Case eHeading.SOUTH
            Pos.Y = Pos.Y + 1
        
        Case eHeading.EAST
            Pos.X = Pos.X + 1
        
        Case eHeading.WEST
            Pos.X = Pos.X - 1

    End Select

    '<EhFooter>
    Exit Sub

HeadtoPos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.HeadtoPos " & "at line " & Erl
        
    '</EhFooter>
End Sub

Function LegalPos(ByVal Map As Integer, _
                  ByVal X As Integer, _
                  ByVal Y As Integer, _
                  Optional ByVal PuedeAgua As Boolean = False, _
                  Optional ByVal PuedeTierra As Boolean = True, _
                  Optional ByVal CheckExitTile As Boolean = False, _
                  Optional ByVal DifPos As Boolean = False) As Boolean

    '***************************************************
    'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 23/01/2007
    'Checks if the position is Legal.
    '***************************************************
    '<EhHeader>
    On Error GoTo LegalPos_Err

    '</EhHeader>

    '¿Es un mapa valido?
    If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPos = False
    Else

        With MapData(Map, X, Y)

            If PuedeAgua And PuedeTierra Then
                LegalPos = (.Blocked <> 1) And (.UserIndex = 0) And (.NpcIndex = 0)
            ElseIf PuedeTierra And Not PuedeAgua Then
                LegalPos = (.Blocked <> 1) And (.UserIndex = 0) And (.NpcIndex = 0) And (Not HayAgua(Map, X, Y))
            ElseIf PuedeAgua And Not PuedeTierra Then
                LegalPos = (.Blocked <> 1) And (.UserIndex = 0) And (.NpcIndex = 0) And (HayAgua(Map, X, Y))
            Else
                LegalPos = False

            End If

        End With
        
        If CheckExitTile Then
            LegalPos = LegalPos And (MapData(Map, X, Y).TileExit.Map = 0)

        End If
        
    End If

    '<EhFooter>
    Exit Function

LegalPos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.LegalPos " & "at line " & Erl
        
    '</EhFooter>
End Function

Function MoveToLegalPos(ByVal Map As Integer, _
                        ByVal X As Integer, _
                        ByVal Y As Integer, _
                        Optional ByVal PuedeAgua As Boolean = False, _
                        Optional ByVal PuedeTierra As Boolean = True) As Boolean

    '***************************************************
    'Autor: ZaMa
    'Last Modification: 13/07/2009
    'Checks if the position is Legal, but considers that if there's a casper, it's a legal movement.
    '13/07/2009: ZaMa - Now it's also legal move where an invisible admin is.
    '***************************************************
    '<EhHeader>
    On Error GoTo MoveToLegalPos_Err

    '</EhHeader>

    Dim UserIndex        As Integer

    Dim IsDeadChar       As Boolean

    Dim IsAdminInvisible As Boolean

    '¿Es un mapa valido?
    If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        MoveToLegalPos = False
    Else

        With MapData(Map, X, Y)
            UserIndex = .UserIndex
        
            If UserIndex > 0 Then
                #If Classic = 0 Then
                    IsDeadChar = True
                #Else
                    IsDeadChar = (UserList(UserIndex).flags.Muerto = 1)
                #End If

                IsAdminInvisible = (UserList(UserIndex).flags.AdminInvisible = 1)
            Else
                IsDeadChar = False
                IsAdminInvisible = False

            End If
        
            If PuedeAgua And PuedeTierra Then
                MoveToLegalPos = (.Blocked <> 1) And (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0)
            ElseIf PuedeTierra And Not PuedeAgua Then
                MoveToLegalPos = (.Blocked <> 1) And (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (Not HayAgua(Map, X, Y))
            ElseIf PuedeAgua And Not PuedeTierra Then
                MoveToLegalPos = (.Blocked <> 1) And (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (HayAgua(Map, X, Y))
            Else
                MoveToLegalPos = False

            End If

        End With

    End If

    '<EhFooter>
    Exit Function

MoveToLegalPos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.MoveToLegalPos " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub FindLegalPos(ByVal UserIndex As Integer, _
                        ByVal Map As Integer, _
                        ByRef X As Integer, _
                        ByRef Y As Integer)

    '<EhHeader>
    On Error GoTo FindLegalPos_Err

    '</EhHeader>

    '***************************************************
    'Autor: ZaMa
    'Last Modification: 26/03/2009
    'Search for a Legal pos for the user who is being teleported.
    '***************************************************

    If MapData(Map, X, Y).UserIndex <> 0 Or MapData(Map, X, Y).NpcIndex <> 0 Then
                    
        ' Se teletransporta a la misma pos a la que estaba
        If MapData(Map, X, Y).UserIndex = UserIndex Then Exit Sub
                            
        Dim FoundPlace     As Boolean

        Dim tX             As Long

        Dim tY             As Long

        Dim Rango          As Long

        Dim OtherUserIndex As Integer
    
        For Rango = 1 To 5
            For tY = Y - Rango To Y + Rango
                For tX = X - Rango To X + Rango

                    'Reviso que no haya User ni NPC
                    If MapData(Map, tX, tY).UserIndex = 0 And MapData(Map, tX, tY).NpcIndex = 0 Then
                        
                        If InMapBounds(Map, tX, tY) Then FoundPlace = True
                        
                        Exit For

                    End If

                Next tX
        
                If FoundPlace Then Exit For
            Next tY
            
            If FoundPlace Then Exit For
        Next Rango
    
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            X = tX
            Y = tY
        Else
            'Muy poco probable, pero..
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
            OtherUserIndex = MapData(Map, X, Y).UserIndex

            If OtherUserIndex <> 0 Then

                'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If UserList(OtherUserIndex).ComUsu.DestUsu > 0 Then

                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(UserList(OtherUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                        Call FinComerciarUsu(UserList(OtherUserIndex).ComUsu.DestUsu)
                        Call WriteConsoleMsg(UserList(OtherUserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                        Call FlushBuffer(UserList(OtherUserIndex).ComUsu.DestUsu)

                    End If

                    'Lo sacamos.
                    If UserList(OtherUserIndex).flags.UserLogged Then
                        Call FinComerciarUsu(OtherUserIndex)
                        Call WriteErrorMsg(OtherUserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                        Call FlushBuffer(OtherUserIndex)

                    End If

                End If
            
                'Call CloseSocket(OtherUserIndex)
                Call WriteDisconnect(OtherUserIndex)
                Call FlushBuffer(OtherUserIndex)
                                                            
                Call CloseSocket(OtherUserIndex)

            End If

        End If

    End If

    '<EhFooter>
    Exit Sub

FindLegalPos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.FindLegalPos " & "at line " & Erl
        
    '</EhFooter>
End Sub

Function LegalPosNPC(ByVal Map As Integer, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal AguaValida As Byte, _
                     Optional ByVal IsPet As Boolean = False, _
                     Optional ByVal TierraInvalida As Boolean = False) As Boolean

    '***************************************************
    'Autor: Unkwnown
    'Last Modification: 09/23/2009
    'Checks if it's a Legal pos for the npc to move to.
    '09/23/2009: Pato - If UserIndex is a AdminInvisible, then is a legal pos.
    '***************************************************
    '<EhHeader>
    On Error GoTo LegalPosNPC_Err

    '</EhHeader>

    Dim IsDeadChar       As Boolean

    Dim UserIndex        As Integer

    Dim IsAdminInvisible As Boolean
 
    If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPosNPC = False
        Exit Function

    End If
 
    With MapData(Map, X, Y)
        UserIndex = .UserIndex

        If UserIndex > 0 Then
            IsDeadChar = UserList(UserIndex).flags.Muerto = 1
            IsAdminInvisible = (UserList(UserIndex).flags.AdminInvisible = 1)
        Else
            IsDeadChar = False
            IsAdminInvisible = False

        End If
 
        ' if it's a pet, check if is going to walk on a tp
        If IsPet And .ObjInfo.ObjIndex <> 0 Then
            If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport Then
                LegalPosNPC = False
                Exit Function

            End If

        End If
 
        If AguaValida = 0 Then
            LegalPosNPC = (.Blocked <> 1) And (.UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (.trigger <> eTrigger.POSINVALIDA Or IsPet) And Not HayAgua(Map, X, Y)
        ElseIf TierraInvalida = False Then
            LegalPosNPC = (.Blocked <> 1) And (.UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (.trigger <> eTrigger.POSINVALIDA Or IsPet)
        Else
            LegalPosNPC = (.Blocked <> 1) And (.UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (.trigger <> eTrigger.POSINVALIDA And HayAgua(Map, X, Y) Or IsPet)

        End If

    End With

    '<EhFooter>
    Exit Function

LegalPosNPC_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.LegalPosNPC " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub LoadHelp()

    '<EhHeader>
    On Error GoTo LoadHelp_Err

    '</EhHeader>
    Dim Manager As clsIniManager
    
    Set Manager = New clsIniManager
    
    Manager.Initialize DatPath & "Help.dat"

    Dim A As Long
    
    HelpLast = val(Manager.GetValue("INIT", "NUMLINES"))
    
    ReDim HelpLines(1 To HelpLast) As String
    
    For A = 1 To HelpLast
        HelpLines(A) = Manager.GetValue("HELP", "LINE" & A)
    Next A
    
    Set Manager = Nothing
    '<EhFooter>
    Exit Sub

LoadHelp_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.LoadHelp " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub SendHelp(ByVal Index As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo SendHelp_Err

    '</EhHeader>
        
    Dim LoopC As Long
        
    For LoopC = 1 To HelpLast
        Call WriteConsoleMsg(Index, HelpLines(LoopC), FontTypeNames.FONTTYPE_INFO)
    Next LoopC

    '<EhFooter>
    Exit Sub

SendHelp_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.SendHelp " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo Expresar_Err

    '</EhHeader>

    If Npclist(NpcIndex).NroExpresiones > 0 Then

        Dim randomi

        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Npclist(NpcIndex).Expresiones(randomi), Npclist(NpcIndex).Char.charindex, vbWhite))

    End If

    '<EhFooter>
    Exit Sub

Expresar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.Expresar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub LookatTile(ByVal UserIndex As Integer, _
               ByVal Map As Integer, _
               ByVal X As Integer, _
               ByVal Y As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 26/03/2009
    '13/02/2009: ZaMa - El nombre del gm que aparece por consola al clickearlo, tiene el color correspondiente a su rango
    '***************************************************
    '<EhHeader>
    On Error GoTo LookatTile_Err

    '</EhHeader>

    'Responde al click del usuario sobre el mapa
    Dim FoundChar      As Byte

    Dim FoundSomething As Byte

    Dim TempCharIndex  As Integer

    Dim Stat           As String

    Dim Ft             As FontTypeNames

    With UserList(UserIndex)
        'If .flags.GmSeguidor > 0 Then
        'Call WriteUpdateListSecurity(.flags.GmSeguidor, .Name, "Click a  Pos X: " & X & ", Y:" & Y & ". UserIndex: " & MapData(Map, X, Y).UserIndex, 1)
        'Call LogError("GM SEGUIDOR: " & UserList(.flags.GmSeguidor).Name)
        'End If
        
        '¿Rango Visión? (ToxicWaste)
        If (Abs(.Pos.Y - Y) > RANGO_VISION_y) Or (Abs(.Pos.X - X) > RANGO_VISION_x) Then

            Exit Sub

        End If
    
        '¿Posicion valida?
        If InMapBounds(Map, X, Y) Then

            With .flags
                .TargetMap = Map
                .TargetX = X
                .TargetY = Y

                '¿Es un obj?
                If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
                    'Informa el nombre
                    .TargetObjMap = Map
                    .TargetObjX = X
                    .TargetObjY = Y
                    FoundSomething = 1
                ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then

                    'Informa el nombre
                    If ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                        .TargetObjMap = Map
                        .TargetObjX = X + 1
                        .TargetObjY = Y
                        FoundSomething = 1

                    End If

                ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then

                    If ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                        'Informa el nombre
                        .TargetObjMap = Map
                        .TargetObjX = X + 1
                        .TargetObjY = Y + 1
                        FoundSomething = 1

                    End If

                ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then

                    If ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                        'Informa el nombre
                        .TargetObjMap = Map
                        .TargetObjX = X
                        .TargetObjY = Y + 1
                        FoundSomething = 1

                    End If

                End If
            
                If FoundSomething = 1 Then
                    .TargetObj = MapData(Map, .TargetObjX, .TargetObjY).ObjInfo.ObjIndex
                    
                    If ObjData(.TargetObj).OBJType = otTeleport Then
                        If MapData(Map, X, Y).TileExit.Map <> 0 Then
                            If EsGm(UserIndex) Then
                                Call WriteConsoleMsg(UserIndex, "Teleport a '" & MapInfo(MapData(Map, X, Y).TileExit.Map).Name & "' (" & MapData(Map, X, Y).TileExit.Map & ", " & MapData(Map, X, Y).TileExit.X & ", " & MapData(Map, X, Y).TileExit.Y & ")", FontTypeNames.FONTTYPE_INFO)
                            Else
                            
                                Call WriteConsoleMsg(UserIndex, "Teleport a '" & MapInfo(MapData(Map, X, Y).TileExit.Map).Name & "'", FontTypeNames.FONTTYPE_INFO)

                            End If

                        End If

                    Else

                        If MostrarCantidad(.TargetObj) Then
                            Call WriteConsoleMsg(UserIndex, ObjData(.TargetObj).Name & " - " & MapData(.TargetObjMap, .TargetObjX, .TargetObjY).ObjInfo.Amount & "", FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(UserIndex, ObjData(.TargetObj).Name, FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If

                End If
            
                '¿Es un personaje?
                If Y + 1 <= YMaxMapSize Then
                    If MapData(Map, X, Y + 1).UserIndex > 0 Then
                        TempCharIndex = MapData(Map, X, Y + 1).UserIndex
                        FoundChar = 1

                    End If

                    If MapData(Map, X, Y + 1).NpcIndex > 0 Then
                        TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
                        FoundChar = 2

                    End If

                End If

                '¿Es un personaje?
                If FoundChar = 0 Then
                    If MapData(Map, X, Y).UserIndex > 0 Then
                        TempCharIndex = MapData(Map, X, Y).UserIndex
                        FoundChar = 1

                    End If

                    If MapData(Map, X, Y).NpcIndex > 0 Then
                        TempCharIndex = MapData(Map, X, Y).NpcIndex
                        FoundChar = 2

                    End If

                End If

            End With
    
            'Reaccion al personaje
            If FoundChar = 1 Then '  ¿Encontro un Usuario?
                If UserList(TempCharIndex).flags.AdminInvisible = 0 Then

                    With UserList(TempCharIndex)

                        If LenB(.DescRM) = 0 And .ShowName Then 'No tiene descRM y quiere que se vea su nombre.
                    
                            Dim Name         As String

                            Dim Desc         As String
                        
                            Dim FactionRange As String

                            Dim GuildName    As String

                            Dim RangeGm      As String

                            Dim Concilio     As Byte

                            Dim Consejo      As Byte

                            Dim sPlayerType  As PlayerType
                        
                            If .GuildIndex > 0 Then
                                GuildName = GuildsInfo(.GuildIndex).Name

                            End If
                        
                            If .Faction.Status <> r_None Then
                                FactionRange = InfoFaction(.Faction.Status).Range(.Faction.Range).Text

                            End If
                        
                            If EsGm(TempCharIndex) Then
                                RangeGm = GetCharRange(UCase$(.Name))

                            End If
                        
                            If Escriminal(TempCharIndex) Then
                                Ft = FontTypeNames.FONTTYPE_FIGHT
                            Else
                                Ft = FontTypeNames.FONTTYPE_CITIZEN

                            End If
                        
                            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                                sPlayerType = PlayerType.RoyalCouncil
                                Ft = FontTypeNames.FONTTYPE_CONSEJOVesA
                            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                                sPlayerType = PlayerType.ChaosCouncil
                                Ft = FontTypeNames.FONTTYPE_EJECUCION
                            Else

                                If Not .flags.Privilegios And PlayerType.User Then
                                    If .flags.Privilegios = PlayerType.Admin Then
                                        sPlayerType = PlayerType.Admin
                                        Ft = FontTypeNames.FONTTYPE_ADMIN
                                    ElseIf .flags.Privilegios = PlayerType.Dios Then
                                        sPlayerType = PlayerType.Dios
                                        Ft = FontTypeNames.FONTTYPE_DIOS
                                        ' Gm
                                    ElseIf .flags.Privilegios = PlayerType.SemiDios Then
                                        sPlayerType = PlayerType.SemiDios
                                        Ft = FontTypeNames.FONTTYPE_GM

                                    End If

                                End If

                            End If

                        Else  'Si tiene descRM la muestro siempre.
                            Stat = .DescRM

                        End If
                    
                        If .ShowName = True Then
                            If UserList(TempCharIndex).flags.SlotEvent = 0 Then
                                Call WriteVesA(UserIndex, .secName, .Desc, .Clase, .Raza, .Faction.Status, FactionRange, GuildName, .GuildRange, RangeGm, .flags.Privilegios, .flags.Oro, .flags.Bronce, .flags.Plata, .flags.Premium, .flags.ModoStream, .flags.Transform, .flags.Muerto, Ft, .flags.StreamUrl, .flags.RachasTemp, .flags.Rachas)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Participante", FontTypeNames.FONTTYPE_INFOBOLD)

                            End If

                        End If

                    End With
                
                    FoundSomething = 1
                    .flags.TargetUser = TempCharIndex
                    .flags.TargetNPC = 0
                    .flags.TargetNpcTipo = eNPCType.Comun

                End If

            End If
    
            With .flags

                If FoundChar = 2 Then '¿Encontro un NPC?

                    Dim estatus As String

                    Dim MinHp   As Long

                    Dim MaxHp   As Long

                    Dim Elv     As Byte

                    Dim sDesc   As String
                        
                    Elv = UserList(UserIndex).Stats.Elv
                    MinHp = Npclist(TempCharIndex).Stats.MinHp
                    MaxHp = Npclist(TempCharIndex).Stats.MaxHp
                
                    If .Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then
                        estatus = "(" & MinHp & "/" & MaxHp & ") "
                    Else

                        If .Muerto = 0 And MinHp > 0 Then
                    
                            If Elv <= 5 Then
                                estatus = "(Dudoso) "
                            
                            ElseIf Elv <= 13 Then

                                If MinHp < (MaxHp / 2) Then
                                    estatus = "(Herido) "
                                Else
                                    estatus = "(Sano) "

                                End If
                            
                            ElseIf Elv <= 21 Then

                                If MinHp < (MaxHp * 0.5) Then
                                    estatus = "(Malherido) "
                                ElseIf MinHp < (MaxHp * 0.75) Then
                                    estatus = "(Herido) "
                                Else
                                    estatus = "(Sano) "

                                End If
                            
                            ElseIf Elv <= 25 Then

                                If MinHp < (MaxHp * 0.25) Then
                                    estatus = "(Muy malherido) "
                                ElseIf MinHp < (MaxHp * 0.5) Then
                                    estatus = "(Herido) "
                                ElseIf MinHp < (MaxHp * 0.75) Then
                                    estatus = "(Levemente herido) "
                                Else
                                    estatus = "(Sano) "

                                End If
                            
                            ElseIf Elv < 36 Then

                                If MinHp < (MaxHp * 0.05) Then
                                    estatus = "(Agonizando) "
                                ElseIf MinHp < (MaxHp * 0.1) Then
                                    estatus = "(Casi muerto) "
                                ElseIf MinHp < (MaxHp * 0.25) Then
                                    estatus = "(Muy Malherido) "
                                ElseIf MinHp < (MaxHp * 0.5) Then
                                    estatus = "(Herido) "
                                ElseIf MinHp < (MaxHp * 0.75) Then
                                    estatus = "(Levemente herido) "
                                ElseIf MinHp < (MaxHp) Then
                                    estatus = "(Sano) "
                                Else
                                    estatus = "(Intacto) "

                                End If
                                
                            ElseIf Elv < 40 Then
                                estatus = "(" & Round(CDbl(MinHp) * CDbl(100) / CDbl(MaxHp), 2) & "%)"

                            Else

                                estatus = "(" & MinHp & "/" & MaxHp & ") "

                            End If
                            
                            If Npclist(TempCharIndex).flags.Invasion = 1 Or Npclist(TempCharIndex).Stats.Elv > UserList(UserIndex).Stats.Elv Then
                                estatus = "(Dudoso) "

                            End If
                            
                            If Npclist(TempCharIndex).Stats.Elv > 0 Then
                                estatus = estatus & "(Lvl " & Npclist(TempCharIndex).Stats.Elv & ")"

                            End If

                        End If

                    End If
                
                    If Len(Npclist(TempCharIndex).Desc) > 1 Then
                        Stat = Npclist(TempCharIndex).Desc
                    
                        ' Informacion de ciertos NPCS:: APARTAR EN ALGUN MODULO :: LAUTARO
                        If Npclist(TempCharIndex).numero = 742 Then
                            If ConfigServer.ModoRetosFast = 0 Then
                                Stat = Stat & " (DESACTIVADO)"
                            Else
                                Stat = Stat & " (ACTIVADO)"

                            End If

                        End If
                        
                        If EsGm(UserIndex) Then
                            Call WriteConsoleMsg(UserIndex, "Npc n°" & Npclist(TempCharIndex).numero & " CharIndex: " & Npclist(TempCharIndex).Char.charindex, FontTypeNames.FONTTYPE_INFO)

                        End If
                    
                        'Enviamos el mensaje propiamente dicho:
                        Call WriteChatOverHead(UserIndex, Stat, Npclist(TempCharIndex).Char.charindex, vbWhite)
                        'Call WriteConsoleMsg(UserIndex, Npclist(TempCharIndex).Name & IIf(Npclist(TempCharIndex).Level > 0, " <Nivel: " & Npclist(TempCharIndex).Level & ">", vbNullString), FontTypeNames.FONTTYPE_INFO)
                        Call WriteConsoleMsg(UserIndex, Npclist(TempCharIndex).Name & " " & estatus, FontTypeNames.FONTTYPE_INFO)
                    Else

                        If Npclist(TempCharIndex).MaestroUser > 0 Then
                            Call WriteConsoleMsg(UserIndex, estatus & Npclist(TempCharIndex).Name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).Name & ".", FontTypeNames.FONTTYPE_INFO)
                        Else
                            sDesc = Npclist(TempCharIndex).Name & " " & estatus

                            If Npclist(TempCharIndex).Owner > 0 Then
                                sDesc = sDesc & " le pertenece a " & UserList(Npclist(TempCharIndex).Owner).Name & "."

                            End If
                                
                            If EsGm(UserIndex) Then
                                Call WriteConsoleMsg(UserIndex, sDesc & "Npc n°" & Npclist(TempCharIndex).numero, FontTypeNames.FONTTYPE_INFO)
                            Else
                                Call WriteConsoleMsg(UserIndex, sDesc, FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                            If .Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                                Call WriteConsoleMsg(UserIndex, "Le pegó primero: " & Npclist(TempCharIndex).flags.AttackedFirstBy & ".", FontTypeNames.FONTTYPE_INFO)

                            End If
                          
                        End If

                    End If
                
                    FoundSomething = 1
                    .TargetNpcTipo = Npclist(TempCharIndex).NPCtype
                    .TargetNPC = TempCharIndex
                    .TargetUser = 0
                    .TargetObj = 0

                End If
            
                If FoundChar = 0 Then
                    .TargetNPC = 0
                    .TargetNpcTipo = eNPCType.Comun
                    .TargetUser = 0

                End If
            
                '*** NO ENCOTRO NADA ***
                If FoundSomething = 0 Then
                    .TargetNPC = 0
                    .TargetNpcTipo = eNPCType.Comun
                    .TargetUser = 0
                    .TargetObj = 0
                    .TargetObjMap = 0
                    .TargetObjX = 0
                    .TargetObjY = 0

                    'Call WriteMultiMessage(UserIndex, eMessages.DontSeeAnything)
                End If

            End With

        Else

            If FoundSomething = 0 Then

                With .flags
                    .TargetNPC = 0
                    .TargetNpcTipo = eNPCType.Comun
                    .TargetUser = 0
                    .TargetObj = 0
                    .TargetObjMap = 0
                    .TargetObjX = 0
                    .TargetObjY = 0

                End With
            
                'Call WriteMultiMessage(UserIndex, eMessages.DontSeeAnything)
            End If

        End If

    End With

    '<EhFooter>
    Exit Sub

LookatTile_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.LookatTile " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub ShowMenu(ByVal UserIndex As Integer, _
                    ByVal Map As Integer, _
                    ByVal X As Integer, _
                    ByVal Y As Integer)
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 10/05/2010
    'Shows menu according to user, npc or object right clicked.
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        ' In Vision Range
        If (Abs(.Pos.Y - Y) > RANGO_VISION_y) Or (Abs(.Pos.X - X) > RANGO_VISION_x) Then Exit Sub
        
        ' Valid position?
        If Not InMapBounds(Map, X, Y) Then Exit Sub
        
        With .flags

            ' Alive?
            If .Muerto = 1 Then Exit Sub
            
            ' Trading?
            If .Comerciando Then Exit Sub
            
            ' Reset flags
            .TargetNPC = 0
            .TargetNpcTipo = eNPCType.Comun
            .TargetUser = 0
            .TargetObj = 0
            .TargetObjMap = 0
            .TargetObjX = 0
            .TargetObjY = 0
            
            .TargetMap = Map
            .TargetX = X
            .TargetY = Y
            
            Dim tmpIndex  As Integer

            Dim FoundChar As Byte

            Dim MenuIndex As Integer
            
            ' Npc or user? (lower position)
            If Y + 1 <= YMaxMapSize Then
                
                ' User?
                tmpIndex = MapData(Map, X, Y + 1).UserIndex

                If tmpIndex > 0 Then

                    ' Invalid: Admin invisible, user invisible or hidden (if not is userindex)
                    If (UserList(tmpIndex).flags.AdminInvisible Or UserList(tmpIndex).flags.Invisible Or UserList(tmpIndex).flags.Oculto) = 0 Or tmpIndex = UserIndex Then
                        
                        FoundChar = 1

                    End If

                End If
                
                ' Npc?
                If MapData(Map, X, Y + 1).NpcIndex > 0 Then
                    tmpIndex = MapData(Map, X, Y + 1).NpcIndex
                    FoundChar = 2

                End If

            End If
             
            ' Npc or user? (upper position)
            If FoundChar = 0 Then
                
                ' User?
                tmpIndex = MapData(Map, X, Y).UserIndex

                If tmpIndex > 0 Then

                    ' Invalid: Admin invisible, user invisible or hidden (if not is userindex)
                    If (UserList(tmpIndex).flags.AdminInvisible Or UserList(tmpIndex).flags.Invisible Or UserList(tmpIndex).flags.Oculto) = 0 Or tmpIndex = UserIndex Then
                        
                        FoundChar = 1

                    End If

                End If
                
                ' Npc?
                If MapData(Map, X, Y).NpcIndex > 0 Then
                    tmpIndex = MapData(Map, X, Y).NpcIndex
                    FoundChar = 2

                End If

            End If
            
            ' User
            If FoundChar = 1 Then
                ' If Interval_Packet250(UserIndex) Then
                '  Call WriteStatsUser(UserIndex, UserList(UserIndex))
                ' End If
                    
                ' Self clicked => pick item
                If tmpIndex = UserIndex Then
                
                    If EsGm(UserIndex) Then Exit Sub
                    If EsNewbie(UserIndex) Then Exit Sub
                    
                    ' Pick item
                    Call GetObj(UserIndex)
                Else

                    ' Sharing npc?
                    If .ShareNpcWith = tmpIndex Then
                        MenuIndex = eMenues.ieOtroUserCompartiendoNpc
                    Else
                        MenuIndex = eMenues.ieOtroUser

                    End If
                    
                    .TargetUser = tmpIndex
                   
                End If
                
                ' Npc
            ElseIf FoundChar = 2 Then

                ' Has menu attached?
                If Npclist(tmpIndex).MenuIndex <> 0 Then
                    MenuIndex = Npclist(tmpIndex).MenuIndex

                End If
                
                .TargetNpcTipo = Npclist(tmpIndex).NPCtype
                .TargetNPC = tmpIndex
                
                'If Npclist(tmpIndex).Stats.MinHp > 0 Then
                'Call WriteSendInfoNpc(UserIndex, Npclist(tmpIndex).Numero)
                '  Else
                Call Accion(UserIndex, Map, X, Y, 0)
                ' End If
                
            End If
            
            ' No user or npc found
            If FoundChar = 0 Then
                
                ' Is there any object?
                tmpIndex = MapData(Map, X, Y).ObjInfo.ObjIndex

                If tmpIndex > 0 Then
                    ' Has menu attached?
                    MenuIndex = ObjData(tmpIndex).MenuIndex
                    
                    If MenuIndex = eMenues.ieFogata Then
                        If .Descansar = 1 Then MenuIndex = eMenues.ieFogataDescansando

                    End If
                    
                    .TargetObjMap = Map
                    .TargetObjX = X
                    .TargetObjY = Y

                End If

            End If

        End With

    End With
    
    ' Show it
    If MenuIndex <> 0 Then Call WriteShowMenu(UserIndex, MenuIndex)
    
    Exit Sub

ErrHandler:
    Call LogError("Error en ShowMenu. Error " & Err.number & " : " & Err.description)

End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As eHeading

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'Devuelve la direccion en la cual el target se encuentra
    'desde pos, 0 si la direc es igual
    '*****************************************************************
    '<EhHeader>
    On Error GoTo FindDirection_Err

    '</EhHeader>

    Dim X As Integer

    Dim Y As Integer
    
    X = Pos.X - Target.X
    Y = Pos.Y - Target.Y
    
    'NE
    If Sgn(X) = -1 And Sgn(Y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.NORTH, eHeading.EAST)

        Exit Function

    End If
    
    'NW
    If Sgn(X) = 1 And Sgn(Y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.NORTH)

        Exit Function

    End If
    
    'SW
    If Sgn(X) = 1 And Sgn(Y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.SOUTH)

        Exit Function

    End If
    
    'SE
    If Sgn(X) = -1 And Sgn(Y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.SOUTH, eHeading.EAST)

        Exit Function

    End If
    
    'Sur
    If Sgn(X) = 0 And Sgn(Y) = -1 Then
        FindDirection = eHeading.SOUTH

        Exit Function

    End If
    
    'norte
    If Sgn(X) = 0 And Sgn(Y) = 1 Then
        FindDirection = eHeading.NORTH

        Exit Function

    End If
    
    'oeste
    If Sgn(X) = 1 And Sgn(Y) = 0 Then
        FindDirection = eHeading.WEST

        Exit Function

    End If
    
    'este
    If Sgn(X) = -1 And Sgn(Y) = 0 Then
        FindDirection = eHeading.EAST

        Exit Function

    End If
    
    'misma
    If Sgn(X) = 0 And Sgn(Y) = 0 Then
        FindDirection = 0

        Exit Function

    End If

    '<EhFooter>
    Exit Function

FindDirection_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.FindDirection " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function ItemNoEsDeMapa(ByVal Map As Integer, _
                               ByVal X As Byte, _
                               ByVal Y As Byte, _
                               ByVal bIsExit As Boolean) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ItemNoEsDeMapa_Err

    '</EhHeader>

    Dim ObjIndex As Integer
    
    ObjIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
    
    With ObjData(ObjIndex)
        ItemNoEsDeMapa = .OBJType <> eOBJType.otPuertas And .OBJType <> eOBJType.otForos And .OBJType <> eOBJType.otCarteles And .OBJType <> eOBJType.otArboles And .OBJType <> eOBJType.otYacimiento And Not (.OBJType = eOBJType.otTeleport And bIsExit) And (MapData(Map, X, Y).Blocked = 0) 'And (MapData(Map, X, Y).trigger = 0)
    
    End With

    '<EhFooter>
    Exit Function

ItemNoEsDeMapa_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.ItemNoEsDeMapa " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function MostrarCantidad(ByVal Index As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo MostrarCantidad_Err

    '</EhHeader>

    With ObjData(Index)
        MostrarCantidad = .OBJType <> eOBJType.otPuertas And .OBJType <> eOBJType.otForos And .OBJType <> eOBJType.otCarteles And .OBJType <> eOBJType.otArboles And .OBJType <> eOBJType.otYacimiento And .OBJType <> eOBJType.otTeleport

    End With

    '<EhFooter>
    Exit Function

MostrarCantidad_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.MostrarCantidad " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo EsObjetoFijo_Err

    '</EhHeader>

    EsObjetoFijo = OBJType = eOBJType.otForos Or OBJType = eOBJType.otCarteles Or OBJType = eOBJType.otArboles Or OBJType = eOBJType.otYacimiento
                   
    '<EhFooter>
    Exit Function

EsObjetoFijo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.EsObjetoFijo " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function RestrictStringToByte(ByRef restrict As String) As Byte

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 04/18/2011
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo RestrictStringToByte_Err

    '</EhHeader>
    restrict = UCase$(restrict)

    Select Case restrict

        Case "NEWBIE"
            RestrictStringToByte = 1
        
        Case "ARMADA"
            RestrictStringToByte = 2
        
        Case "CAOS"
            RestrictStringToByte = 3
        
        Case "FACCION"
            RestrictStringToByte = 4
        
        Case Else
            RestrictStringToByte = 0

    End Select

    '<EhFooter>
    Exit Function

RestrictStringToByte_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.RestrictStringToByte " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function RestrictByteToString(ByVal restrict As Byte) As String

    '<EhHeader>
    On Error GoTo RestrictByteToString_Err

    '</EhHeader>

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 04/18/2011
    '
    '***************************************************
    Select Case restrict

        Case 1
            RestrictByteToString = "NEWBIE"
        
        Case 2
            RestrictByteToString = "ARMADA"
        
        Case 3
            RestrictByteToString = "CAOS"
        
        Case 4
            RestrictByteToString = "FACCION"
        
        Case 0
            RestrictByteToString = "NO"

    End Select

    '<EhFooter>
    Exit Function

RestrictByteToString_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.RestrictByteToString " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function TerrainStringToByte(ByRef restrict As String) As Byte

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 04/18/2011
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo TerrainStringToByte_Err

    '</EhHeader>
    restrict = UCase$(restrict)

    Select Case restrict

        Case "NIEVE"
            TerrainStringToByte = 1
        
        Case "DESIERTO"
            TerrainStringToByte = 2
        
        Case "CIUDAD"
            TerrainStringToByte = 3
        
        Case "CAMPO"
            TerrainStringToByte = 4
        
        Case "DUNGEON"
            TerrainStringToByte = 5
        
        Case Else
            TerrainStringToByte = 0

    End Select

    '<EhFooter>
    Exit Function

TerrainStringToByte_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.TerrainStringToByte " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function TerrainByteToString(ByVal restrict As Byte) As String

    '<EhHeader>
    On Error GoTo TerrainByteToString_Err

    '</EhHeader>

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 04/18/2011
    '
    '***************************************************
    Select Case restrict

        Case 1
            TerrainByteToString = "NIEVE"
        
        Case 2
            TerrainByteToString = "DESIERTO"
        
        Case 3
            TerrainByteToString = "CIUDAD"
        
        Case 4
            TerrainByteToString = "CAMPO"
        
        Case 5
            TerrainByteToString = "DUNGEON"
        
        Case 0
            TerrainByteToString = "BOSQUE"

    End Select

    '<EhFooter>
    Exit Function

TerrainByteToString_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.TerrainByteToString " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub SetAreaResuTheNpc(ByVal iNpc As Integer)

    '<EhHeader>
    On Error GoTo SetAreaResuTheNpc_Err

    '</EhHeader>

    ' @@ Miqueas
    ' @@ 17-10-2015
    ' @@ Set Trigger in this NPC area
    Const Range = 5 ' @@ + 5 Tildes a la redonda del npc

    Dim X      As Long

    Dim Y      As Long
     
    Dim NpcPos As WorldPos
     
    NpcPos.Map = Npclist(iNpc).Pos.Map
    NpcPos.X = Npclist(iNpc).Pos.X
    NpcPos.Y = Npclist(iNpc).Pos.Y
        
    For X = NpcPos.X - Range To NpcPos.X + Range
        For Y = NpcPos.Y - Range To NpcPos.Y + Range

            If InMapBounds(NpcPos.Map, X, Y) Then
                If (MapData(NpcPos.Map, X, Y).trigger = NADA) Then
                    MapData(NpcPos.Map, X, Y).trigger = eTrigger.AutoResu

                End If

            End If

        Next Y
    Next X

    '<EhFooter>
    Exit Sub

SetAreaResuTheNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.SetAreaResuTheNpc " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub DeleteAreaResuTheNpc(ByVal iNpc As Integer)

    '<EhHeader>
    On Error GoTo DeleteAreaResuTheNpc_Err

    '</EhHeader>

    ' @@ Miqueas
    ' @@ 17-10-2015
    ' @@ Not Set Trigger in this NPC area
    Const Range = 5 ' @@ + 4 Tildes a la redonda del npc

    Dim X      As Long

    Dim Y      As Long
     
    Dim NpcPos As WorldPos
     
    NpcPos.Map = Npclist(iNpc).Pos.Map
    NpcPos.X = Npclist(iNpc).Pos.X
    NpcPos.Y = Npclist(iNpc).Pos.Y

    For X = NpcPos.X - Range To NpcPos.X + Range
        For Y = NpcPos.Y - Range To NpcPos.Y + Range

            If InMapBounds(NpcPos.Map, X, Y) Then
                If (MapData(NpcPos.Map, X, Y).trigger = eTrigger.AutoResu) Then
                    MapData(NpcPos.Map, X, Y).trigger = 0

                End If

            End If

        Next Y
    Next X

    '<EhFooter>
    Exit Sub

DeleteAreaResuTheNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.DeleteAreaResuTheNpc " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function IsAreaResu(ByVal UserIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo IsAreaResu_Err

    '</EhHeader>

    ' @@ Miqueas
    ' @@ 17/10/2015
    ' @@ Validate Trigger Area
    With UserList(UserIndex)

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.AutoResu Then
            IsAreaResu = True

            Exit Function

        End If

    End With

    IsAreaResu = False
    '<EhFooter>
    Exit Function

IsAreaResu_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.IsAreaResu " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub AutoCurar(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo AutoCurar_Err

    '</EhHeader>

    ' @@ Miqueas
    ' @@ 17-10-15
    ' @@ Zona de auto curacion
     
    With UserList(UserIndex)

        If .flags.Muerto = 1 Then
            Call RevivirUsuario(UserIndex)
            'Call WriteConsoleMsg(UserIndex, "El sacerdote te ha resucitado", FontTypeNames.FONTTYPE_INFO)
            GoTo Temp

        End If

        If .Stats.MinHp < .Stats.MaxHp Then
            .Stats.MinHp = .Stats.MaxHp
            Call WriteUpdateHP(UserIndex)

            'Call WriteConsoleMsg(UserIndex, "El sacerdote te ha curado.", FontTypeNames.FONTTYPE_INFO)
        End If

Temp:

        If .flags.Ceguera = 1 Then
            .flags.Ceguera = 0

        End If

        If .flags.Envenenado = 1 Then
            .flags.Envenenado = 0

        End If

    End With

    '<EhFooter>
    Exit Sub

AutoCurar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.AutoCurar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function isNPCResucitador(ByVal iNpc As Integer) As Boolean

    '<EhHeader>
    On Error GoTo isNPCResucitador_Err

    '</EhHeader>

    ' @@ Miqueas
    ' @@ 17/10/2015
    ' @@ Validate NPC
    With Npclist(iNpc)

        If (.NPCtype = eNPCType.Revividor) Or (.NPCtype = eNPCType.ResucitadorNewbie) Then
            isNPCResucitador = True

            Exit Function

        End If

    End With

    isNPCResucitador = False
    '<EhFooter>
    Exit Function

isNPCResucitador_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Extra.isNPCResucitador " & "at line " & Erl
        
    '</EhFooter>
End Function

Sub LoadPacketRatePolicy()

    On Error GoTo LoadPacketRatePolicy_Err

    Dim Lector As clsIniManager

    Dim i      As Long

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando PacketRatePolicy."
    
    Set Lector = New clsIniManager
    Call Lector.Initialize(IniPath & "PacketRatePolicy.ini")

    For i = 1 To MAX_PACKET_COUNTERS

        Dim PacketName As String

        PacketName = PacketIdToString(i)
        MacroIterations(i) = val(Lector.GetValue(PacketName, "Iterations"))
        PacketTimerThreshold(i) = val(Lector.GetValue(PacketName, "Limit"))
    Next i

    Set Lector = Nothing

    Exit Sub

LoadPacketRatePolicy_Err:
    Set Lector = Nothing
        
End Sub

Public Function PacketIdToString(ByVal PacketID As Long) As String

    Select Case PacketID

        Case 1
            PacketIdToString = "CastSpell"
            Exit Function

        Case 2
            PacketIdToString = "WorkLeftClick"
            Exit Function

        Case 3
            PacketIdToString = "LeftClick"
            Exit Function

        Case 4
            PacketIdToString = "UseItem"
            Exit Function

        Case 5
            PacketIdToString = "UseItemU"
            Exit Function

        Case 6
            PacketIdToString = "Walk"
            Exit Function

        Case 7
            PacketIdToString = "Sailing"
            Exit Function

        Case 8
            PacketIdToString = "Talk"
            Exit Function

        Case 9
            PacketIdToString = "Attack"
            Exit Function

        Case 10
            PacketIdToString = "Drop"
            Exit Function

        Case 11
            PacketIdToString = "Work"
            Exit Function

        Case 12
            PacketIdToString = "EquipItem"
            Exit Function

        Case 13
            PacketIdToString = "GuildMessage"
            Exit Function

        Case 14
            PacketIdToString = "QuestionGM"
            Exit Function

        Case 15
            PacketIdToString = "ChangeHeading"
            Exit Function

    End Select
    
End Function

Public Function HayPuerta(ByVal Map As Integer, _
                          ByVal X As Integer, _
                          ByVal Y As Integer) As Boolean

    If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
        HayPuerta = (ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas) And ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Cerrada = 1 And (ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0)

    End If

End Function

' Autor: WyroX - 20/01/2021
' Retorna el heading recibo como parámetro pero rotado, según el valor R.
' Si R es 1, rota en sentido horario. Si R es -1, en sentido antihorario.
Function Rotate_Heading(ByVal Heading As eHeading, ByVal r As Integer) As eHeading
    
    Rotate_Heading = (Heading + r + 3) Mod 4 + 1
    
End Function

Function Status(ByVal UserIndex As Integer) As eTipoFaction
        
    On Error GoTo Status_Err

    Status = UserList(UserIndex).Faction.Status
        
    Exit Function

Status_Err:
        
End Function
