Attribute VB_Name = "modSendData"
'**************************************************************
' SendData.bas - Has all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' Implemented by Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@gmail.com)
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

''
' Contains all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20070107
' Reparado por Lorwik

Option Explicit

Public Enum SendTarget

    ToAll = 1
    ToOne
    toMap
    toMapSecure
    ToPCArea
    ToAllButIndex
    ToMapButIndex
    ToGM
    ToNPCArea
    ToGuildMembers
    ToAdmins
    ToPCAreaButIndex
    ToAdminsAreaButConsejeros
    ToDiosesYclan
    ToConsejo
    ToConsejoYCaos
    ToClanArea
    ToConsejoCaos
    ToDeadArea
    ToCiudadanos
    ToCriminales
    ToPartyArea
    ToReal
    ToCaos
    ToCiudadanosYRMs
    ToCriminalesYRMs
    ToRealYRMs
    ToCaosYRMs
    ToHigherAdmins
    ToGMsAreaButRmsOrCounselors
    ToUsersAreaButGMs
    ToUsersAndRmsAndCounselorsAreaButGMs
    ToFaction

End Enum

Public Sub SendData(ByVal sndRoute As SendTarget, _
                    ByVal sndIndex As Integer, _
                    ByVal sndData As String, _
                    Optional ByVal IsDenounce As Boolean = False, _
                    Optional ByVal IsUrgent As Boolean = False)
        
    '<EhHeader>
    On Error GoTo OnError

    '</EhHeader>

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus) - Rewrite of original
    'Last Modify Date: 14/11/2010
    'Last modified by: ZaMa
    '14/11/2010: ZaMa - Now denounces can be desactivated.
    '**************************************************************

    Dim LoopC As Long
    
    frmMain.lstDebug.AddItem " > [" & sndIndex & "] Data (" & Len(sndData) & "): " & sndData
    
    Select Case sndRoute

        Case SendTarget.ToOne

            If UserList(sndIndex).ConnIDValida Then
                Call Server.Send(sndIndex, IsUrgent, Writer)

            End If
            
        Case SendTarget.ToPCArea
            Call SendToUserArea(sndIndex, sndData)
        
        Case SendTarget.ToGM

            For LoopC = 1 To LastUser

                If UserList(LoopC).ConnIDValida Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then

                        ' Denounces can be desactivated
                        If IsDenounce Then
                            If UserList(LoopC).flags.SendDenounces Then
                                Call Server.Send(LoopC, False, Writer)

                            End If

                        Else
                            Call Server.Send(LoopC, False, Writer)

                        End If

                    End If

                End If

            Next LoopC
            
        Case SendTarget.ToAdmins

            For LoopC = 1 To LastUser

                If UserList(LoopC).ConnIDValida Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin) Then
                        If EsGmPriv(LoopC) Then

                            ' Denou(ces can be desactivated
                            If IsDenounce Then
                                If UserList(LoopC).flags.SendDenounces Then
                                    Call Server.Send(LoopC, False, Writer)

                                End If
    
                            Else
                                Call Server.Send(LoopC, False, Writer)

                            End If

                        End If

                    End If

                End If

            Next LoopC
        
        Case SendTarget.ToAll

            For LoopC = 1 To LastUser

                If UserList(LoopC).ConnIDValida Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call Server.Send(LoopC, False, Writer)

                    End If

                End If

            Next LoopC
        
        Case SendTarget.ToAllButIndex

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnIDValida) And (LoopC <> sndIndex) Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call Server.Send(LoopC, False, Writer)

                    End If

                End If

            Next LoopC
        
        Case SendTarget.toMap
            Call SendToMap(sndIndex, sndData)
        
        Case SendTarget.toMapSecure

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnIDValida) And (LoopC <> sndIndex) Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        If Not MapInfo(UserList(LoopC).Pos.Map).Pk Then
                            Call Server.Send(LoopC, False, Writer)

                        End If

                    End If

                End If

            Next LoopC
                    
        Case SendTarget.ToGuildMembers
            'LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

            'While LoopC > 0

            'If (UserList(LoopC).ConnIDValida) Then
            'Call Server.send(LoopC, false, Writer)
            'End If

            'LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

            'Wend
        
        Case SendTarget.ToDeadArea
            Call SendToDeadUserArea(sndIndex, sndData)
        
        Case SendTarget.ToPCAreaButIndex
            Call SendToUserAreaButindex(sndIndex, sndData)
        
        Case SendTarget.ToClanArea
            Call SendToUserGuildArea(sndIndex, sndData)
        
        Case SendTarget.ToPartyArea
            Call SendToUserPartyArea(sndIndex, sndData)
        
        Case SendTarget.ToAdminsAreaButConsejeros
            Call SendToAdminsButConsejerosArea(sndIndex, sndData)
        
        Case SendTarget.ToNPCArea
            Call SendToNpcArea(sndIndex, sndData)

        Case SendTarget.ToDiosesYclan

            For LoopC = 1 To MAX_GUILD_MEMBER
                
                If (UserList(LoopC).ConnIDValida) And (GuildsInfo(sndIndex).Members(LoopC).UserIndex > 0) Then
                    If UserList(GuildsInfo(sndIndex).Members(LoopC).UserIndex).flags.UserLogged Then 'Esta logeado como usuario?
                        Call Server.Send(GuildsInfo(sndIndex).Members(LoopC).UserIndex, False, Writer)

                    End If

                End If

            Next LoopC

        Case SendTarget.ToConsejoYCaos

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnIDValida) Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.RoyalCouncil Or PlayerType.RoyalCouncil) Then
                        Call Server.Send(LoopC, False, Writer)

                    End If

                End If

            Next LoopC
            
        Case SendTarget.ToConsejo

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnIDValida) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call Server.Send(LoopC, False, Writer)

                    End If

                End If

            Next LoopC
        
        Case SendTarget.ToConsejoCaos

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnIDValida) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call Server.Send(LoopC, False, Writer)

                    End If

                End If

            Next LoopC
        
        Case SendTarget.ToCiudadanos

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnIDValida) Then
                    If Not Escriminal(LoopC) Then
                        Call Server.Send(LoopC, False, Writer)

                    End If

                End If

            Next LoopC
        
        Case SendTarget.ToCriminales

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnIDValida) Then
                    If Escriminal(LoopC) Then
                        Call Server.Send(LoopC, False, Writer)

                    End If

                End If

            Next LoopC
        
        Case SendTarget.ToReal

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnIDValida) Then
                    If UserList(LoopC).Faction.Status = r_Armada Then
                        Call Server.Send(LoopC, False, Writer)

                    End If

                End If

            Next LoopC
        
        Case SendTarget.ToCaos

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnIDValida) Then
                    If UserList(LoopC).Faction.Status = r_Caos Then
                        Call Server.Send(LoopC, False, Writer)

                    End If

                End If

            Next LoopC
        
        Case SendTarget.ToCiudadanosYRMs

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnIDValida) Then
                    If Not Escriminal(LoopC) Then
                        Call Server.Send(LoopC, False, Writer)

                    End If

                End If

            Next LoopC
        
        Case SendTarget.ToCriminalesYRMs

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnIDValida) Then
                    If Escriminal(LoopC) Then
                        Call Server.Send(LoopC, False, Writer)

                    End If

                End If

            Next LoopC
        
        Case SendTarget.ToRealYRMs

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnIDValida) Then
                    If UserList(LoopC).Faction.Status = r_Armada Then
                        Call Server.Send(LoopC, False, Writer)

                    End If

                End If

            Next LoopC
        
        Case SendTarget.ToCaosYRMs

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnIDValida) Then
                    If UserList(LoopC).Faction.Status = r_Caos Then
                        Call Server.Send(LoopC, False, Writer)

                    End If

                End If

            Next LoopC
        
        Case SendTarget.ToHigherAdmins

            For LoopC = 1 To LastUser

                If UserList(LoopC).ConnIDValida Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
                        Call Server.Send(LoopC, False, Writer)

                    End If

                End If

            Next LoopC
            
        Case SendTarget.ToGMsAreaButRmsOrCounselors
            Call SendToGMsAreaButRmsOrCounselors(sndIndex, sndData)
            
        Case SendTarget.ToUsersAreaButGMs
            Call SendToUsersAreaButGMs(sndIndex, sndData)

        Case SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs
            Call SendToUsersAndRmsAndCounselorsAreaButGMs(sndIndex, sndData)
        
        Case SendTarget.ToFaction
            Call SendToUsersFaction(sndIndex, sndData)

    End Select
    
OnError:
    Writer.Clear
        
    If Err.number <> 0 Then
        LogError Err.description & vbCrLf & "in ServidorArgentum.modSendData.SendData " & "at line " & Erl

    End If

End Sub

Private Sub SendToUserArea(ByVal UserIndex As Integer, _
                           ByVal sdData As String, _
                           Optional ByVal IsUrgent As Boolean = False)

    '<EhHeader>
    On Error GoTo SendToUserArea_Err

    '</EhHeader>
    Dim query() As Collision.UUID

    Dim i       As Long
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
        Call Server.Send(query(i).Name, IsUrgent, Writer)
    Next i
    
    Call Server.Send(UserIndex, IsUrgent, Writer)
    '<EhFooter>
    Exit Sub

SendToUserArea_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSendData.SendToUserArea " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, _
                                   ByVal sdData As String, _
                                   Optional ByVal IsUrgent As Boolean = False)

    '<EhHeader>
    On Error GoTo SendToUserAreaButindex_Err

    '</EhHeader>
    Dim query() As Collision.UUID

    Dim i       As Long
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
        Call Server.Send(query(i).Name, IsUrgent, Writer)
    Next i

    '<EhFooter>
    Exit Sub

SendToUserAreaButindex_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSendData.SendToUserAreaButindex " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub SendToDeadUserArea(ByVal UserIndex As Integer, _
                               ByVal sdData As String, _
                               Optional ByVal IsUrgent As Boolean = False)

    '<EhHeader>
    On Error GoTo SendToDeadUserArea_Err

    '</EhHeader>
    Dim query() As Collision.UUID

    Dim i       As Long
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)

        With UserList(query(i).Name)

            If (.flags.Muerto = 1 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0) Then
                Call Server.Send(query(i).Name, IsUrgent, Writer)

            End If

        End With

    Next i

    Call Server.Send(UserIndex, IsUrgent, Writer)
    '<EhFooter>
    Exit Sub

SendToDeadUserArea_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSendData.SendToDeadUserArea " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, _
                                ByVal sdData As String, _
                                Optional ByVal IsUrgent As Boolean = False)

    '<EhHeader>
    On Error GoTo SendToUserGuildArea_Err

    '</EhHeader>
    Dim query()    As Collision.UUID

    Dim i          As Long

    Dim GuildIndex As Integer
    
    GuildIndex = UserList(UserIndex).GuildIndex
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)

        With UserList(query(i).Name)

            If (.GuildIndex = GuildIndex Or (.flags.Privilegios And PlayerType.Dios)) Then
                Call Server.Send(query(i).Name, IsUrgent, Writer)

            End If

        End With

    Next i

    Call Server.Send(UserIndex, IsUrgent, Writer)
    '<EhFooter>
    Exit Sub

SendToUserGuildArea_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSendData.SendToUserGuildArea " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub SendToUserPartyArea(ByVal UserIndex As Integer, _
                                ByVal sdData As String, _
                                Optional ByVal IsUrgent As Boolean = False)

    '<EhHeader>
    On Error GoTo SendToUserPartyArea_Err

    '</EhHeader>
    Dim query()    As Collision.UUID

    Dim i          As Long

    Dim GroupIndex As Long

    GroupIndex = UserList(UserIndex).GroupIndex

    If GroupIndex = 0 Then Exit Sub
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)

        If (UserList(query(i).Name).GroupIndex = GroupIndex) Then
            Call Server.Send(query(i).Name, IsUrgent, Writer)

        End If

    Next i

    Call Server.Send(UserIndex, IsUrgent, Writer)
    '<EhFooter>
    Exit Sub

SendToUserPartyArea_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSendData.SendToUserPartyArea " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub SendToAdminsButConsejerosArea(ByVal UserIndex As Integer, _
                                          ByVal sdData As String, _
                                          Optional ByVal IsUrgent As Boolean = False)

    '<EhHeader>
    On Error GoTo SendToAdminsButConsejerosArea_Err

    '</EhHeader>
    Dim query() As Collision.UUID

    Dim i       As Long
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)

        If (UserList(query(i).Name).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
            Call Server.Send(query(i).Name, IsUrgent, Writer)

        End If

    Next i
    
    If (UserList(UserIndex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
        Call Server.Send(UserIndex, IsUrgent, Writer)

    End If

    '<EhFooter>
    Exit Sub

SendToAdminsButConsejerosArea_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSendData.SendToAdminsButConsejerosArea " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub SendToNpcArea(ByVal NpcIndex As Long, _
                          ByVal sdData As String, _
                          Optional ByVal IsUrgent As Boolean = False)

    '<EhHeader>
    On Error GoTo SendToNpcArea_Err

    '</EhHeader>
    Dim query() As Collision.UUID

    Dim i       As Long
    
    For i = 0 To ModAreas.QueryObservers(NpcIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)
        Call Server.Send(query(i).Name, IsUrgent, Writer)
    Next i

    '<EhFooter>
    Exit Sub

SendToNpcArea_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSendData.SendToNpcArea " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub SendToAreaByPos(ByVal Map As Integer, _
                           ByVal AreaX As Integer, _
                           ByVal AreaY As Integer, _
                           ByVal sdData As String, _
                           Optional ByVal IsUrgent As Boolean = False)

    '<EhHeader>
    On Error GoTo SendToAreaByPos_Err

    '</EhHeader>

    Dim query() As Collision.UUID

    Dim i       As Long

    Dim ItemID  As Long

    ItemID = Pack(Map, AreaX, AreaY)
    
    For i = 0 To ModAreas.QueryObservers(ItemID, ENTITY_TYPE_OBJECT, query, ENTITY_TYPE_PLAYER)
        Call Server.Send(query(i).Name, IsUrgent, Writer)
    Next i
    
    '<EhFooter>
    Exit Sub

SendToAreaByPos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSendData.SendToAreaByPos " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub SendToMap(ByVal Map As Integer, _
                     ByVal sdData As String, _
                     Optional ByVal IsUrgent As Boolean = False)

    '<EhHeader>
    On Error GoTo SendToMap_Err

    '</EhHeader>
    Call Server.Broadcast(MapInfo(Map).Players, IsUrgent, Writer)
    '<EhFooter>
    Exit Sub

SendToMap_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSendData.SendToMap " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub SendToGMsAreaButRmsOrCounselors(ByVal UserIndex As Integer, _
                                            ByVal sdData As String, _
                                            Optional ByVal IsUrgent As Boolean = False)

    '<EhHeader>
    On Error GoTo SendToGMsAreaButRmsOrCounselors_Err

    '</EhHeader>
    Dim query() As Collision.UUID

    Dim i       As Long
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)

        With UserList(query(i).Name)

            If ((.flags.Privilegios And Not PlayerType.User) = .flags.Privilegios) Then
                Call Server.Send(query(i).Name, IsUrgent, Writer)

            End If

        End With

    Next i
    
    With UserList(UserIndex)

        If ((.flags.Privilegios And Not PlayerType.User) = .flags.Privilegios) Then
            Call Server.Send(UserIndex, IsUrgent, Writer)

        End If

    End With

    '<EhFooter>
    Exit Sub

SendToGMsAreaButRmsOrCounselors_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSendData.SendToGMsAreaButRmsOrCounselors " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub SendToUsersAreaButGMs(ByVal UserIndex As Integer, _
                                  ByVal sdData As String, _
                                  Optional ByVal IsUrgent As Boolean = False)

    '<EhHeader>
    On Error GoTo SendToUsersAreaButGMs_Err

    '</EhHeader>
    Dim query() As Collision.UUID

    Dim i       As Long
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)

        If (UserList(query(i).Name).flags.Privilegios And PlayerType.User) Then
            Call Server.Send(query(i).Name, IsUrgent, Writer)

        End If

    Next i

    If (UserList(UserIndex).flags.Privilegios And PlayerType.User) Then
        Call Server.Send(UserIndex, IsUrgent, Writer)

    End If

    '<EhFooter>
    Exit Sub

SendToUsersAreaButGMs_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSendData.SendToUsersAreaButGMs " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub SendToUsersAndRmsAndCounselorsAreaButGMs(ByVal UserIndex As Integer, _
                                                     ByVal sdData As String, _
                                                     Optional ByVal IsUrgent As Boolean = False)

    '<EhHeader>
    On Error GoTo SendToUsersAndRmsAndCounselorsAreaButGMs_Err

    '</EhHeader>
    Dim query() As Collision.UUID

    Dim i       As Long
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)

        If (UserList(query(i).Name).flags.Privilegios And (PlayerType.User)) Then
            Call Server.Send(query(i).Name, IsUrgent, Writer)

        End If

    Next i

    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
        Call Server.Send(UserIndex, IsUrgent, Writer)

    End If

    '<EhFooter>
    Exit Sub

SendToUsersAndRmsAndCounselorsAreaButGMs_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSendData.SendToUsersAndRmsAndCounselorsAreaButGMs " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub SendToUsersFaction(ByVal UserIndex As Integer, _
                               ByVal sdData As String, _
                               Optional ByVal IsUrgent As Boolean = False)

    '<EhHeader>
    On Error GoTo SendToUsersAreaButGMs_Err

    '</EhHeader>
    Dim query() As Collision.UUID

    Dim i       As Long
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)

        If (UserList(query(i).Name).Faction.Status = UserList(UserIndex).Faction.Status) Then
            Call Server.Send(query(i).Name, IsUrgent, Writer)

        End If

    Next i

    If (UserList(UserIndex).flags.Privilegios And PlayerType.User) Then
        Call Server.Send(UserIndex, IsUrgent, Writer)

    End If

    '<EhFooter>
    Exit Sub

SendToUsersAreaButGMs_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSendData.SendToUsersAreaButGMs " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub AlertarFaccionarios(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo AlertarFaccionarios_Err

    '</EhHeader>

    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 17/11/2009
    'Alerta a los faccionarios, dandoles una orientacion
    '**************************************************************
    Dim TempIndex As Integer

    Dim i         As Long

    Dim Font      As FontTypeNames

    Dim query()   As Collision.UUID
        
    With UserList(UserIndex)

        If esCaos(UserIndex) Then
            Font = FontTypeNames.FONTTYPE_CONSEJOCAOS
        Else
            Font = FontTypeNames.FONTTYPE_CONSEJO

        End If
            
        Call SendData(SendTarget.ToFaction, UserIndex, PrepareMessageConsoleMsg("Escuchas el llamado de un líder faccionario que proviene de " & MapInfo(.Pos.Map).Name & " (" & .Pos.Map & " " & .Pos.X & " " & .Pos.Y & ")", Font))

    End With

    '<EhFooter>
    Exit Sub

AlertarFaccionarios_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSendData.AlertarFaccionarios " & "at line " & Erl

    '</EhFooter>
End Sub
