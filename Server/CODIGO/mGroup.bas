Attribute VB_Name = "mGroup"
Option Explicit

Public Const MAX_MEMBERS_GROUP       As Byte = 5

Public Const MAX_REQUESTS_GROUP      As Byte = 10

Private Const MAX_GROUPS             As Byte = 100

Private Const SLOT_LEADER            As Byte = 1

Private Const GROUPS_MAXDISTANCIA    As Byte = 18

Private Const GROUPS_MIN_LEVEL_FOUND As Byte = 15

Private Const GROUPS_MIN_LEVEL       As Byte = 13

Private Const GROUPS_DIFERENCE_LEVEL As Byte = 3      ' Diferencia de Nivel para poder cambiar los porcentajes...

Private Const GROUPS_REQUEST_TIME    As Integer = 20000      ' Diferencia de Nivel para poder cambiar los porcentajes...

Public Enum eBonusGroup

    GroupFull = 1
    LeaderPremium = 2
    LeaderPendient = 3
    LeaderMaxLevel = 4

End Enum

Public Type tUserGroup

    Index As Integer
    Exp As Long
    PorcExp As Byte

End Type

Public Type tGroups

    Members As Byte
    User(1 To MAX_MEMBERS_GROUP) As tUserGroup
    Requests(1 To MAX_REQUESTS_GROUP) As String
    Acumular As Boolean

End Type

Public Groups(1 To MAX_GROUPS) As tGroups

' Buscamos un SLOT LIBRE para CREAR GRUPO (HASTA 100)
Private Function FreeGroup() As Byte

    '<EhHeader>
    On Error GoTo FreeGroup_Err

    '</EhHeader>

    Dim A As Long
    
    For A = 1 To MAX_GROUPS

        If Groups(A).User(SLOT_LEADER).Index = 0 Then
            FreeGroup = A

            Exit For

        End If

    Next A

    '<EhFooter>
    Exit Function

FreeGroup_Err:
    LogError Err.description & vbCrLf & "in FreeGroup " & "at line " & Erl

    '</EhFooter>
End Function

' Buscamos un SLOT LIBRE para meter un NUEVO MIEMBRO.
Private Function FreeGroupMember(ByVal GroupIndex As Byte) As Byte

    '<EhHeader>
    On Error GoTo FreeGroupMember_Err

    '</EhHeader>

    Dim A As Long
    
    For A = 1 To MAX_MEMBERS_GROUP

        If Groups(GroupIndex).User(A).Index = 0 Then
            FreeGroupMember = A

            Exit For

        End If

    Next A

    '<EhFooter>
    Exit Function

FreeGroupMember_Err:
    LogError Err.description & vbCrLf & "in FreeGroupMember " & "at line " & Erl

    '</EhFooter>
End Function

Private Sub SetGroupMember(ByVal GroupIndex As Integer, _
                           ByVal SlotMember As Byte, _
                           ByVal UserIndex As Integer, _
                           ByVal PorcExp As Byte)

    '<EhHeader>
    On Error GoTo SetGroupMember_Err

    '</EhHeader>

    With Groups(GroupIndex)
        .User(SlotMember).Index = UserIndex
        .User(SlotMember).PorcExp = PorcExp
        
        'Call SendData(SendTarget.ToPartyArea, UserIndex, PrepareMessageUpdateGroupIndex(UserList(UserIndex).Char.CharIndex, GroupIndex))
    End With

    '<EhFooter>
    Exit Sub

SetGroupMember_Err:
    LogError Err.description & vbCrLf & "in SetGroupMember " & "at line " & Erl

    '</EhFooter>
End Sub

' Creamos un NUEVO GRUPO.
Public Sub CreateGroup(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo CreateGroup_Err

    '</EhHeader>

    Dim Slot As Byte
    
    With UserList(UserIndex)
        
        If EsGm(UserIndex) Then Exit Sub
        
        If .GroupIndex > 0 Then
            WriteConsoleMsg UserIndex, "Ya perteneces a una party.", FontTypeNames.FONTTYPE_INFO

            Exit Sub

        End If
        
        If .Stats.Elv < GROUPS_MIN_LEVEL_FOUND Then
            WriteConsoleMsg UserIndex, "Tu nivel no te permite crear una Party.", FontTypeNames.FONTTYPE_INFO

            Exit Sub

        End If
        
        Slot = FreeGroup
        
        If Slot > 0 Then
            SetGroupMember Slot, SLOT_LEADER, UserIndex, 100
            Groups(Slot).Members = SLOT_LEADER
            Groups(Slot).Acumular = True ' Por defecto acumula la experiencia...
            UserList(UserIndex).GroupIndex = Slot
            UserList(UserIndex).GroupSlotUser = SLOT_LEADER
            
            WriteConsoleMsg UserIndex, "¡Eres el lider de un nuevo grupo! Debes invitar a usuarios haciendo clic sobre ellos y tecleando F3", FontTypeNames.FONTTYPE_INFO
            WriteGroupPrincipal UserIndex
        Else
            WriteConsoleMsg UserIndex, "Existen demasiados grupos creados, por favor espera que se disuelva alguno.", FontTypeNames.FONTTYPE_INFO

        End If
    
    End With

    '<EhFooter>
    Exit Sub

CreateGroup_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mGroup.CreateGroup " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Enviamos solicitud a UN GRUPO
Public Sub SendInvitationGroup(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo SendInvitationGroup_Err

    '</EhHeader>

    ' Un personaje decide solicitar entrar a una party.
    Dim tUser As Integer

    Dim Slot  As Byte

    Dim Time  As Long
    
    Time = GetTime
    
    With UserList(UserIndex)
    
        If .GroupIndex = 0 Then
            WriteConsoleMsg UserIndex, "Debes pertenecer a una party para poder invitar usuarios...", FontTypeNames.FONTTYPE_INFO

            Exit Sub

        End If
            
        If Groups(.GroupIndex).Members = MAX_MEMBERS_GROUP Then
            WriteConsoleMsg UserIndex, "La party está llena.", FontTypeNames.FONTTYPE_INFO

            Exit Sub

        End If
        
        tUser = .flags.TargetUser
        
        If tUser > 0 Then

            With UserList(tUser)
                
                ' Pertenece a un Grupo
                If .GroupIndex > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El personaje ya pertenece a una party.", FontTypeNames.FONTTYPE_INFORED)
                    Exit Sub

                End If
                
                ' Ya le hemos enviado una solicitud al flaco
                If StrComp(.GroupRequest, UCase$(UserList(UserIndex).Name)) = 0 Then
                    Call WriteConsoleMsg(UserIndex, "El personaje ya ha recibido una solicitud de tu parte. Aguarda respuesta...", FontTypeNames.FONTTYPE_INFORED)
                    Exit Sub

                End If
                
                ' Anti Spam de Solicitudes
                If Time - .GroupRequestTime < GROUPS_REQUEST_TIME Then
                    Call WriteConsoleMsg(UserIndex, "El usuario necesita esperar algunos segundos para recibir una solicitud de grupo.", FontTypeNames.FONTTYPE_INFORED)
                    Exit Sub

                End If
                
                ' No tiene el nivel correspondiente
                If .Stats.Elv < GROUPS_MIN_LEVEL Then
                    WriteConsoleMsg UserIndex, "El personaje no puede pertenecer a un grupo por su nivel.", FontTypeNames.FONTTYPE_INFORED

                    Exit Sub

                End If
                    
                .GroupRequest = UCase$(UserList(UserIndex).Name)
                .GroupRequestTime = Time
                
                WriteConsoleMsg UserIndex, "Has invitado a " & .Name & " para que se una a tu grupo. Espera pronta respuesta...", FontTypeNames.FONTTYPE_INFO
                WriteConsoleMsg tUser, "El personaje " & UserList(UserIndex).Name & " te ha invitado a formar un grupo. Tipea /SIPARTY para confirmar su invitación...", FontTypeNames.FONTTYPE_INFO
            
            End With
        
        End If
    
    End With

    '<EhFooter>
    Exit Sub

SendInvitationGroup_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mGroup.SendInvitationGroup " & "at line " & Erl
        
    '</EhFooter>
End Sub

' El personaje acepta la invitación para conformar una party
Public Sub AcceptInvitationGroup(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo AcceptInvitationGroup_Err

    '</EhHeader>

    Dim Slot       As Byte

    Dim tUser      As Integer
    
    Dim GroupIndex As Integer
    
    With UserList(UserIndex)

        If .GroupIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Ya perteneces a un grupo y debes salir del mismo para participar en otro...", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If
        
        If .GroupRequest = vbNullString Then
            Call WriteConsoleMsg(UserIndex, "Nadie te ha invitado a formar un grupo.. Vaya parece que necesitas amigos...", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If
        
        tUser = NameIndex(.GroupRequest)
        
        If tUser <= 0 Then
            .GroupRequest = vbNullString
            .GroupRequestTime = 0
            Call WriteConsoleMsg(UserIndex, "Vaya.. Parece ser que te han dejado solo ¡El lider ha escapado!", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub
        Else
            GroupIndex = UserList(tUser).GroupIndex

            If GroupIndex <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Vaya.. Parece ser que te han dejado solo ¡El lider ha escapado!", FontTypeNames.FONTTYPE_INFORED)
                .GroupRequest = vbNullString
                .GroupRequestTime = 0
                Exit Sub

            End If
                
            If UserList(tUser).flags.SlotEvent > 0 Then
                Call WriteConsoleMsg(UserIndex, "Vaya.. Parece ser que te han dejado solo ¡El lider ha escapado!", FontTypeNames.FONTTYPE_INFORED)
                .GroupRequest = vbNullString
                .GroupRequestTime = 0
                Exit Sub
                
            End If

        End If
        
        If Groups(GroupIndex).Members = MAX_MEMBERS_GROUP Then
            WriteConsoleMsg UserIndex, "La party está llena. Busca formar un grupo con otros usuarios...", FontTypeNames.FONTTYPE_INFO

            Exit Sub

        End If
        
        Slot = FreeGroupMember(GroupIndex)
        
        Groups(GroupIndex).User(Slot).Index = UserIndex
        Groups(GroupIndex).Members = Groups(GroupIndex).Members + 1
            
        .GroupSlotUser = Slot
        .GroupIndex = GroupIndex
            
        SendMessageGroup .GroupIndex, UserList(tUser).Name, "El personaje " & .Name & " se ha unido al grupo ¡Bienvenido!"
            
        UpdatePorcentaje GroupIndex

    End With

    '<EhFooter>
    Exit Sub

AcceptInvitationGroup_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mGroup.AcceptInvitationGroup " & "at line " & Erl

    '</EhFooter>
End Sub

Private Function CheckGroupMap(ByVal GroupIndex As Byte) As Boolean

    '<EhHeader>
    On Error GoTo CheckGroupMap_Err

    '</EhHeader>

    Dim A As Byte
    
    CheckGroupMap = True
    
    For A = 1 To MAX_MEMBERS_GROUP

        With Groups(GroupIndex)

            If .User(A).Index > 0 Then
                If UserList(.User(A).Index).Pos.Map <> UserList(.User(SLOT_LEADER).Index).Pos.Map Then
                    CheckGroupMap = False

                    Exit For

                End If

            End If

        End With

    Next A

    '<EhFooter>
    Exit Function

CheckGroupMap_Err:
    LogError Err.description & vbCrLf & "in CheckGroupMap " & "at line " & Erl

    '</EhFooter>
End Function

' El grupo acumula experiencia. // Al morir da el ORO
Public Sub AddExpGroup(ByVal UserIndex As Integer, _
                       ByRef Exp As Long, _
                       Optional ByVal Gld As Long = 0)

    '<EhHeader>
    On Error GoTo AddExpGroup_Err

    '</EhHeader>
    
    Dim A              As Long

    Dim TempExp        As Long

    Dim ExpTemp        As Long

    Dim GldTemp        As Long
    
    Dim GroupIndex     As Byte
    
    Dim PendienteGroup As Boolean
    
    GroupIndex = UserList(UserIndex).GroupIndex
         
    With Groups(GroupIndex)

        For A = 1 To MAX_MEMBERS_GROUP
                
            If .User(A).Index > 0 Then

                ' Tiene un porcentaje inválido con pendiente bug?
                If .User(A).PorcExp >= 70 Then
                    If UserList(.User(A).Index).Invent.PendientePartyObjIndex = 0 Then
                        
                    End If

                End If
                    
                If Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(.User(A).Index).Pos.X, UserList(.User(A).Index).Pos.Y) <= GROUPS_MAXDISTANCIA And UserList(UserIndex).Pos.Map = UserList(.User(A).Index).Pos.Map And (UserList(.User(A).Index).flags.Muerto = 0) Then
                        
                    If Exp > 0 Then

                        ' // NUEVO
                        If UserIndex = .User(A).Index Then
                            Call CalcularDarExp_Bonus_Party(UserIndex, A, Exp)

                        End If
                  
                        TempExp = Porcentaje(Exp, .User(A).PorcExp)

                        If PartyTime Then
                            If .Members > 1 Then
                                TempExp = TempExp + (TempExp * 0.25)

                            End If

                        End If

                        If .Acumular Then
                            .User(A).Exp = .User(A).Exp + TempExp
                        Else
                            UserList(.User(A).Index).Stats.Exp = UserList(.User(A).Index).Stats.Exp + TempExp
                            Call WriteUpdateExp(.User(A).Index)
                            Call CheckUserLevel(.User(A).Index)

                        End If
                                               
                        If .User(A).Exp > MAXEXP Then
                            SaveExpAndGldMember GroupIndex, .User(A).Index

                        End If

                        Call SendData(SendTarget.ToOne, .User(A).Index, PrepareMessageRenderConsole("Exp " & .User(A).PorcExp & "% +" & CStr(Format(TempExp, "###,###,###")), d_Exp, 3000, 0))
                            
                    End If
                        
                    If Gld > 0 Then
                        GldTemp = Porcentaje(Gld, .User(A).PorcExp)
                        UserList(.User(A).Index).Stats.Gld = UserList(.User(A).Index).Stats.Gld + GldTemp
                        Call WriteUpdateGold(.User(A).Index)
                        Call SendData(SendTarget.ToOne, .User(A).Index, PrepareMessageRenderConsole("Oro " & .User(A).PorcExp & "% +" & CStr(Format(GldTemp, "###,###,###")), d_Exp, 3000, 0))

                    End If
                    
                    Call WriteGroupUpdateExp(.User(A).Index, GroupIndex)

                End If
                    
            End If
            
        Next A

    End With

    '<EhFooter>
    Exit Sub

AddExpGroup_Err:
    LogError Err.description & vbCrLf & "in AddExpGroup " & "at line " & Erl

    '</EhFooter>
End Sub

' Distribuimos las experiencias de las partys
Public Sub DistributeExpAndGldGroups()

    '<EhHeader>
    On Error GoTo DistributeExpAndGldGroups_Err

    '</EhHeader>

    Dim A As Long, B As Long
    
    For A = 1 To MAX_GROUPS

        With Groups(A)

            If .User(SLOT_LEADER).Index > 0 Then

                For B = 1 To MAX_MEMBERS_GROUP

                    If .User(B).Index > 0 Then
                        SaveExpAndGldMember A, .User(B).Index

                    End If

                Next B

            End If

        End With

    Next A

    '<EhFooter>
    Exit Sub

DistributeExpAndGldGroups_Err:
    LogError Err.description & vbCrLf & "in DistributeExpAndGldGroups " & "at line " & Erl

    '</EhFooter>
End Sub

' Actualizamos Experiencia y Oro del personaje.
Public Sub SaveExpAndGldMember(ByVal GroupIndex As Byte, ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo SaveExpAndGldMember_Err

    '</EhHeader>
    
    Dim SlotUser As Byte
    
    With UserList(UserIndex)
        SlotUser = .GroupSlotUser

        .Stats.Exp = .Stats.Exp + Groups(GroupIndex).User(SlotUser).Exp
        
        If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
        
        CheckUserLevel UserIndex
        
        WriteConsoleMsg UserIndex, "Hemos actualizado tu experiencia y Oro. Has conseguido " & Groups(GroupIndex).User(SlotUser).Exp & " puntos de experiencia.", FontTypeNames.FONTTYPE_GUILD
        Groups(GroupIndex).User(SlotUser).Exp = 0

    End With

    '<EhFooter>
    Exit Sub

SaveExpAndGldMember_Err:
    LogError Err.description & vbCrLf & "in SaveExpAndGldMember " & "at line " & Erl

    '</EhFooter>
End Sub

' Enviamos un mensaje al grupo.
Public Sub SendMessageGroup(ByVal GroupIndex As Byte, _
                            ByVal Emisor As String, _
                            ByVal Message As String)

    '<EhHeader>
    On Error GoTo SendMessageGroup_Err

    '</EhHeader>

    Dim A As Long

    For A = 1 To MAX_MEMBERS_GROUP

        With Groups(GroupIndex)

            If .User(A).Index > 0 Then
                If Emisor <> vbNullString Then
                    WriteConsoleMsg .User(A).Index, Emisor & "» " & Message, FontTypeNames.FONTTYPE_PARTY
                Else
                    WriteConsoleMsg .User(A).Index, Message, FontTypeNames.FONTTYPE_PARTY

                End If

            End If

        End With

    Next A

    '<EhFooter>
    Exit Sub

SendMessageGroup_Err:
    LogError Err.description & vbCrLf & "in SendMessageGroup " & "at line " & Erl

    '</EhFooter>
End Sub

' Reiniciamos la información de un miembro del Grupo
Private Sub ResetMemberGroup(ByVal GroupIndex As Byte, ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo ResetMemberGroup_Err

    '</EhHeader>

    With Groups(GroupIndex)
        
        ' Asignamos Experiencia y Oro obtenido hasta el momento
        mGroup.SaveExpAndGldMember GroupIndex, UserIndex
        
        .Members = .Members - 1
        .User(UserList(UserIndex).GroupSlotUser).Index = 0
        .User(UserList(UserIndex).GroupSlotUser).Exp = 0
        .User(UserList(UserIndex).GroupSlotUser).PorcExp = 0
        
    End With
    
    With UserList(UserIndex)
        .GroupIndex = 0
        .GroupRequest = 0
        .GroupRequestTime = 0
        .GroupSlotUser = 0
            
        'Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateGroupIndex(.Char.CharIndex, 0))
    End With
    
    UpdatePorcentaje GroupIndex
    '<EhFooter>
    Exit Sub

ResetMemberGroup_Err:
    LogError Err.description & vbCrLf & "in ResetMemberGroup " & "at line " & Erl

    '</EhFooter>
End Sub

' Reiniciamos la información del grupo
Private Sub ResetGroup(ByVal GroupIndex As Byte)

    '<EhHeader>
    On Error GoTo ResetGroup_Err

    '</EhHeader>

    Dim A As Long
    
    With Groups(GroupIndex)
        
        For A = 1 To MAX_MEMBERS_GROUP

            If .User(A).Index > 0 Then
                ResetMemberGroup GroupIndex, .User(A).Index

            End If

        Next A
        
        For A = 1 To MAX_REQUESTS_GROUP
            .Requests(A) = vbNullString
        Next A

    End With

    '<EhFooter>
    Exit Sub

ResetGroup_Err:
    LogError Err.description & vbCrLf & "in ResetGroup " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub ChangeObtainExp(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo ChangeObtainExp_Err

    '</EhHeader>
    
    With UserList(UserIndex)

        If Groups(.GroupIndex).User(SLOT_LEADER).Index = UserIndex Then
            Groups(.GroupIndex).Acumular = Not Groups(.GroupIndex).Acumular
            
        End If

    End With
    
    '<EhFooter>
    Exit Sub

ChangeObtainExp_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mGroup.ChangeObtainExp " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub AbandonateGroup(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo AbandonateGroup_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        ' ¿Lider disuelve el grupo?
        If Groups(.GroupIndex).User(SLOT_LEADER).Index = UserIndex Then
            ResetGroup .GroupIndex
            WriteConsoleMsg UserIndex, "Has disuelto el grupo.", FontTypeNames.FONTTYPE_INFO
        Else
            ResetMemberGroup .GroupIndex, UserIndex
            WriteConsoleMsg UserIndex, "Has abandonado el grupo.", FontTypeNames.FONTTYPE_INFO

        End If

    End With

    '<EhFooter>
    Exit Sub

AbandonateGroup_Err:
    LogError Err.description & vbCrLf & "in AbandonateGroup " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub UpdatePorcentaje(ByVal GroupIndex As Byte)

    '<EhHeader>
    On Error GoTo UpdatePorcentaje_Err

    '</EhHeader>
    
    Dim A     As Integer

    Dim Value As Byte
    
    With Groups(GroupIndex)
        
        For A = 1 To MAX_MEMBERS_GROUP

            If .User(A).Index > 0 Then
                .User(A).PorcExp = Int(100 / .Members)
                
            End If

        Next A
        
        ' Caso de 3 miembros
        If .Members = 3 Then
            .User(SLOT_LEADER).PorcExp = 34

        End If
        
    End With

    '<EhFooter>
    Exit Sub

UpdatePorcentaje_Err:
    LogError Err.description & vbCrLf & "in UpdatePorcentaje " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub GroupSetPorcentaje(ByVal UserIndex As Integer, _
                              ByVal GroupIndex As Byte, _
                              ByRef Exp() As Byte)

    '<EhHeader>
    On Error GoTo GroupSetPorcentaje_Err

    '</EhHeader>

    Dim A        As Long

    Dim TotalExp As Long, TotalGld As Long

    Dim Valid    As Boolean
        
    Dim MaxCero  As Boolean
        
    Valid = True

    With Groups(GroupIndex)

        If .User(SLOT_LEADER).Index <> UserIndex Then Exit Sub
              
        For A = 1 To MAX_MEMBERS_GROUP

            If .User(A).Index > 0 Then
                TotalExp = TotalExp + Exp(A - 1)
                    
                If .Members > 2 Then
                    If Exp(A - 1) < 10 Then
                        MaxCero = True

                    End If

                End If

            End If

        Next A
            
        If TotalExp <> 100 Then
            WriteConsoleMsg UserIndex, "La suma de los porcentajes debe ser 100.", FontTypeNames.FONTTYPE_INFO

            Exit Sub

        End If
            
        If MaxCero Then
            WriteConsoleMsg UserIndex, "¡Tienes que dar al menos un 10% a cada miembro!", FontTypeNames.FONTTYPE_INFO

            Exit Sub
            
        End If
            
        For A = 1 To MAX_MEMBERS_GROUP

            If .User(A).Index > 0 Then
                .User(A).PorcExp = Exp(A - 1)

            End If

        Next A
        
        SendMessageGroup GroupIndex, UserList(UserIndex).Name, "Ha cambiado los porcentajes de experiencia y oro."

    End With

    '<EhFooter>
    Exit Sub

GroupSetPorcentaje_Err:
    LogError Err.description & vbCrLf & "in GroupSetPorcentaje " & "at line " & Erl

    '</EhFooter>
End Sub

