Attribute VB_Name = "mGroup"
Option Explicit

Public Const MAX_MEMBERS_GROUP As Byte = 5
Public Const MAX_REQUESTS_GROUP As Byte = 10
Private Const MAX_GROUPS As Byte = 100
Private Const SLOT_LEADER As Byte = 1

Private Const EXP_BONUS_MAX_MEMBERS As Single = 1.05 '%
Private Const EXP_BONUS_LEADER_PREMIUM As Single = 1.05 '%
Private Const EXP_BONUS_LEADER_ELV_MAX As Single = 1.1 '%
Private Const EXP_BONUS_LEADER_PENDIENT As Single = 1.5 '%

Private Const PENDIENT_GROUP As Integer = 1322

Public Enum eBonusGroup
    GroupFull = 1
    LeaderPremium = 2
    LeaderPendient = 3
    LeaderMaxLevel = 4
End Enum

Public Type tUserGroup
    Index As Integer
    Exp As Long
    Gld As Long
    PorcExp As Byte
    PorcGld As Byte
End Type

Public Type tGroups
    Members As Byte
    User(1 To MAX_MEMBERS_GROUP) As tUserGroup
    Requests(1 To MAX_REQUESTS_GROUP) As String
End Type

Public Groups(1 To MAX_GROUPS) As tGroups

' Buscamos un SLOT LIBRE para CREAR GRUPO (HASTA 100)
Private Function FreeGroup() As Byte
    Dim A As Long
    
    For A = 1 To MAX_GROUPS
        If Groups(A).User(SLOT_LEADER).Index = 0 Then
            FreeGroup = A
            Exit For
        End If
    Next A
End Function

' Buscamos un SLOT LIBRE para enviar SOLICITUD AL GRUPO. (HASTA 10)
Private Function FreeGroupRequest(ByVal GroupIndex As Byte) As Byte
    Dim A As Long
    
    For A = 1 To MAX_REQUESTS_GROUP
        If Groups(GroupIndex).Requests(A) = vbNullString Then
            FreeGroupRequest = A
            Exit For
        End If
    Next A
End Function

' Buscamos UNA SOLICITUD en el GRUPO.
Private Function SearchGroupRequest(ByVal GroupIndex As Byte, ByVal UserName As String) As Byte
    Dim A As Long
    
    For A = 1 To MAX_REQUESTS_GROUP
        If Groups(GroupIndex).Requests(A) = UserName Then
            SearchGroupRequest = A
            Exit For
        End If
    Next A
End Function

' Buscamos un SLOT LIBRE para meter un NUEVO MIEMBRO.
Private Function FreeGroupMember(ByVal GroupIndex As Byte) As Byte
    Dim A As Long
    
    For A = 1 To MAX_MEMBERS_GROUP
        If Groups(GroupIndex).User(A).Index = 0 Then
            FreeGroupMember = A
            Exit For
        End If
    Next A
End Function

Private Sub SetGroupMember(ByVal GroupIndex As Integer, _
                            ByVal SlotMember As Byte, _
                            ByVal UserIndex As Integer, _
                            ByVal PorcExp As Byte, _
                            ByVal PorcGld As Byte)
    With Groups(GroupIndex)
        .User(SlotMember).Index = UserIndex
        .User(SlotMember).PorcExp = PorcExp
        .User(SlotMember).PorcGld = PorcGld
    End With
End Sub

' Creamos un NUEVO GRUPO.
Public Sub CreateGroup(ByVal UserIndex As Integer)
    Dim Slot As Byte
    
    With UserList(UserIndex)
        
        If .GroupIndex > 0 Then
            WriteConsoleMsg UserIndex, "Ya perteneces a una party.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
        Slot = FreeGroup
        
        If Slot > 0 Then
            SetGroupMember Slot, SLOT_LEADER, UserIndex, 100, 100
            Groups(Slot).Members = SLOT_LEADER
            UserList(UserIndex).GroupIndex = Slot
            UserList(UserIndex).GroupSlotUser = SLOT_LEADER
            
            WriteConsoleMsg UserIndex, "Has creado un nuevo grupo. Podrás ingresar al Panel con la tecla F7.", FontTypeNames.FONTTYPE_INFO
            WriteGroupPrincipal UserIndex
        Else
            WriteConsoleMsg UserIndex, "Existen 100 grupos creados. Por favor espera que se disuelva alguno.", FontTypeNames.FONTTYPE_INFO
        End If
    
    End With
End Sub

' Enviamos solicitud a UN GRUPO
Public Sub SendInvitationGroup(ByVal UserIndex As Integer)

    ' Un personaje decide solicitar entrar a una party.
    Dim tUser As Integer
    Dim Slot As Byte
    
    With UserList(UserIndex)
    
        If .GroupIndex > 0 Then
            WriteConsoleMsg UserIndex, "Ya perteneces a una party.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
        tUser = .flags.TargetUser
        
        If tUser > 0 Then
            With UserList(tUser)
                If .GroupIndex > 0 Then
                
                    If SearchGroupRequest(.GroupIndex, UCase$(UserList(UserIndex).Name)) > 0 Then Exit Sub
                    
                    If Groups(.GroupIndex).Members = MAX_MEMBERS_GROUP Then
                        WriteConsoleMsg UserIndex, "La party está llena.", FontTypeNames.FONTTYPE_INFO
                        Exit Sub
                    End If
                    
                    Slot = FreeGroupRequest(.GroupIndex)
                    
                    If Slot > 0 Then
                        Groups(.GroupIndex).Requests(Slot) = UCase$(UserList(UserIndex).Name)
                        WriteConsoleMsg UserIndex, "Has enviado solicitud al grupo de " & .Name & " . Espera pronta noticias", FontTypeNames.FONTTYPE_INFO
                        WriteConsoleMsg tUser, "Has recibido una solicitud de Party del personaje " & UserList(UserIndex).Name & ". Para aceptar ingresa al panel con la tecla F7.", FontTypeNames.FONTTYPE_INFO
                    Else
                        WriteConsoleMsg UserIndex, "La party tiene muchas solicitudes sin responder. Pídele al lider que descarte algunas para que puedas solicitar entrar.", FontTypeNames.FONTTYPE_INFO
                    End If
                Else
                    WriteConsoleMsg UserIndex, "El personaje no pertenece a ninguna party", FontTypeNames.FONTTYPE_INFO
                End If
            End With
        
        
        End If
        
    
    End With
End Sub
' El LIDER acepta la SOLICITUD RECIBIDA.
Public Sub AcceptInvitationGroup(ByVal UserIndex As Integer, ByVal UserName As String)

    Dim Slot As Byte
    Dim tUser As Integer
    Dim SlotRequest As Byte
    
    With UserList(UserIndex)
        If .GroupIndex = 0 Then Exit Sub
        If Groups(.GroupIndex).User(SLOT_LEADER).Index <> UserIndex Then Exit Sub
        SlotRequest = SearchGroupRequest(.GroupIndex, UCase$(UserName))
        If SlotRequest <= 0 Then Exit Sub
        
        
        tUser = NameIndex(UserName)
        
        If tUser <= 0 Then
            WriteConsoleMsg UserIndex, "El personaje está offline.", FontTypeNames.FONTTYPE_INFO
            WriteGroupRequests UserIndex
            Exit Sub
        End If
        
        If UserList(tUser).GroupIndex > 0 Then
            WriteConsoleMsg UserIndex, "El personaje ya está en otro Grupo.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
        Slot = FreeGroupMember(.GroupIndex)
        
        If Slot > 0 Then
            Groups(.GroupIndex).Requests(SlotRequest) = vbNullString
            Groups(.GroupIndex).User(Slot).Index = tUser
            Groups(.GroupIndex).Members = Groups(.GroupIndex).Members + 1
            
            UserList(tUser).GroupSlotUser = Slot
            UserList(tUser).GroupIndex = .GroupIndex
            SendMessageGroup .GroupIndex, .Name, "El personaje " & UserName & " ha sido aceptado en el grupo."
            
            UpdatePorcentaje .GroupIndex
            WriteGroupPrincipal UserIndex
        Else
            WriteConsoleMsg UserIndex, "La party está llena.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
    End With
End Sub

' El LIDER RECHAZA la SOLICITUD RECIBIDA.
Public Sub RechaceInvitationGroup(ByVal UserIndex As Integer, ByVal UserName As String)

    Dim Slot As Byte
    Dim tUser As Integer
    
    With UserList(UserIndex)
        If .GroupIndex = 0 Then Exit Sub
        If Groups(.GroupIndex).User(SLOT_LEADER).Index <> UserIndex Then Exit Sub
        
        
        tUser = NameIndex(UserName)
        Slot = SearchGroupRequest(.GroupIndex, UCase$(UserName))
        
        If Slot <= 0 Then Exit Sub
        
        If tUser <= 0 Then
            WriteConsoleMsg UserIndex, "El personaje está offline.", FontTypeNames.FONTTYPE_INFO
            Groups(.GroupIndex).Requests(Slot) = vbNullString
            WriteGroupRequests UserIndex
            Exit Sub
        End If

        Groups(.GroupIndex).Requests(Slot) = vbNullString
        SendMessageGroup .GroupIndex, .Name, "El personaje " & UserName & " ha sido rechazado para entrar al grupo."
            
        WriteGroupRequests UserIndex
    End With
End Sub

Private Function CheckGroupMap(ByVal GroupIndex As Byte) As Boolean
    Dim A As Byte
    
    CheckGroupMap = True
    
    For A = 1 To MAX_MEMBERS_GROUP
        With Groups(GroupIndex)
            If .User(A).Index > 0 Then
                If UserList(.User(A).Index).Pos.map <> UserList(.User(SLOT_LEADER).Index).Pos.map Then
                    CheckGroupMap = False
                    Exit For
                End If
            End If
        End With
    Next A
End Function
' El grupo acumula experiencia
Public Sub AddExpGroup(ByVal GroupIndex As Byte, ByRef Exp As Long)
    
    Dim A As Long
    Dim TempExp As Long
    Dim ExpTemp As Long
    
    With Groups(GroupIndex)
        If Not CheckGroupMap(GroupIndex) Then Exit Sub
        
        ExpTemp = Exp
        
        If .Members <> 1 Then
            ' Bonus al tener MÁXIMO DE MIEMBROS.
            If .Members = MAX_MEMBERS_GROUP Then
                ExpTemp = ExpTemp * EXP_BONUS_MAX_MEMBERS
            End If
                
            ' Bonus al tener lider premium
            If UserList(.User(SLOT_LEADER).Index).flags.Premium Then
                ExpTemp = ExpTemp * EXP_BONUS_LEADER_PREMIUM
            End If
            
            ' Bonus al tener lider máximo
            If UserList(.User(SLOT_LEADER).Index).Stats.ELV Then
                ExpTemp = ExpTemp * EXP_BONUS_LEADER_ELV_MAX
            End If
            
            If TieneObjetos(PENDIENT_GROUP, 1, .User(SLOT_LEADER).Index) Then
                ExpTemp = ExpTemp * EXP_BONUS_LEADER_PENDIENT
            End If
        End If
                
        For A = 1 To MAX_MEMBERS_GROUP
            If .User(A).Index > 0 Then
                If .User(A).Exp > MAXEXP Then
                    SaveExpAndGldMember GroupIndex, .User(A).Index
                End If
                
                TempExp = Porcentaje(ExpTemp, .User(A).PorcExp)
                .User(A).Exp = .User(A).Exp + TempExp
                
                WriteConsoleMsg .User(A).Index, "Experiencia acumulada en Party» +" & TempExp, FontTypeNames.FONTTYPE_PARTY
                
            End If
            
        Next A
    End With
End Sub

' El grupo acumula Oro
Public Sub AddGldGroup(ByVal GroupIndex As Byte, ByVal Gld As Long)
    
    Dim A As Long
    
    With Groups(GroupIndex)
        If Not CheckGroupMap(GroupIndex) Then Exit Sub
        
        For A = 1 To MAX_MEMBERS_GROUP
            If .User(A).Index > 0 Then
                .User(A).Gld = .User(A).Gld + Porcentaje(Gld, .User(A).PorcGld)
                
                
                WriteConsoleMsg .User(A).Index, "Oro acumulado en grupo» +" & Porcentaje(Gld, .User(A).PorcGld), FontTypeNames.FONTTYPE_PARTY
            End If
        Next A
        
        
    End With
End Sub

' Distribuimos las experiencias de las partys
Public Sub DistributeExpAndGldGroups()
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
End Sub

' Actualizamos Experiencia y Oro del personaje.
Public Sub SaveExpAndGldMember(ByVal GroupIndex As Byte, ByVal UserIndex As Integer)
    
    Dim SlotUser As Byte
    
    With UserList(UserIndex)
        SlotUser = .GroupSlotUser

        .Stats.Exp = .Stats.Exp + Groups(GroupIndex).User(SlotUser).Exp
        .Stats.Gld = .Stats.Gld + Groups(GroupIndex).User(SlotUser).Gld
        
         If .Stats.Exp > MAXEXP Then _
            .Stats.Exp = MAXEXP
        
        CheckUserLevel UserIndex
        WriteUpdateGold UserIndex
        
        WriteConsoleMsg UserIndex, "Hemos actualizado tu experiencia y Oro. Has conseguido " & Groups(GroupIndex).User(SlotUser).Exp & " puntos de experiencia y " & Groups(GroupIndex).User(SlotUser).Gld & " monedas de oro", FontTypeNames.FONTTYPE_ORO
        Groups(GroupIndex).User(SlotUser).Exp = 0
        Groups(GroupIndex).User(SlotUser).Gld = 0
    End With
End Sub
' Enviamos un mensaje al grupo.
Public Sub SendMessageGroup(ByVal GroupIndex As Byte, ByVal Emisor As String, ByVal message As String)
    Dim A As Long

    For A = 1 To MAX_MEMBERS_GROUP
        With Groups(GroupIndex)
            If .User(A).Index > 0 Then
                If Emisor <> vbNullString Then
                    WriteConsoleMsg .User(A).Index, Emisor & "» " & message, FontTypeNames.FONTTYPE_PARTY
                Else
                    WriteConsoleMsg .User(A).Index, message, FontTypeNames.FONTTYPE_PARTY
                End If
            End If
        End With
    Next A
End Sub

' Reiniciamos la información de un miembro del Grupo
Private Sub ResetMemberGroup(ByVal GroupIndex As Byte, ByVal UserIndex As Integer)

    With Groups(GroupIndex)
        
        
        ' Asignamos Experiencia y Oro obtenido hasta el momento
        mGroup.SaveExpAndGldMember GroupIndex, UserIndex
        
        .Members = .Members - 1
        .User(UserList(UserIndex).GroupSlotUser).Index = 0
        .User(UserList(UserIndex).GroupSlotUser).Exp = 0
        .User(UserList(UserIndex).GroupSlotUser).Gld = 0
        .User(UserList(UserIndex).GroupSlotUser).PorcExp = 0
        .User(UserList(UserIndex).GroupSlotUser).PorcGld = 0
        
    End With
    
    With UserList(UserIndex)
        .GroupIndex = 0
        .GroupRequired = 0
        .GroupSlotUser = 0
    End With
    
    UpdatePorcentaje GroupIndex
End Sub

' Reiniciamos la información del grupo
Private Sub ResetGroup(ByVal GroupIndex As Byte)

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
End Sub

Public Sub AbandonateGroup(ByVal UserIndex As Integer)
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
End Sub

Public Function CheckBonusGroup(ByVal GroupIndex As Integer, ByVal Bonus As eBonusGroup) As Boolean

    With Groups(GroupIndex)
        Select Case Bonus
            Case eBonusGroup.GroupFull
                If .Members = MAX_MEMBERS_GROUP Then
                    CheckBonusGroup = True
                    Exit Function
                End If
            Case eBonusGroup.LeaderPremium
                If UserList(.User(SLOT_LEADER).Index).flags.Premium Then
                    CheckBonusGroup = True
                    Exit Function
                End If
            Case eBonusGroup.LeaderPendient
                If TieneObjetos(PENDIENT_GROUP, 1, .User(SLOT_LEADER).Index) Then
                    CheckBonusGroup = True
                    Exit Function
                End If
                
            Case eBonusGroup.LeaderMaxLevel
                If UserList(.User(SLOT_LEADER).Index).Stats.ELV >= STAT_MAXELV Then
                    CheckBonusGroup = True
                    Exit Function
                End If
        End Select
    End With
End Function


Private Sub UpdatePorcentaje(ByVal GroupIndex As Byte)
    
    Dim A As Integer
    Dim Value As Byte
    
    With Groups(GroupIndex)
        
        For A = 1 To MAX_MEMBERS_GROUP
            If .User(A).Index > 0 Then
                .User(A).PorcExp = Int(100 / .Members)
                .User(A).PorcGld = Int(100 / .Members)
                
            End If
        Next A
        
        ' Caso de 3 miembros
        If .Members = 3 Then
            .User(SLOT_LEADER).PorcExp = 34
            .User(SLOT_LEADER).PorcGld = 34
        End If
        
    End With
End Sub

Public Sub GroupSetPorcentaje(ByVal UserIndex As Integer, ByVal GroupIndex As Byte, ByRef Exp() As Byte, ByRef Gld() As Byte)
    Dim A As Long
    Dim TotalExp As Long, TotalGld As Long
    Dim Valid As Boolean
    
    
    Valid = True
    With Groups(GroupIndex)
        If .User(SLOT_LEADER).Index <> UserIndex Then Exit Sub
        
        For A = 1 To MAX_MEMBERS_GROUP
            If .User(A).Index > 0 Then
                If Exp(A - 1) < 10 Then
                    Valid = False
                End If
                
                If Gld(A - 1) < 10 Then
                    Valid = False
                End If
                
                TotalExp = TotalExp + Exp(A - 1)
                TotalGld = TotalGld + Gld(A - 1)
            End If
        Next A
        
        
        If TotalExp <> 100 Or TotalGld <> 100 Or Valid = False Then
            WriteConsoleMsg UserIndex, "La suma de los porcentajes debe ser 100 y el miembro debe tener 10 como mínimo.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
        
        For A = 1 To MAX_MEMBERS_GROUP
            If .User(A).Index > 0 Then
                .User(A).PorcExp = Exp(A - 1)
                .User(A).PorcGld = Gld(A - 1)
            End If
        Next A
        
        SendMessageGroup GroupIndex, UserList(UserIndex).Name, "Ha cambiado los porcentajes de experiencia y oro."
    End With
End Sub
