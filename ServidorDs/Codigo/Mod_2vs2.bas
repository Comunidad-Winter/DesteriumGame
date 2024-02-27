Attribute VB_Name = "Mod_2vs2"
Option Explicit

Public Const RETO_COLOR As String = "~65~190~156~0~0"

Public Type ruleStruct

    drop_inv   As Boolean
    gold_gamble As Long

End Type

Public Type teamStruct

    user_Index(1) As Integer
    round_count As Byte
    return_city As Byte

End Type

Public Type retoStruct

    team_array(1) As teamStruct
    general_rules As ruleStruct
    count_Down As Byte
    used_ring  As Boolean

    nextRoundCount As Integer

End Type

Public Type userStruct

    tempStruct As retoStruct
    accept_count As Byte
    reto_Index As Integer
    nick_sender As String
    reto_used  As Boolean
    return_city As Byte
    acceptedOK As Boolean
    acceptLimit As Integer

End Type

Private Type tempPos

    X          As Integer
    Y          As Integer

End Type

Public reto_2Map As Integer
Public reto_List() As retoStruct
Public reto_RingPos() As tempPos

Public Sub initRetoData2()

'
' @ amishar

    Dim bRead  As New clsIniManager

    Dim nRing  As Integer

    Set bRead = New clsIniManager

    Call bRead.Initialize(App.Path & "\Reto2vs2.ini")

    nRing = val(bRead.GetValue("INIT", "Arenas"))

    If (nRing = 0) Then Exit Sub

    ReDim reto_List(0 To nRing - 1) As retoStruct
    ReDim reto_RingPos(1 To nRing, 1 To 2, 1 To 2) As tempPos

    reto_2Map = val(bRead.GetValue("INIT", "MapaArenas"))

    Dim i      As Long
    Dim j      As Long
    Dim p      As Long
    Dim S      As String

    For i = 1 To nRing
        For j = 1 To 2
            For p = 1 To 2
                S = bRead.GetValue("ARENA" & CStr(i), "Equipo" & CStr(j) & "Jugador" & CStr(p))

                reto_RingPos(i, j, p).X = val(ReadField(1, S, Asc("-")))
                reto_RingPos(i, j, p).Y = val(ReadField(2, S, Asc("-")))

            Next p
        Next j
    Next i

    Set bRead = Nothing

End Sub

Public Sub loop_reto()

'
' @ amishar

    Dim loopC  As Long

    For loopC = 0 To UBound(reto_List())

        If (reto_List(loopC).used_ring) Then
            Call loop_reto_index(loopC)
        End If

    Next loopC

End Sub

Private Function check_player_List(ByVal Userindex As Integer) As Boolean

'
' @ amishar

    With UserList(Userindex).reto2Data

        Dim tmp(2) As Integer

        With .tempStruct

            check_player_List = False

            tmp(0) = .team_array(0).user_Index(1)
            tmp(1) = .team_array(1).user_Index(0)
            tmp(2) = .team_array(1).user_Index(1)

            If Userindex = tmp(0) Or Userindex = tmp(1) Or Userindex = tmp(2) Then Exit Function

            If tmp(0) = tmp(1) Or tmp(0) = tmp(2) Then Exit Function

            If tmp(1) = tmp(2) Then Exit Function

            check_player_List = True
        End With
    End With

End Function

Public Function can_Attack(ByVal AttackerIndex As Integer, _
                           ByVal victimIndex As Integer) As Boolean

'
' @ amishar

    Dim RetoIndex As Integer
    Dim teamIndex As Integer
    Dim tempIndex As Integer
    Dim teamLoop As Long

    can_Attack = True

    RetoIndex = UserList(AttackerIndex).reto2Data.reto_Index

    teamIndex = -1

    If reto_List(RetoIndex).used_ring Then

        For teamLoop = 0 To 1

            If reto_List(RetoIndex).team_array(teamLoop).user_Index(0) = AttackerIndex Or reto_List(RetoIndex).team_array(teamLoop).user_Index(1) = AttackerIndex Then
                teamIndex = teamLoop

                Exit For

            End If

        Next teamLoop

        If teamIndex <> -1 Then
            tempIndex = IIf(reto_List(RetoIndex).team_array(teamIndex).user_Index(0) = AttackerIndex, 1, 0)

            If reto_List(RetoIndex).team_array(teamIndex).user_Index(tempIndex) = victimIndex Then
                can_Attack = False
            End If
        End If
    End If

End Function


Private Sub loop_reto_index(ByVal reto_Index As Integer)

'
' @ amishar

    Dim i      As Long
    Dim j      As Long
    Dim h      As Integer
    Dim m      As String

    With reto_List(reto_Index)

        If (.nextRoundCount <> 0) Then
            .nextRoundCount = .nextRoundCount - 1

            If (.nextRoundCount = 0) Then
                Call warp_Teams(reto_Index, True)

                .count_Down = 15
            End If
        End If

        If (.count_Down <> 0) Then
            .count_Down = (.count_Down - 1)

            If (.count_Down > 0) Then
                m = "Retos>" & CStr(.count_Down)
            End If

            For i = 0 To 1
                For j = 0 To 1
                    h = .team_array(i).user_Index(j)

                    If (h <> 0) Then
                        If UserList(h).ConnID <> -1 Then

                            If (.count_Down > 0) Then
                                Call Protocol.WriteConsoleMsg(h, m, FontTypeNames.FONTTYPE_RETO)
                            Else
                                WriteConsoleMsg h, "Retos> YA!", FontTypeNames.FONTTYPE_FIGHT
                            End If

                            If (.count_Down = 0) Then Call Protocol.WritePauseToggle(h)
                        End If
                    End If

                Next j
            Next i

        End If

    End With

End Sub

Public Function get_reto_index() As Integer

'
' @ amishar

    Dim loopC  As Long

    For loopC = 0 To UBound(reto_List())

        If (reto_List(loopC).used_ring = False) Then
            get_reto_index = CInt(loopC)

            Exit Function

        End If

    Next loopC

    get_reto_index = -1

End Function

Public Sub set_reto_struct(ByVal user_Index As Integer, _
                           ByVal my_team As String, _
                           ByRef enemy_name As String, _
                           ByRef team_enemy As String, _
                           ByVal invDrop As Boolean, _
                           ByVal goldAmount As Long)

'
' @ amishar

    With UserList(user_Index).reto2Data
        .accept_count = 0

        With .tempStruct
            .count_Down = 0
            .used_ring = False

            With .team_array(0)
                .user_Index(0) = user_Index
                .user_Index(1) = NameIndex(my_team)
            End With

            With .team_array(1)
                .user_Index(0) = NameIndex(enemy_name)
                .user_Index(1) = NameIndex(team_enemy)
            End With

            With .general_rules
                .drop_inv = invDrop
                .gold_gamble = goldAmount
            End With

        End With

    End With

End Sub

Public Sub user_retoLoop(ByVal user_Index As Integer)

'
' @ amishar

    With UserList(user_Index).reto2Data

        If (.acceptLimit <> 0) Then
            .acceptLimit = .acceptLimit - 1

            If (.acceptLimit <= 0) Then
                Call message_reto(.tempStruct, "El reto se ha autocancelado debido a que el tiempo para aceptar ha llegado a su límite.")

                Dim j As Long
                Dim i As Long
                Dim N As Integer
                Dim b As userStruct

                For j = 0 To 1
                    For i = 0 To 1
                        N = .tempStruct.team_array(j).user_Index(i)

                        If N > 0 Then
                            If UCase$(UserList(N).reto2Data.nick_sender) = UCase$(UserList(user_Index).Name) Then
                                UserList(N).reto2Data.nick_sender = vbNullString
                                UserList(N).reto2Data.acceptedOK = False
                            End If
                        End If

                    Next i
                Next j

                UserList(user_Index).reto2Data = b
            End If
        End If

        If (.return_city <> 0) Then
            .return_city = .return_city - 1

            If (.return_city = 0) Then

                Dim p As WorldPos

                p = Arghal

                Call FindLegalPos(user_Index, p.Map, p.X, p.Y)
                Call WarpUserChar(user_Index, p.Map, p.X, p.Y, True)

                'Call Protocol.WriteConsoleMsg(user_Index, "Regresas a la ciudad." & RETO_COLOR, FontTypeNames.FONTTYPE_GUILD)
            End If

        End If

    End With

End Sub

Public Sub erase_userData(ByVal user_Index As Integer)

'
' @ amishar

    With UserList(user_Index).reto2Data

        Dim dumpStruct As retoStruct

        .accept_count = 0
        .nick_sender = vbNullString
        .reto_Index = 0
        .reto_used = False
        .tempStruct = dumpStruct

    End With

End Sub

Public Function can_send_reto(ByVal user_Index As Integer, _
                              ByRef fERROR As String) As Boolean

'
' @ amishar

    can_send_reto = False

    With UserList(user_Index)

        If UserList(user_Index).flags.Comerciando Then
            fERROR = "¡Estás Comerciando!"
            Exit Function
        End If

        If (.flags.Muerto <> 0) Then
            fERROR = "¡Estás muerto!"
            Exit Function
        End If

        If (.Counters.Pena <> 0) Then
            fERROR = "Estás en la cárcel"
            Exit Function
        End If

        If (.reto2Data.reto_Index <> 0) Or (reto_List(.reto2Data.reto_Index).used_ring) Then
            fERROR = "Ya estás en reto"
            Exit Function
        End If

        If (.Stats.GLD < .reto2Data.tempStruct.general_rules.gold_gamble) Then
            Call Protocol.WriteConsoleMsg(user_Index, "No tienes el oro necesario.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If


        If (.reto2Data.tempStruct.general_rules.gold_gamble < 25000) Then
            Call Protocol.WriteConsoleMsg(user_Index, "El mínimo de oro paraa retar es de 25.000 monedas de oro.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If

        If (.reto2Data.tempStruct.general_rules.gold_gamble > 2000000) Then
            Call Protocol.WriteConsoleMsg(user_Index, "El máximo de oro para apostar en el reto son 2.000.000 monedas de oro.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If

        If (.Stats.GLD < 25000) Then
            Call Protocol.WriteConsoleMsg(user_Index, "El mínimo de oro para retar es de 25.000 monedas de oro.", FontTypeNames.FONTTYPE_INFO)
            Exit Function

        End If

        If (.Stats.ELV < 35) Then
            Call Protocol.WriteConsoleMsg(user_Index, "Tu nivel es insuficiente para entrar a este desafío. El nivel mínimo es 35.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If

        With .reto2Data.tempStruct
            can_send_reto = check_User(.team_array(0).user_Index(1), fERROR)

            If (can_send_reto) Then
                can_send_reto = check_User(.team_array(1).user_Index(0), fERROR)
            Else

                Exit Function

            End If

            If (can_send_reto) Then
                can_send_reto = check_User(.team_array(1).user_Index(1), fERROR)
            Else

                Exit Function

            End If

            If (can_send_reto) Then
                can_send_reto = check_player_List(user_Index)

                If Not can_send_reto Then fERROR = "No puedes repetir el nombre de un usuario!"
            Else

                Exit Function

            End If

        End With
    End With

End Function

Private Function check_User(ByVal user_Index As Integer, _
                            ByRef fERROR As String) As Boolean

'
' @ amishar

    check_User = False

    If (user_Index = 0) Then
        fERROR = "No se ha enviado la solicitud del reto debido a que uno de los usuarios se encuentra desconectado."

        Exit Function

    End If

    With UserList(user_Index)

        If (.flags.Muerto <> 0) Then
            fERROR = .Name & " está muerto"

            Exit Function

        End If

        If (.Counters.Pena <> 0) Then
            fERROR = .Name & " está en la cárcel!"

            Exit Function

        End If

        If (.reto2Data.reto_Index <> 0) Then
            fERROR = .Name & " ya está en reto!"

            Exit Function

        End If

        If (.Stats.GLD < .reto2Data.tempStruct.general_rules.gold_gamble) Then
            fERROR = .Name & " no tiene el oro necesario!"

            Exit Function

        End If

        If (.Pos.Map <> 1) Then
            fERROR = .Name & " debe estar en la Ciudad de Ullathorpe para retar."
            Exit Function
        End If

        If (MapInfo(.Pos.Map).Pk = True) Then
            fERROR = .Name & " está en zona insegura."

            Exit Function

        End If

        If (.Stats.ELV < 35) Then
            fERROR = .Name & " debe ser mayor a nivel 35!"

            Exit Function

        End If

        check_User = True
    End With

End Function

Public Sub send_Reto(ByVal user_Index As Integer)

'
' @ amishar

    With UserList(user_Index).reto2Data

        Dim i  As Long
        Dim j  As Long

        Dim team_str As String
        Dim gamble_str As String

        team_str = UserList(.tempStruct.team_array(0).user_Index(0)).Name & "(" & UserList(.tempStruct.team_array(0).user_Index(0)).Stats.ELV & ") y " & UserList(.tempStruct.team_array(0).user_Index(1)).Name & "(" & UserList(.tempStruct.team_array(0).user_Index(1)).Stats.ELV & ") vs " & UserList(.tempStruct.team_array(1).user_Index(0)).Name & "(" & UserList(.tempStruct.team_array(1).user_Index(0)).Stats.ELV & ") y " & UserList(.tempStruct.team_array(1).user_Index(1)).Name & "(" & UserList(.tempStruct.team_array(1).user_Index(1)).Stats.ELV & ")"

        gamble_str = ". Apuesta por " & Format$(.tempStruct.general_rules.gold_gamble, "####") & " monedas de oro"

        If (.tempStruct.general_rules.drop_inv) Then
            gamble_str = " y los items"
        End If

        For i = 0 To 1
            For j = 0 To 1
                UserList(.tempStruct.team_array(i).user_Index(j)).reto2Data.nick_sender = UCase$(UserList(user_Index).Name)

                If (.tempStruct.team_array(i).user_Index(j) <> user_Index) Then
                    Call Protocol.WriteConsoleMsg(.tempStruct.team_array(i).user_Index(j), UserList(user_Index).Name & " te invita a participar en el reto  " & team_str & " " & gamble_str & " . Para aceptar escribe /ACEPTAR " & UCase$(UserList(user_Index).Name) & ".", FontTypeNames.FONTTYPE_GUILD)
                End If

            Next j
        Next i

        Call Protocol.WriteConsoleMsg(user_Index, "Tipea /ACEPTAR.", FontTypeNames.FONTTYPE_INFO)
        Call Protocol.WriteConsoleMsg(user_Index, "Se han enviado las solicitudes.", FontTypeNames.FONTTYPE_INFO)

        .acceptLimit = 60
    End With

End Sub

Public Sub disconnect_Reto(ByVal user_Index As Integer)

'
' @ amishar

    Dim team_Index As Integer
    Dim team_winner As Byte
    Dim reto_Index As Integer

    reto_Index = UserList(user_Index).reto2Data.reto_Index

    team_Index = find_Team(user_Index, reto_Index)

    If (team_Index <> -1) Then
        team_winner = IIf(team_Index = 1, 0, 1)
        Call finish_reto(UserList(user_Index).reto2Data.reto_Index, team_winner, True)
    End If

End Sub

Public Sub accept_Reto(ByVal user_Index As Integer, ByVal requestName As String)

'
' @ amishar

    Dim SendIndex As Integer


    If UserList(user_Index).flags.Comerciando Then
        Call WriteConsoleMsg(user_Index, "¡Estás Comerciando!", FontTypeNames.FONTTYPE_TALK)

        Exit Sub

    End If

    If Not UserList(user_Index).Pos.Map = 1 Then
        Call Protocol.WriteConsoleMsg(user_Index, "Debes estar en la Ciudad de Ullathorpe para aceptar el reto..", FontTypeNames.FONTTYPE_INFO)

        Exit Sub

    End If

    SendIndex = NameIndex(requestName)
    If UserList(user_Index).reto2Data.acceptedOK = False Then
        If (UserList(user_Index).Stats.GLD < UserList(user_Index).reto2Data.tempStruct.general_rules.gold_gamble) Then
            Call Protocol.WriteConsoleMsg(user_Index, "No tenes suficientes monedas de oro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

    If (UserList(user_Index).Stats.GLD < UserList(SendIndex).reto2Data.tempStruct.general_rules.gold_gamble) Then
        Call Protocol.WriteConsoleMsg(user_Index, "No tenes suficientes monedas de oro.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

        If (SendIndex = 0) Or (UCase$(requestName) <> UserList(user_Index).reto2Data.nick_sender) Then
            Call Protocol.WriteConsoleMsg(user_Index, requestName & " no te invito a ningún reto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If

        If UserList(SendIndex).Name = UserList(user_Index).reto2Data.nick_sender Then
            WriteConsoleMsg user_Index, "Tú no puedes aceptar tu misma solicitud!", FontTypeNames.FONTTYPE_INFO

            Exit Sub

        End If
    End If

    If (SendIndex = 0) Then Exit Sub

    If UserList(user_Index).reto2Data.acceptedOK Then
        Call Protocol.WriteConsoleMsg(user_Index, "Tú ya has aceptado, espera a que otros acepten.", FontTypeNames.FONTTYPE_INFO)

        Exit Sub

    End If

    UserList(SendIndex).reto2Data.accept_count = (UserList(SendIndex).reto2Data.accept_count + 1)

    Call message_reto(UserList(SendIndex).reto2Data.tempStruct, UserList(user_Index).Name & " aceptó el reto.")

    UserList(user_Index).flags.BeforeMap = UserList(user_Index).Pos.Map
    UserList(user_Index).flags.BeforeX = UserList(user_Index).Pos.X
    UserList(user_Index).flags.BeforeY = UserList(user_Index).Pos.Y

    Dim BoludoSinoro As Integer
    
    If CheckeoCompleto(SendIndex, BoludoSinoro) = False Then
        Call message_reto(UserList(SendIndex).reto2Data.tempStruct, "El usuario " & UserList(BoludoSinoro).Name & " no tiene el oro suficiente.")
        Exit Sub
    End If
    
    If (UserList(SendIndex).reto2Data.accept_count = 4) Then
        Call init_reto(SendIndex)
    End If

    UserList(user_Index).reto2Data.acceptedOK = True

End Sub

Private Function CheckeoCompleto(ByVal SendIndex As Integer, ByRef BoludoSinoro As Integer) As Boolean

    On Error Resume Next
    
    Dim i As Long, T As Long, u As Integer
    
    CheckeoCompleto = False
    
    With UserList(SendIndex).reto2Data
       
        For T = 0 To 1
        
            For i = 0 To 1
            
            u = .tempStruct.team_array(T).user_Index(i)
            
            If (u <> 0) Then
                If UserList(u).Stats.GLD < UserList(SendIndex).reto2Data.tempStruct.general_rules.gold_gamble Then
                    BoludoSinoro = u
                    Exit Function
                End If
            End If
            
            Next i
            
        Next T
        
    End With

    CheckeoCompleto = True
    
End Function

Private Sub init_reto(ByVal userSendIndex As Integer)
 
        '
        ' @ amishar
   
        Dim reto_Index As Integer
   
        reto_Index = get_reto_index()
   
        If (reto_Index = -1) Then
                Call message_reto(UserList(userSendIndex).reto2Data.tempStruct, "Reto cancelado, todas las arenas están ocupadas.")
 
                Exit Sub
 
        End If
   
        UserList(userSendIndex).reto2Data.acceptLimit = 0
        reto_List(reto_Index) = UserList(userSendIndex).reto2Data.tempStruct
        reto_List(reto_Index).used_ring = True
        reto_List(reto_Index).count_Down = 15
   
        Call warp_Teams(reto_Index)
   
End Sub
 
Private Sub warp_Teams(ByVal reto_Index As Integer, _
                       Optional ByVal respawnUser As Boolean = False)

        '
        ' @ amishar
    
        With reto_List(reto_Index)

                Dim loopC As Long
                Dim mPosX As Byte
                Dim mPosY As Byte
                Dim NuSer As Integer
         
                .count_Down = 15
         
                For loopC = 0 To 1
                        NuSer = .team_array(0).user_Index(loopC)
              
                        If (NuSer <> 0) Then
                                If (UserList(NuSer).ConnID <> -1) Then
                                        mPosX = get_pos_x(reto_Index + 1, 1, loopC + 1)
                                        mPosY = get_pos_y(reto_Index + 1, 1, loopC + 1)
                     
                                        UserList(NuSer).reto2Data.reto_used = True
                                     
                                        Call WarpUserChar(NuSer, reto_2Map, mPosX, mPosY, True)
                                        Call Protocol.WritePauseToggle(NuSer)
                                          UserList(NuSer).flags.MapU = UserList(NuSer).Pos.Map
                                        UserList(NuSer).flags.MapX = UserList(NuSer).Pos.X
                                        UserList(NuSer).flags.MapY = UserList(NuSer).Pos.Y
                                        If (respawnUser) Then
                                                If (UserList(NuSer).flags.Muerto) Then
                                                        Call RevivirUsuario(NuSer)
                                                End If
                         
                                                UserList(NuSer).Stats.MinHp = UserList(NuSer).Stats.MaxHp
                                                UserList(NuSer).Stats.MinMAN = UserList(NuSer).Stats.MaxMAN
                                                UserList(NuSer).Stats.MinHam = 100
                                                UserList(NuSer).Stats.MinAGU = 100
                                                UserList(NuSer).Stats.MinSta = UserList(NuSer).Stats.MaxSta
                         
                                                Call Protocol.WriteUpdateUserStats(NuSer)
                                        End If

                                Else
                                        UserList(NuSer).reto2Data.acceptedOK = False
                                End If
                        End If

                Next loopC
         
                For loopC = 0 To 1
                        NuSer = .team_array(1).user_Index(loopC)
              
                        If (NuSer <> 0) Then
                                If (UserList(NuSer).ConnID <> -1) Then
                                        mPosX = get_pos_x(reto_Index + 1, 2, loopC + 1)
                                        mPosY = get_pos_y(reto_Index + 1, 2, loopC + 1)
                    
                                        UserList(NuSer).reto2Data.reto_used = True

                                        Call WarpUserChar(NuSer, reto_2Map, mPosX, mPosY, True)
                                        Call Protocol.WritePauseToggle(NuSer)
                                        UserList(NuSer).flags.MapU = UserList(NuSer).Pos.Map
                                        UserList(NuSer).flags.MapX = UserList(NuSer).Pos.X
                                        UserList(NuSer).flags.MapY = UserList(NuSer).Pos.Y
                                        If (respawnUser) Then
                                                If (UserList(NuSer).flags.Muerto) Then
                                                        Call RevivirUsuario(NuSer)
                                                End If
                         
                                                UserList(NuSer).Stats.MinHp = UserList(NuSer).Stats.MaxHp
                                                UserList(NuSer).Stats.MinMAN = UserList(NuSer).Stats.MaxMAN
                                                UserList(NuSer).Stats.MinHam = 100
                                                UserList(NuSer).Stats.MinAGU = 100
                                                UserList(NuSer).Stats.MinSta = UserList(NuSer).Stats.MaxSta
                         
                                                Call Protocol.WriteUpdateUserStats(NuSer)
                                        Else
                                                UserList(NuSer).reto2Data.acceptedOK = False
                                        End If
                                End If
                        End If

                Next loopC

        End With

End Sub
 
Private Sub message_reto(ByRef retoStr As retoStruct, ByRef sMessage As String)
 
        '
        ' @ amishar
   
        With retoStr
 
                Dim i As Long
                Dim j As Long
                Dim u As Integer
         
                For i = 0 To 1
                        For j = 0 To 1
                                u = .team_array(i).user_Index(j)
                 
                                If (u <> 0) Then
                                        If (UserList(u).ConnID <> -1) Then
                                                Call Protocol.WriteConsoleMsg(u, sMessage, FontTypeNames.FONTTYPE_GUILD)
                                        End If
                                End If
 
                        Next j
                Next i
 
        End With
   
End Sub
 
Public Sub user_die_reto(ByVal user_Index As Integer)

        '
        ' @ amishar
    
        Dim team_Index As Integer
        Dim user_slot  As Integer
        Dim other_user As Integer
        Dim reto_Index As Integer
    
        reto_Index = UserList(user_Index).reto2Data.reto_Index
    
        team_Index = find_Team(user_Index, reto_Index)

        If (team_Index <> -1) Then
                user_slot = find_user(team_Index, user_Index, reto_Index)
                If team_Index = 0 Then
                WarpUserChar user_Index, UserList(user_Index).flags.MapU, UserList(user_Index).flags.MapX - 2, UserList(user_Index).flags.MapY, True
        ElseIf team_Index = 1 Then
        WarpUserChar user_Index, UserList(user_Index).flags.MapU, UserList(user_Index).flags.MapX + 2, UserList(user_Index).flags.MapY, True
            
        End If
        End If
 
        If (user_slot = -1) Then Exit Sub

        other_user = IIf(user_slot = 0, 1, 0)
        other_user = reto_List(reto_Index).team_array(team_Index).user_Index(other_user)
    
        'is dead?

        If (other_user) Then
                If UserList(other_user).flags.Muerto Then
                        Call team_winner(reto_Index, IIf(team_Index = 0, 1, 0))
                End If

        Else
                Call team_winner(reto_Index, IIf(team_Index = 0, 1, 0))
        End If
    
End Sub
 
Private Function find_Team(ByVal user_Index As Integer, _
                           ByVal reto_Index As Integer) As Integer
 
        '
        ' @ amishar
   
        Dim i As Long
        Dim j As Long
   
        For i = 0 To 1
                For j = 0 To 1
 
                        If reto_List(reto_Index).team_array(i).user_Index(j) = user_Index Then
                                find_Team = i
 
                                Exit Function
 
                        End If
 
                Next j
        Next i
 
        find_Team = -1
End Function
 
Private Function find_user(ByVal team_Index As Integer, _
                           ByVal user_Index As Integer, _
                           ByVal reto_Index As Integer) As Integer
 
        '
        ' @ amishar
   
        Dim i As Long
   
        For i = 0 To 1
 
                If reto_List(reto_Index).team_array(team_Index).user_Index(i) = user_Index Then
                        find_user = i
 
                        Exit Function
 
                End If
 
        Next i
   
        find_user = -1
 
End Function
 
Private Sub team_winner(ByVal reto_Index As Integer, ByVal team_winner As Byte)
 
        '
        ' @ amishar
   
        With reto_List(reto_Index)
                .team_array(team_winner).round_count = (.team_array(team_winner).round_count + 1)
         
                If (.team_array(team_winner).round_count = 2) Then
                        Call finish_reto(reto_Index, team_winner)
                Else
                        Call respawn_reto(reto_Index, team_winner)
                End If
         
        End With
 
End Sub
 
Private Sub respawn_reto(ByVal reto_Index As Integer, ByVal team_winner As Integer)

'
' @ amishar

'Call warp_Teams(reto_Index, True)

Dim loopx As Long
Dim loopC As Long
Dim mStr As String
Dim index As Integer

With reto_List(reto_Index)

mStr = "Ganador equipo de " & UserList(.team_array(team_winner).user_Index(0)).Name & " y " & UserList(.team_array(team_winner).user_Index(1)).Name & "." & vbNewLine & "Resultado parcial : " & CStr(.team_array(0).round_count) & "-" & CStr(.team_array(1).round_count)

For loopx = 0 To 1
For loopC = 0 To 1
index = .team_array(loopx).user_Index(loopC)

If (index <> 0) Then
If UserList(index).ConnID <> -1 Then
Call Protocol.WriteConsoleMsg(index, mStr, FontTypeNames.FONTTYPE_GUILD)
End If
End If

Next loopC
Next loopx

.nextRoundCount = 1

End With

End Sub
 
Private Sub finish_reto(ByVal reto_Index As Integer, _
                        ByVal team_winner As Byte, _
                        Optional ByVal bClose As Boolean = False)
 
        '
        ' @ amishar
   
        With reto_List(reto_Index)
         
                Dim retoMessage As String
                Dim team_looser As Byte
                Dim temp_index  As Integer
         
                retoMessage = get_reto_message(reto_Index)
         
retoMessage = retoMessage & ". Ganador equipo de " & UserList(.team_array(team_winner).user_Index(0)).Name & " y " & UserList(.team_array(team_winner).user_Index(1)).Name & "."

Call SendData(SendTarget.ToAll, 0, Protocol.PrepareMessageConsoleMsg(retoMessage, FontTypeNames.FONTTYPE_INFO))

                team_looser = IIf(team_winner = 0, 1, 0)
         
                Dim loopC  As Long
                Dim byDrop As Boolean
                Dim byGold As Long
         
                byDrop = (.general_rules.drop_inv = True)
                byGold = .general_rules.gold_gamble
         
                With .team_array(team_looser)
 
                        For loopC = 0 To 1
                                temp_index = .user_Index(loopC)
                 
                                UserList(temp_index).reto2Data.reto_used = False
                                UserList(temp_index).reto2Data.acceptedOK = False
                 
                                If (byDrop) Then
                                        Call TirarTodosLosItems(temp_index)
                                End If
                                   
                                Call WarpUserChar(temp_index, Arghal.Map, Arghal.X + loopC, Arghal.Y, True)
                                   
                                UserList(temp_index).Stats.GLD = (UserList(temp_index).Stats.GLD - byGold)
                 
                                UserList(temp_index).reto2Data.nick_sender = vbNullString
                                UserList(temp_index).reto2Data.reto_Index = 0
                 
                                Call Protocol.WriteUpdateGold(temp_index)
                 
                        Next loopC
 
                End With
         
                With .team_array(team_winner)
 
                        For loopC = 0 To 1
                                temp_index = .user_Index(loopC)
                 
                                UserList(temp_index).reto2Data.reto_used = False
                                UserList(temp_index).reto2Data.acceptedOK = False
                 
                                If (byDrop) Then
                                        UserList(temp_index).reto2Data.return_city = 15
                     
                                       Call Protocol.WriteConsoleMsg(temp_index, "Bienvenido a la sala de retos, tienes 15 segundos para recoger todos los items.", FontTypeNames.FONTTYPE_GUILD)
                                Else
                                        Call WarpUserChar(temp_index, 151, 30 + loopC, 72, True)
                                End If
                 
                                UserList(temp_index).Stats.GLD = (UserList(temp_index).Stats.GLD + byGold)
                 
                                UserList(temp_index).reto2Data.nick_sender = vbNullString
                                UserList(temp_index).reto2Data.reto_Index = 0
                 
                                Call Protocol.WriteUpdateGold(temp_index)
                 
                        Next loopC
 
                End With
         
                Call clear_data(reto_Index)
         
        End With
 
End Sub
 
Private Sub clear_data(ByVal reto_Index As Integer)
 
        '
        ' @ amishar
   
        With reto_List(reto_Index)
                .count_Down = 0
         
                With .general_rules
                        .drop_inv = False
                        .gold_gamble = 0
                End With
         
                .used_ring = False
         
                Dim i As Long
         
                For i = 0 To 1
       
                        .team_array(i).user_Index(0) = 0
                        .team_array(i).user_Index(1) = 0
                        .team_array(i).round_count = 0
 
                Next i
         
        End With
 
End Sub
 
Private Function get_reto_message(ByVal reto_Index As Integer) As String
 
        '
        ' @ amishar
   
        Dim tempStr  As String
        Dim tempUser As Integer
   
        With reto_List(reto_Index)
         
                tempStr = "Retos> "
         
                With .team_array(0)
                        tempUser = .user_Index(0)
             
                        If (tempUser <> 0) Then
                                If UserList(tempUser).ConnID <> -1 Then
                                        tempStr = tempStr & UserList(tempUser).Name
                                End If
                        End If
             
                        tempUser = .user_Index(1)
             
                        If (tempUser <> 0) Then
                                If UserList(tempUser).ConnID <> -1 Then
                                        tempStr = tempStr & " y " & UserList(tempUser).Name
                                End If
                        End If
             
                End With
         
                With .team_array(1)
                        tempUser = .user_Index(0)
             
                        If (tempUser <> 0) Then
                                If UserList(tempUser).ConnID <> -1 Then
                                        tempStr = tempStr & " vs " & UserList(tempUser).Name
                                End If
                        End If
             
                        tempUser = .user_Index(1)
             
                        If (tempUser <> 0) Then
                                If UserList(tempUser).ConnID <> -1 Then
                                        tempStr = tempStr & " y " & UserList(tempUser).Name
                                End If
                        End If
             
                End With
         
                With .general_rules
                        tempStr = tempStr & " apuesta " & Format$(.gold_gamble, "####") & " monedas de oro"
             
                        If (.drop_inv) Then
                                tempStr = tempStr & " y los items del inventario"
                        End If
 
                End With
         
        End With
   
        get_reto_message = tempStr
 
End Function
 
Public Function get_pos_x(ByVal ring_index As Integer, _
                          ByVal team_Index As Integer, _
                          ByVal user_Index As Integer)
 
        '
        ' @ amishar
   
        get_pos_x = reto_RingPos(ring_index, team_Index, user_Index).X
 
End Function
 
Public Function get_pos_y(ByVal ring_index As Integer, _
                          ByVal team_Index As Integer, _
                          ByVal user_Index As Integer)
 
        '
        ' @ matipanupro
   
        get_pos_y = reto_RingPos(ring_index, team_Index, user_Index).Y
   
End Function



