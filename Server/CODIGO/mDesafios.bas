Attribute VB_Name = "mDesafios"
Option Explicit

Public Const CHALLENGE_GLD      As Long = 50000

Private Const MAX_MAP_CHALLENGE As Byte = 4

Public Type tMapChallenge

    Map As Integer
    X As Byte
    Y As Byte
    CHALLENGE_MAP_Y_DISTANCE As Byte

End Type

Public MapChallenge(1 To MAX_MAP_CHALLENGE) As tMapChallenge

Public Type tDataChallenge

    Users(1) As Integer
    MapSelected As Byte
    
End Type

Public Challenge As tDataChallenge

#If Classic = 0 Then
    ' Cargamos los Mapas de Desafios
    Public Sub Challenge_SetMap()
    
        Challenge.MapSelected = 1
    
        ' Desierto cl�sico
        With MapChallenge(1)
            .Map = 63
            .X = 59
            .Y = 38
            .CHALLENGE_MAP_Y_DISTANCE = 23

        End With
    
        ' Bosque
        With MapChallenge(2)
            .Map = 64
            .X = 57
            .Y = 35
            .CHALLENGE_MAP_Y_DISTANCE = 29

        End With
    
        ' Nieve
        With MapChallenge(3)
            .Map = 65
            .X = 55
            .Y = 34
            .CHALLENGE_MAP_Y_DISTANCE = 33

        End With
    
        ' Lava
        With MapChallenge(4)
            .Map = 66
            .X = 34
            .Y = 34
            .CHALLENGE_MAP_Y_DISTANCE = 31

        End With
    
    End Sub

#Else
' Cargamos los Mapas de Desafios
Public Sub Challenge_SetMap()
    
    Challenge.MapSelected = 1
    
    ' Desierto cl�sico
    With MapChallenge(1)
        .Map = 129
        .X = 57
        .Y = 36
        .CHALLENGE_MAP_Y_DISTANCE = 27
    End With
    
    ' Bosque
    With MapChallenge(2)
        .Map = 130
        .X = 57
        .Y = 36
        .CHALLENGE_MAP_Y_DISTANCE = 27
    End With
    
    ' Nieve
    With MapChallenge(3)
        .Map = 128
        .X = 56
        .Y = 34
        .CHALLENGE_MAP_Y_DISTANCE = 27
    End With
    
    ' Lava
    With MapChallenge(4)
        .Map = 127
        .X = 32
        .Y = 32
        .CHALLENGE_MAP_Y_DISTANCE = 27
    End With
    
End Sub

#End If

Public Sub Challenge_SelectedMap()

End Sub

Public Sub Desafio_UserAdd(ByVal UserIndex As Integer)

    With UserList(UserIndex)
                
        If .flags.Muerto Then Exit Sub
        
        If MapInfo(.Pos.Map).Pk Then
            WriteConsoleMsg UserIndex, "Debes estar en zona segura para participar del evento.", FontTypeNames.FONTTYPE_INFORED

            Exit Sub

        End If
        
        If .Stats.Elv < 40 Then
            WriteConsoleMsg UserIndex, "Debes ser al menos nivel 40 para participar de los desafios.", FontTypeNames.FONTTYPE_INFORED

            Exit Sub

        End If
        
        If .Stats.Gld < CHALLENGE_GLD Then
            WriteConsoleMsg UserIndex, "Debes poseer " & CHALLENGE_GLD & " Monedas de Oro participar de los desafios.", FontTypeNames.FONTTYPE_INFORED
            Exit Sub

        End If
        
        If .flags.Desafiando > 0 Then
            Call WriteConsoleMsg(UserIndex, "Ya est�s en desafio.", FontTypeNames.FONTTYPE_INFORED)

            Exit Sub

        End If
        
        If .Clase = eClass.Warrior Or .Clase = eClass.Hunter Then
            Call WriteConsoleMsg(UserIndex, "Tu Clase no participa en este tipo de eventos.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If
        
        .PosAnt = .Pos
        
        If Not (Challenge.Users(0) = 0) And Not (Challenge.Users(1) = 0) Then
            Call WriteConsoleMsg(UserIndex, "El desafio se est� realizando entre " & UserList(Challenge.Users(0)).Name & " vs " & UserList(Challenge.Users(1)).Name & ".", FontTypeNames.FONTTYPE_DESAFIOS)
            Exit Sub

        End If
        
        If Challenge.Users(0) = 0 Then
            Challenge.Users(0) = UserIndex
            
            EventWarpUser Challenge.Users(0), MapChallenge(Challenge.MapSelected).Map, MapChallenge(Challenge.MapSelected).X, MapChallenge(Challenge.MapSelected).Y
        ElseIf Challenge.Users(1) = 0 Then
            Challenge.Users(1) = UserIndex
            
            EventWarpUser Challenge.Users(1), MapChallenge(Challenge.MapSelected).Map, MapChallenge(Challenge.MapSelected).X, MapChallenge(Challenge.MapSelected).Y + MapChallenge(Challenge.MapSelected).CHALLENGE_MAP_Y_DISTANCE

        End If

        If Not (Challenge.Users(1) = 0) And Not (Challenge.Users(1) = 0) Then
        
            SendData SendTarget.toMapSecure, 0, PrepareMessageConsoleMsg("Desafios� " & UserList(Challenge.Users(0)).Name & " (" & ListaClases(UserList(Challenge.Users(0)).Clase) & " " & ListaRazas(UserList(Challenge.Users(0)).Raza) & " Lvl " & UserList(Challenge.Users(0)).Stats.Elv & ") vs  " & UserList(Challenge.Users(1)).Name & " (" & ListaClases(UserList(Challenge.Users(1)).Clase) & " " & ListaRazas(UserList(Challenge.Users(1)).Raza) & " Lvl " & UserList(Challenge.Users(1)).Stats.Elv & ")", FontTypeNames.FONTTYPE_DESAFIOS)
        
            EventWarpUser Challenge.Users(0), MapChallenge(Challenge.MapSelected).Map, MapChallenge(Challenge.MapSelected).X, MapChallenge(Challenge.MapSelected).Y
            EventWarpUser Challenge.Users(1), MapChallenge(Challenge.MapSelected).Map, MapChallenge(Challenge.MapSelected).X, MapChallenge(Challenge.MapSelected).Y + MapChallenge(Challenge.MapSelected).CHALLENGE_MAP_Y_DISTANCE
                  
            With UserList(Challenge.Users(0))
                .Stats.Gld = .Stats.Gld - CHALLENGE_GLD
                Call WriteUpdateGold(Challenge.Users(0))

            End With
            
            With UserList(Challenge.Users(1))
                .Stats.Gld = .Stats.Gld - CHALLENGE_GLD
                Call WriteUpdateGold(Challenge.Users(1))

            End With
            
            Challenge.MapSelected = Challenge.MapSelected + 1
        
            If Challenge.MapSelected >= MAX_MAP_CHALLENGE Then
                Challenge.MapSelected = 1

            End If

        Else
            SendData SendTarget.toMapSecure, 0, PrepareMessageConsoleMsg("Desafios� El personaje " & UserList(UserIndex).Name & " (" & ListaClases(UserList(UserIndex).Clase) & " " & ListaRazas(UserList(UserIndex).Raza) & " Lvl " & UserList(UserIndex).Stats.Elv & ") ha entrado a la sala.", FontTypeNames.FONTTYPE_DESAFIOS)

        End If
        
        .flags.Desafiando = 1

    End With

End Sub

Public Sub Desafio_UserKill(ByVal VictimIndex As Integer)

    '<EhHeader>
    On Error GoTo Desafio_UserKill_Err

    '</EhHeader>

    Dim AttackerIndex As Integer

    If Challenge.Users(0) = VictimIndex Then
        Challenge.Users(0) = 0
        AttackerIndex = Challenge.Users(1)
        
    ElseIf Challenge.Users(1) = VictimIndex Then
        Challenge.Users(1) = 0
        AttackerIndex = Challenge.Users(0)
       
    End If
    
    UserList(VictimIndex).flags.Desafiando = 0
    UserList(VictimIndex).flags.DesafiosGanados = 0
    'UserList(VictimIndex).Stats.DesafiosJugados = UserList(VictimIndex).Stats.DesafiosJugados + 1
    WriteConsoleMsg VictimIndex, "Has pasado a la siguiente sala. Obtiene 5 Victorias consecutivas y ganar�s tu primer Punto de Honor.", FontTypeNames.FONTTYPE_INFO
    WarpUserChar VictimIndex, 1, 27, 53, False

    If AttackerIndex = 0 Then Exit Sub
        
    UserList(AttackerIndex).Stats.Gld = UserList(AttackerIndex).Stats.Gld + CHALLENGE_GLD
    Call WriteUpdateGold(AttackerIndex)
    EventWarpUser AttackerIndex, MapChallenge(Challenge.MapSelected).Map, MapChallenge(Challenge.MapSelected).X, MapChallenge(Challenge.MapSelected).Y
    
    ' Variable Temporal
    'UserList(AttackerIndex).flags.DesafiosGanados = UserList(AttackerIndex).flags.DesafiosGanados + 1
    ' Variable Ranking
    'UserList(AttackerIndex).Stats.DesafiosGanados = UserList(AttackerIndex).Stats.DesafiosGanados + 1
    'UserList(AttackerIndex).Stats.DesafiosJugados = UserList(AttackerIndex).Stats.DesafiosJugados + 1
                            
    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Desafios� Gana " & UserList(AttackerIndex).Name & " (" & ListaClases(UserList(AttackerIndex).Clase) & " " & ListaRazas(UserList(AttackerIndex).Raza) & " Lvl " & UserList(AttackerIndex).Stats.Elv & ") y espera contrincante en la sala. Lleva " & UserList(AttackerIndex).flags.DesafiosGanados & " desafios ganados de forma consecutiva.", FontTypeNames.FONTTYPE_DESAFIOS)
    
    If General.AntiFrags_CheckUser(AttackerIndex, VictimIndex, 900) Then
        Desafio_CheckPremio AttackerIndex

    End If

    '<EhFooter>
    Exit Sub

Desafio_UserKill_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mDesafios.Desafio_UserKill " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Desafio_CheckPremio(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo Desafio_CheckPremio_Err

    '</EhHeader>

    Dim Sound  As Integer

    Dim Points As Integer
    
    Select Case UserList(UserIndex).flags.DesafiosGanados

        Case 2
            Sound = eSound.sDoubleKill

        Case 3
            Sound = eSound.sTripleKill

        Case 4
            Sound = eSound.sUltraKill

        Case 5
            Sound = eSound.sPerspal

        Case 10
            Sound = eSound.sHolyShit

        Case 15
            Sound = eSound.sUnstoppable

        Case 20
            Sound = eSound.sMonsterKill

    End Select
        
    If Sound > 0 Then Call SendData(SendTarget.toMapSecure, 0, PrepareMessagePlayEffect(Sound, NO_3D_SOUND, NO_3D_SOUND))
    
    If UserList(UserIndex).flags.DesafiosGanados >= 5 Then

        'UserList(UserIndex).Stats.Eldhir = UserList(UserIndex).Stats.Eldhir + 1
    End If
    
    '<EhFooter>
    Exit Sub

Desafio_CheckPremio_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mDesafios.Desafio_CheckPremio " & "at line " & Erl
        
    '</EhFooter>
End Sub
