Attribute VB_Name = "mFacciones"

Public Const PATH_FACTION As String = "\DAT\FACTION.DAT"

Public MAX_FACTION        As Byte

Public Enum eTipoFaction

    r_None = 0
    r_Armada = 1
    r_Caos = 2

End Enum

Public Type tRange

    Text As String
    Frags As Long
    Elv As Byte
    Gld As Long
    
    MinDef As Integer
    MaxDef As Integer

End Type

Public Type tFaction

    Status As eTipoFaction
    FragsCiu As Long
    FragsCri As Long
    FragsOther As Long
    Range As Byte
    
    StartDate As String
    StartElv As Byte
    StartFrags As Integer
    
    ExFaction As eTipoFaction

End Type

Public Type tInfoFaction

    Name As String
    TeamFaction As eTipoFaction
    AttackFaction As Byte
    TotalRange As Byte
    Range() As tRange

End Type

Public InfoFaction() As tInfoFaction

'Rangos de las facciones con sus requisitos
Public Sub LoadFactions()

    '<EhHeader>
    On Error GoTo LoadFactions_Err

    '</EhHeader>

    Dim Read As clsIniManager

    Dim A    As Long, B As Long

    Dim Temp As String
   
    Set Read = New clsIniManager
   
    Read.Initialize App.Path & PATH_FACTION
   
    MAX_FACTION = val(Read.GetValue("INIT", "MAX_FACTION"))
   
    ReDim InfoFaction(1 To MAX_FACTION) As tInfoFaction
   
    For A = 1 To MAX_FACTION

        With InfoFaction(A)
            .Name = Read.GetValue("FACTION" & A, "Name")
          
            .TeamFaction = val(Read.GetValue("FACTION" & A, "TeamFaction"))
            .TotalRange = val(Read.GetValue("FACTION" & A, "TotalRange"))
            .AttackFaction = val(Read.GetValue("FACTION" & A, "AttackFaction"))
          
            ReDim .Range(0 To .TotalRange) As tRange
          
            For B = 0 To .TotalRange
                Temp = Read.GetValue("FACTION" & A, "Range" & B)
              
                .Range(B).Text = ReadField(1, Temp, Asc("-"))
                .Range(B).Frags = val(ReadField(2, Temp, Asc("-")))
                .Range(B).Elv = val(ReadField(3, Temp, Asc("-")))
                .Range(B).Gld = val(ReadField(4, Temp, Asc("-")))
                .Range(B).MinDef = val(ReadField(5, Temp, Asc("-")))
                .Range(B).MaxDef = val(ReadField(6, Temp, Asc("-")))
            Next B

        End With

    Next A
   
    Set Read = Nothing
   
    '<EhFooter>
    Exit Sub

LoadFactions_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mFacciones.LoadFactions " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Agregamos el usuario a la facción correspondiente
Public Sub Faction_AddUser(ByVal UserIndex As Integer, ByVal Faction As eTipoFaction)

    '<EhHeader>
    On Error GoTo Faction_AddUser_Err

    '</EhHeader>

    With UserList(UserIndex)

        If Not Faction_CheckRequired(UserIndex, Faction, 0) Then Exit Sub
        
        If .Faction.ExFaction > 0 Then
            Call WriteConsoleMsg(UserIndex, "Ya has pertenecido a una facción. Deberás pedir perdón para volver a ser miembro.", FontTypeNames.FONTTYPE_WARNING)

            Exit Sub

        End If
        
        If Faction = r_Armada Then
            If .Faction.FragsCiu > 0 Then
                Call WriteConsoleMsg(UserIndex, "Has asesinado gente inocente. Deberás pedir perdón para volver a ser miembro.", FontTypeNames.FONTTYPE_WARNING)

                Exit Sub

            End If

        End If
        
        .Faction.Status = Faction
        .Faction.Range = 0
        
        Call Faction_RewardUser(UserIndex)
        'Call RankUser_AddPoint(UserIndex, 5)
        WriteConsoleMsg UserIndex, "¡Te has enlistado! El rey ha decidido entregarte una armadura única para que te defienda al momento de defender su honor ¡Usala bien!", FontTypeNames.FONTTYPE_INFOGREEN
        
    End With

    '<EhFooter>
    Exit Sub

Faction_AddUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mFacciones.Faction_AddUser " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Faction_RemoveUser(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo Faction_RemoveUser_Err

    '</EhHeader>
    
    With UserList(UserIndex).Faction
        .ExFaction = .Status
        .Status = 0
        .Range = 0
        .StartDate = vbNullString
        .StartElv = 0
        .StartFrags = 0
        
        Call WriteConsoleMsg(UserIndex, "¡Facción removida!", FontTypeNames.FONTTYPE_INFO)
        
        Call Guilds_CheckAlineation(UserIndex, a_Neutral)

    End With
    
    '<EhFooter>
    Exit Sub

Faction_RemoveUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mFacciones.Faction_RemoveUser " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Function Faction_CheckRequired(ByVal UserIndex As Integer, _
                                       ByVal Faction As eTipoFaction, _
                                       ByVal FactionRange As Byte) As Boolean

    '<EhHeader>
    On Error GoTo Faction_CheckRequired_Err

    '</EhHeader>
    
    Dim Frags As Long
    
    With UserList(UserIndex)

        Select Case Faction

            Case eTipoFaction.r_Armada
                Frags = .Faction.FragsCri

            Case eTipoFaction.r_Caos
                Frags = .Faction.FragsCiu

        End Select
        
        If Frags < InfoFaction(Faction).Range(FactionRange).Frags Then
            WriteConsoleMsg UserIndex, "Necesitas " & InfoFaction(Faction).Range(FactionRange).Frags & " Asesinados. Y tu tienes '" & Frags & "'.", FontTypeNames.FONTTYPE_INFO
            Faction_CheckRequired = False

            Exit Function

        End If
        
        If .Stats.Elv < InfoFaction(Faction).Range(FactionRange).Elv Then
            WriteConsoleMsg UserIndex, "Mataste suficientes criminales, pero te faltan " & InfoFaction(Faction).Range(FactionRange).Elv - .Stats.Elv & " niveles para poder recibir la próxima recompensa.", FontTypeNames.FONTTYPE_INFO
            Faction_CheckRequired = False

            Exit Function

        End If
        
        If .Stats.Gld < InfoFaction(Faction).Range(FactionRange).Gld Then
            WriteConsoleMsg UserIndex, "Necesitas " & InfoFaction(Faction).Range(FactionRange).Gld & " monedas de oro para poder recibir la próxima recompensa.", FontTypeNames.FONTTYPE_INFO
            Faction_CheckRequired = False

            Exit Function

        End If

    End With
    
    Faction_CheckRequired = True
    '<EhFooter>
    Exit Function

Faction_CheckRequired_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mFacciones.Faction_CheckRequired " & "at line " & Erl
        
    '</EhFooter>
End Function

' Otorgamos armaduras iniciales
' Esto despues tiene que ir cargado desde .dat en el FACTION.DAT
Public Sub Faction_RewardUser(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo Faction_RewardUser_Err

    '</EhHeader>
    Dim Obj As Obj
    
    With UserList(UserIndex)
        Obj.Amount = 1
        
        If .Faction.Status = r_Armada Then

            Select Case .Clase

                Case eClass.Cleric, eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin

                    If (.Raza = eRaza.Humano) Or (.Raza = eRaza.Drow) Or (.Raza = eRaza.Elfo) Then
                        Obj.ObjIndex = 1046
                    Else
                        Obj.ObjIndex = 1496

                    End If
                    
                Case eClass.Paladin, eClass.Warrior

                    If (.Raza = eRaza.Humano) Or (.Raza = eRaza.Drow) Or (.Raza = eRaza.Elfo) Then
                        Obj.ObjIndex = 1044
                    Else
                        Obj.ObjIndex = 1045

                    End If
                    
                Case eClass.Mage

                    If (.Raza = eRaza.Humano) Or (.Raza = eRaza.Drow) Or (.Raza = eRaza.Elfo) Then
                        If .Genero = eGenero.Hombre Then
                            Obj.ObjIndex = 1049
                        Else
                            Obj.ObjIndex = 1048

                        End If

                    Else
                        Obj.ObjIndex = 1047

                    End If

            End Select
            
        Else
            ' Legión Oscura

            Select Case .Clase

                Case eClass.Cleric, eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin

                    If (.Raza = eRaza.Humano) Or (.Raza = eRaza.Drow) Or (.Raza = eRaza.Elfo) Then
                        Obj.ObjIndex = 1057
                    Else
                        Obj.ObjIndex = 1497

                    End If
                    
                Case eClass.Paladin, eClass.Warrior

                    If (.Raza = eRaza.Humano) Or (.Raza = eRaza.Drow) Or (.Raza = eRaza.Elfo) Then
                        Obj.ObjIndex = 1055
                    Else
                        Obj.ObjIndex = 1056

                    End If
                
                Case eClass.Mage

                    If (.Raza = eRaza.Humano) Or (.Raza = eRaza.Drow) Or (.Raza = eRaza.Elfo) Then
                        If .Genero = eGenero.Hombre Then
                            Obj.ObjIndex = 1060
                        Else
                            Obj.ObjIndex = 1059

                        End If

                    Else
                        Obj.ObjIndex = 1058

                    End If

            End Select
            
        End If
        
        If Not MeterItemEnInventario(UserIndex, Obj) Then
            Call Logs_User(.Name, eUser, eNone, "Enlistado en facción no le dió armadura inicial. Darsela manual")

        End If
        
    End With

    '<EhFooter>
    Exit Sub

Faction_RewardUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mFacciones.Faction_RewardUser " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Asignamos un rango al personaje.
' Este procedimiento se llama cada vez que un usuario mata a alguien opuesto a su rango.
Public Sub Faction_CheckRangeUser(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo Faction_CheckRangeUser_Err

    '</EhHeader>
    
    Dim Faction As eTipoFaction

    Dim Frags   As Long

    With UserList(UserIndex)

        If .Faction.Status = 0 Then
            Call WriteConsoleMsg(UserIndex, "No perteneces a ninguna facción", FontTypeNames.FONTTYPE_INFO)

            Exit Sub
        
        End If
        
        If (.Faction.Range) = InfoFaction(.Faction.Status).TotalRange Then Exit Sub
        If Not Faction_CheckRequired(UserIndex, .Faction.Status, .Faction.Range + 1) Then Exit Sub
        
        .Faction.Range = .Faction.Range + 1
        
        If InfoFaction(.Faction.Status).Range(.Faction.Range).Gld > 0 Then
            .Stats.Gld = .Stats.Gld - InfoFaction(.Faction.Status).Range(.Faction.Range).Gld
            Call WriteUpdateGold(UserIndex)

        End If
        
        ' Ultimo Rango Flod Consola.
        If .Faction.Range = InfoFaction(.Faction.Status).TotalRange Then
            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El personaje " & .Name & " ha alcanzado el último rango de su facción. Felicitaciones", FontTypeNames.FONTTYPE_INFO)

        End If
            
    End With

    '<EhFooter>
    Exit Sub

Faction_CheckRangeUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mFacciones.Faction_CheckRangeUser " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Comprobamos si un personaje puede atacar a otro.
' Las reglas son básicas. Usuarios de la misma facción se pueden atacar si la variable configurable AttackFaction=1
' Si son de distinta facción, podra atacar a la víctima si la variable TeamFaction no es igual al índice de facción del enemigo.
Public Function Faction_CanAttack(ByVal AttackerIndex As Integer, _
                                  ByVal VictimIndex As Integer)

    '<EhHeader>
    On Error GoTo Faction_CanAttack_Err

    '</EhHeader>

    Dim StatusAttacker As Byte

    Dim StatusVictim   As Byte
    
    Faction_CanAttack = False
    
    With UserList(AttackerIndex)
        
        StatusAttacker = .Faction.Status
        StatusVictim = UserList(VictimIndex).Faction.Status

        If StatusAttacker = StatusVictim Then
            If InfoFaction(StatusAttacker).AttackFaction > 0 Then
                Faction_CanAttack = True

                Exit Function

            End If
        
        Else

            If InfoFaction(StatusAttacker).TeamFaction <> StatusVictim Then
                Faction_CanAttack = True

                Exit Function

            End If
        
        End If
        
    End With

    '<EhFooter>
    Exit Function

Faction_CanAttack_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mFacciones.Faction_CanAttack " & "at line " & Erl
        
    '</EhFooter>
End Function
