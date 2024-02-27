Attribute VB_Name = "mInvasiones"
 Option Explicit
 
 Private Type tNpcInvasion
    Numero As Integer
    Amount As Byte
    DropIndex As Byte
End Type

' //PROPIAMENTE INVASIONES
Private Type tInvasion
    Activa As Boolean
    Name As String
    desc As String
    Npcs As Byte
    Npc() As tNpcInvasion
    map As Integer
End Type

Private Type ObjDrop
    Amount As Integer
    ObjIndex As Integer
    Probability As Byte
End Type

Private Type PointsDrop
    Points As Integer
    Probability As Byte
End Type

Private Type tDropInvasiones
    CantObj As Byte
    Obj() As ObjDrop
    
    Points As Integer
    
End Type


Public NumInvasiones As Byte
Public Invasiones() As tInvasion
Public DropInvasiones() As tDropInvasiones


' FIN INVASIONES

Public Sub LoadInvasiones()
    
    
    
    Dim LoopC As Integer
    Dim LoopB As Integer
    Dim strTemp As String
    Dim CantTemp As Integer
    
    Dim Read As clsIniManager
    Set Read = New clsIniManager

    Call Read.Initialize(App.Path & "\DAT\" & "INVASIONES.DAT")
    
    NumInvasiones = val(Read.GetValue("INIT", "NumInvasiones"))
    
    ReDim Invasiones(1 To NumInvasiones) As tInvasion
    
    ' /CARGAMOS LAS INVASIONES
    For LoopC = 1 To NumInvasiones
        With Invasiones(LoopC)
            .Name = vbNullString
            .desc = vbNullString
            .Npcs = Read.GetValue("INVASION" & LoopC, "CantNpcs")
            
            ReDim .Npc(1 To .Npcs) As tNpcInvasion
            
            For LoopB = 1 To .Npcs
                strTemp = Read.GetValue("INVASION" & LoopC, "NPC" & LoopB)
                .Npc(LoopB).Numero = val(ReadField(1, strTemp, Asc("-")))
                .Npc(LoopB).Amount = val(ReadField(2, strTemp, Asc("-")))
                .Npc(LoopB).DropIndex = val(ReadField(3, strTemp, Asc("-")))
            Next LoopB
        End With
    Next LoopC
    
    
    ' /CARGAMOS LOS DROPS :)
    CantTemp = val(Read.GetValue("INIT", "NumDrops"))
    
    ReDim DropInvasiones(1 To CantTemp) As tDropInvasiones
    
    For LoopC = 1 To CantTemp
        With DropInvasiones(LoopC)
            .Points = val(Read.GetValue("DROP" & LoopC, "Points"))
            .CantObj = val(Read.GetValue("DROP" & LoopC, "NumObj"))
            
            ReDim .Obj(1 To .CantObj) As ObjDrop
            
            For LoopB = 1 To .CantObj
                strTemp = Read.GetValue("DROP" & LoopC, "Obj" & LoopB)
                .Obj(LoopB).ObjIndex = val(ReadField(1, strTemp, Asc("-")))
                .Obj(LoopB).Amount = val(ReadField(2, strTemp, Asc("-")))
                .Obj(LoopB).Probability = val(ReadField(3, strTemp, Asc("-")))
            Next LoopB
        End With
    Next LoopC
    
    
End Sub

Public Sub NewInvasion(ByVal UserIndex As Integer, ByVal InvasionIndex As Byte, ByVal Name As String, ByVal desc As String, ByVal map As Integer)
    ' Creamos una nueva invasion.
    
    Dim strTemp As String
    
    
    With Invasiones(InvasionIndex)
        If .Activa Then
            WriteConsoleMsg UserIndex, "La invasión puede hacerse una sola vez.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
        .Activa = True
        .Name = Name
        .desc = desc
        .map = map
        
        RespawnInvasion InvasionIndex
    End With
    
    strTemp = "Invasion activa: " & Name & ", " & desc
    strTemp = strTemp & vbCrLf
    strTemp = strTemp & "Mapa de la invasión: Mapa" & map & " (" & MapInfo(map).Name & ")" & vbCrLf
    strTemp = strTemp & "Utiliza el comando /INFOEVENTO para saber toda la información sobre la invasión en curso."
    
    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strTemp, FontTypeNames.FONTTYPE_BRONCE)
    
End Sub

Private Sub RespawnInvasion(ByVal InvasionIndex As Integer)

    Dim NpcIndex As Integer
    Dim Pos As WorldPos
    Dim tmpCant As Byte
    Dim LoopC As Integer
    Dim LoopB As Integer
    
    
    With Invasiones(InvasionIndex)
        Pos.map = .map
        
        For LoopC = 1 To .Npcs
            For LoopB = 1 To .Npc(LoopC).Amount
                Pos.X = RandomNumber(20, 80)
                Pos.Y = RandomNumber(20, 80)
                
                NpcIndex = SpawnNpc(.Npc(LoopC).Numero, Pos, False, False)
                
                If NpcIndex = 0 Then
                    LogEventos "Npc nulo en invasiòn en curso"
                Else
                    Npclist(NpcIndex).Invasion = InvasionIndex
                    Npclist(NpcIndex).DropIndex = .Npc(LoopC).DropIndex
                End If
            Next LoopB
        Next LoopC
    End With
End Sub

Public Sub MuereNpcInvasion(ByVal UserIndex As Integer, ByVal InvasionIndex As Byte, ByVal DropIndex As Byte)
    
    Dim tmpStr As String
    Dim Obj As Obj
    Dim LoopC As Integer
    
    
    With DropInvasiones(DropIndex)
        If .CantObj > 0 Then
            For LoopC = 1 To .CantObj
                If RandomNumber(1, 100) <= .Obj(LoopC).Probability Then
                    Obj.ObjIndex = .Obj(LoopC).ObjIndex
                    Obj.Amount = .Obj(LoopC).Amount
                    
                    If Not MeterItemEnInventario(UserIndex, Obj) Then
                        LogEventos "El personaje " & UserList(UserIndex).Name & " no recibió el DropIndex"
                        WriteConsoleMsg UserIndex, "No has recibido el DropIndex de la criatura por falta de espacio. CONTACTA A UN GAME MASTER", FontTypeNames.FONTTYPE_INFO
                    Else
                        
                        If Not LoopC = .CantObj Then
                            tmpStr = tmpStr & "La criatura te ha dropeado: " & ObjData(Obj.ObjIndex).Name & " (" & Obj.Amount & ")" & vbCrLf
                        Else
                            tmpStr = tmpStr & "La criatura te ha dropeado: " & ObjData(Obj.ObjIndex).Name & " (" & Obj.Amount & ")"
                        End If
                    End If
                End If
            Next LoopC
            
        End If
        
        If .Points > 0 Then
            tmpStr = tmpStr & vbCrLf
            
            With UserList(UserIndex)
                .Stats.Points = .Stats.Points + DropInvasiones(DropIndex).Points
                tmpStr = "La criatura te ha dropeado PUNTOS DE CANJE: " & DropInvasiones(DropIndex).Points & "."
                WriteUpdatePoints UserIndex
            End With
        End If
        
        WriteConsoleMsg UserIndex, tmpStr, FontTypeNames.FONTTYPE_CENTINELA
    End With
End Sub

Public Function GenerateInfoInvasion() As String
    
    Dim LoopC As Integer
    Dim tmpStr As String
    
    If NumInvasiones = 0 Then Exit Function
    
    For LoopC = 1 To NumInvasiones
        If Invasiones(LoopC).Activa Then
            With Invasiones(LoopC)
                tmpStr = "INVASIÓN " & .Name & "» " & .desc & ". "
                tmpStr = tmpStr & " Deberás dirigirte al mapa " & .map & " (" & MapInfo(.map).Name & ")"
                tmpStr = tmpStr & vbCrLf
            End With
        End If
    Next LoopC
    GenerateInfoInvasion = tmpStr
End Function
Public Sub CloseInvasion(ByVal InvasionIndex As Byte)
    Dim X As Integer
    Dim Y As Integer
    Dim map As Integer
    
    With Invasiones(InvasionIndex)
        map = .map
        .Activa = False
    End With
    
    For Y = YMinMapSize To YMaxMapSize
           For X = XMinMapSize To XMaxMapSize
               With MapData(map, X, Y)
                    If .NpcIndex > 0 Then
                        If Npclist(.NpcIndex).Invasion > 0 Then
                            Call QuitarNPC(.NpcIndex)
                        End If
                    End If
               End With
            Next X
    Next Y
    
    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Invasión en el mapa: " & map & " cancelada por Game Master", FontTypeNames.FONTTYPE_WARNING)
End Sub


