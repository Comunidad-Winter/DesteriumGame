Attribute VB_Name = "mInvocation"
Option Explicit

' INVOCACIONES CON USUARIOS

Public Type tInvocaciones
    
    Activo As Byte
    
    'INFORMACION CARGADA
    Desc As String
    NpcIndex As Integer
    CantidadUsuarios As Byte
    mapa As Byte
    X() As Byte
    Y() As Byte
    
End Type

Public NumInvocaciones As Byte

Public Invocaciones()  As tInvocaciones

'[INIT]
'NumInvocaciones = 1

'[INVOCACION1] 'Mago del inframundo
'NpcIndex = 410

'Mapa = 1
'CantidadUsuarios = 2
'Pos1 = 40 - 60
'Pos2 = 70 - 80
Public Sub LoadInvocaciones()

    '<EhHeader>
    On Error GoTo LoadInvocaciones_Err

    '</EhHeader>
          
    Dim i        As Integer

    Dim X        As Integer

    Dim ln       As String
    
    Dim NpcIndex As Integer
    
    Dim Pos      As WorldPos
          
    NumInvocaciones = val(GetVar(DatPath & "Invocaciones.dat", "INIT", "NumInvocaciones"))
          
    ReDim Invocaciones(0 To NumInvocaciones) As tInvocaciones

    For i = 1 To NumInvocaciones

        With Invocaciones(i)
            .Activo = 0
            .CantidadUsuarios = val(GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "CantidadUsuarios"))
            .mapa = val(GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "Mapa"))
            .NpcIndex = val(GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "NpcIndex"))
            .Desc = GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "Desc")
                      
            ReDim .X(1 To .CantidadUsuarios)
            ReDim .Y(1 To .CantidadUsuarios)
                      
            For X = 1 To .CantidadUsuarios
                ln = GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "Pos" & X)
                          
                .X(X) = val(ReadField(1, ln, 45))
                .Y(X) = val(ReadField(2, ln, 45))
            Next X
                
            Pos.Map = .mapa
            Pos.X = .X(1)
            Pos.Y = .Y(1)
            
            NpcIndex = SpawnNpc(.NpcIndex, Pos, False, False)
            
            If NpcIndex Then
                .Activo = 1
                Npclist(NpcIndex).flags.Invocation = i
                Call UpdateInfoNpcs(Pos.Map, NpcIndex)

            End If
            
        End With

    Next i
          
    '<EhFooter>
    Exit Sub

LoadInvocaciones_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mInvocation.LoadInvocaciones " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function InvocacionIndex(ByVal mapa As Integer, _
                                ByVal X As Byte, _
                                ByVal Y As Byte) As Byte

    '<EhHeader>
    On Error GoTo InvocacionIndex_Err

    '</EhHeader>

    Dim i As Integer

    Dim Z As Integer
          
    InvocacionIndex = 0
          
    '// Devuelve el Index del mapa de invocación en el que está
    For i = 1 To NumInvocaciones

        With Invocaciones(i)

            For Z = 1 To .CantidadUsuarios

                If .mapa = mapa And (.X(Z) = X) And .Y(Z) = Y Then
                    InvocacionIndex = i

                    Exit For

                End If

            Next Z

        End With

    Next i
              
    '<EhFooter>
    Exit Function

InvocacionIndex_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mInvocation.InvocacionIndex " & "at line " & Erl
        
    '</EhFooter>
End Function

' if invocacacionindex = 0 then
Public Function PuedeSpawn(ByVal Index As Byte) As Boolean

    '<EhHeader>
    On Error GoTo PuedeSpawn_Err

    '</EhHeader>
          
    Dim Contador As Byte

    Dim i        As Integer
          
    PuedeSpawn = False

    For i = 1 To Invocaciones(Index).CantidadUsuarios

        If MapData(Invocaciones(Index).mapa, Invocaciones(Index).X(i), Invocaciones(Index).Y(i)).UserIndex Then
            Contador = Contador + 1
                  
            If Contador = Invocaciones(Index).CantidadUsuarios Then
                PuedeSpawn = True

            End If

        End If

    Next i
          
    '<EhFooter>
    Exit Function

PuedeSpawn_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mInvocation.PuedeSpawn " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function PuedeRealizarInvocacion(ByVal UserIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo PuedeRealizarInvocacion_Err

    '</EhHeader>
    PuedeRealizarInvocacion = False
          
    With UserList(UserIndex)

        If .flags.Muerto Then Exit Function
        If .flags.Mimetizado Then Exit Function
        If Not MapInfo(.Pos.Map).Pk Then Exit Function

    End With
          
    PuedeRealizarInvocacion = True
    '<EhFooter>
    Exit Function

PuedeRealizarInvocacion_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mInvocation.PuedeRealizarInvocacion " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub RealizarInvocacion(ByVal UserIndex As Integer, ByVal Index As Byte)

    '<EhHeader>
    On Error GoTo RealizarInvocacion_Err

    '</EhHeader>
          
    Dim Pos As WorldPos
          
    ' ¿Los usuarios están en las pos?
    If PuedeSpawn(Index) Then
              
        Dim NpcIndex As Integer

        Pos.Map = Invocaciones(Index).mapa
        Pos.X = RandomNumber(Invocaciones(Index).X(1) - 3, Invocaciones(Index).X(1) + 3)
        Pos.Y = RandomNumber(Invocaciones(Index).Y(1) - 3, Invocaciones(Index).Y(1) + 3)
              
        FindLegalPos UserIndex, Pos.Map, Pos.X, Pos.Y
        NpcIndex = SpawnNpc(Invocaciones(Index).NpcIndex, Pos, True, False)
              
        If Not NpcIndex = 0 Then
            Invocaciones(Index).Activo = 1
            Npclist(NpcIndex).flags.Invocation = Index
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Invocaciones(Index).Desc, FontTypeNames.FONTTYPE_GUILD))

        End If

    End If
          
    '<EhFooter>
    Exit Sub

RealizarInvocacion_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mInvocation.RealizarInvocacion " & "at line " & Erl
        
    '</EhFooter>
End Sub

