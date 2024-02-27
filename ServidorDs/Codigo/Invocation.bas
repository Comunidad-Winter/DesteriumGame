Attribute VB_Name = "Invocation"
Option Explicit

' INVOCACIONES CON USUARIOS

Public Type tInvocaciones
    
    Activo As Byte
    
    
    'INFORMACION CARGADA
    desc As String
    NpcIndex As Integer
    CantidadUsuarios As Byte
    Mapa As Byte
    X() As Byte
    Y() As Byte
    
        
    
    
End Type


Public NumInvocaciones As Byte
Public Invocaciones() As tInvocaciones

'[INIT]
'NumInvocaciones = 1

'[INVOCACION1] 'Mago del inframundo
'NpcIndex = 410

'Mapa = 1
'CantidadUsuarios = 2
'Pos1 = 40 - 60
'Pos2 = 70 - 80
Public Sub LoadInvocaciones()
          
          Dim i As Integer
          Dim X As Integer
          Dim ln As String
          
10        NumInvocaciones = val(GetVar(DatPath & "Invocaciones.dat", "INIT", "NumInvocaciones"))
          
20        ReDim Invocaciones(1 To NumInvocaciones) As tInvocaciones
30            For i = 1 To NumInvocaciones
40                With Invocaciones(i)
50                    .Activo = 0
60                    .CantidadUsuarios = val(GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "CantidadUsuarios"))
70                    .Mapa = val(GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "Mapa"))
80                    .NpcIndex = val(GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "NpcIndex"))
90                    .desc = GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "Desc")
                      
100                   ReDim .X(1 To .CantidadUsuarios)
110                   ReDim .Y(1 To .CantidadUsuarios)
                      
                      
120                   For X = 1 To .CantidadUsuarios
130                       ln = GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "Pos" & X)
                          
140                       .X(X) = val(ReadField(1, ln, 45))
150                       .Y(X) = val(ReadField(2, ln, 45))
160                   Next X
170               End With
180           Next i

          
End Sub

Public Function InvocacionIndex(ByVal Mapa As Byte, ByVal X As Byte, ByVal Y As Byte) As Byte

          Dim i As Integer
          Dim Z As Integer
          
10        InvocacionIndex = 0
          
          '// Devuelve el Index del mapa de invocación en el que está
20        For i = 1 To NumInvocaciones
30            With Invocaciones(i)
40                For Z = 1 To .CantidadUsuarios
50                    If .Mapa = Mapa And (.X(Z) = X) And .Y(Z) = Y Then
60                        InvocacionIndex = i
70                        Exit For
80                    End If
90                Next Z
100           End With
110       Next i
                  
              
End Function

' if invocacacionindex = 0 then
Public Function PuedeSpawn(ByVal index As Byte) As Boolean
          
          Dim Contador As Byte
          Dim i As Integer
          
10        PuedeSpawn = False
20        For i = 1 To Invocaciones(index).CantidadUsuarios
30            If MapData(Invocaciones(index).Mapa, Invocaciones(index).X(i), Invocaciones(index).Y(i)).Userindex Then
40                Contador = Contador + 1
                  
50                If Contador = Invocaciones(index).CantidadUsuarios Then
60                    PuedeSpawn = True
70                End If
80            End If
90        Next i
          
End Function

Public Function PuedeRealizarInvocacion(ByVal Userindex As Integer) As Boolean
10        PuedeRealizarInvocacion = False
          
20        With UserList(Userindex)
30            If .flags.Muerto Then Exit Function
40            If .flags.Mimetizado Then Exit Function
              
50        End With
          
          
60        PuedeRealizarInvocacion = True
End Function

Public Sub RealizarInvocacion(ByVal Userindex As Integer, ByVal index As Byte)
          
          Dim Pos As WorldPos
          
          ' ¿Los usuarios están en las pos?
10        If PuedeSpawn(index) Then
              
              Dim NpcIndex As Integer
20            Pos.map = Invocaciones(index).Mapa
30            Pos.X = RandomNumber(Invocaciones(index).X(1) - 5, Invocaciones(index).X(1) + 5)
40            Pos.Y = RandomNumber(Invocaciones(index).Y(1) - 5, Invocaciones(index).Y(1) + 5)
              
50            FindLegalPos Userindex, Pos.map, Pos.X, Pos.Y
60            NpcIndex = SpawnNpc(Invocaciones(index).NpcIndex, Pos, True, False)
              
70            If Not NpcIndex = 0 Then
80                Invocaciones(index).Activo = 1
90                Npclist(NpcIndex).flags.Invocacion = 1
100               Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Invocaciones(index).desc, FontTypeNames.FONTTYPE_GUILD))
110           End If
120       End If
          
          
End Sub


