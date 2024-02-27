Attribute VB_Name = "mRanking"
Option Explicit

Public Const MAX_TOP As Byte = 10
Public Const MAX_RANKINGS As Byte = 6

Public Type tRanking
    value(1 To MAX_TOP) As Long
    Nombre(1 To MAX_TOP) As String
End Type

Public Ranking(1 To MAX_RANKINGS) As tRanking

Public Enum eRanking
    TopFrags = 1
    TopTorneos = 2
    TopLevel = 3
    TopOro = 4
    TopRetos = 5
    TopClanes = 6
End Enum



Public Function RenameRanking(ByVal Ranking As eRanking) As String

          'ASDJKASJKASKJ ESTE LO HICE YO GIL
          '@ Devolvemos el nombre del TAG [] del archivo .DAT
10        Select Case Ranking
              Case eRanking.TopClanes
20                RenameRanking = "Criminales Matados"
30            Case eRanking.TopFrags
40                RenameRanking = "Usuarios Matados"
50            Case eRanking.TopLevel
60                RenameRanking = "Ciudadanos Matados"
70            Case eRanking.TopOro
80                RenameRanking = "Oro"
90            Case eRanking.TopRetos
100               RenameRanking = "Retos"
110           Case eRanking.TopTorneos
120               RenameRanking = "Torneos"
130           Case Else
140               RenameRanking = vbNullString
150       End Select
End Function
Public Function RenameValue(ByVal Userindex As Integer, ByVal Ranking As eRanking) As Long
          ' @ Devolvemos a que hace referencia el ranking
10        With UserList(Userindex)
20            Select Case Ranking
                  Case eRanking.TopClanes
30                    RenameValue = .Faccion.CriminalesMatados
                      'RenameValue = guilds(.GuildIndex).Puntos
40                Case eRanking.TopFrags
50                    RenameValue = .Stats.UsuariosMatados
60                Case eRanking.TopLevel
70                    RenameValue = .Faccion.CiudadanosMatados
80                Case eRanking.TopOro
90                    RenameValue = .Stats.Gld
100               Case eRanking.TopRetos
110                   RenameValue = .Stats.RetosGanados
120               Case eRanking.TopTorneos
130                   RenameValue = .Stats.TorneosGanados
140           End Select
150       End With
End Function

Public Sub LoadRanking()
          ' @ Cargamos los rankings
          
          Dim LoopI As Integer
          Dim LoopX As Integer
          Dim ln As String
          
10        For LoopX = 1 To MAX_RANKINGS
20            For LoopI = 1 To MAX_TOP
30                ln = GetVar(App.Path & "\Dat\" & "Ranking.dat", RenameRanking(LoopX), "Top" & LoopI)
40                Ranking(LoopX).Nombre(LoopI) = ReadField(1, ln, 45)
50                Ranking(LoopX).value(LoopI) = val(ReadField(2, ln, 45))
60            Next LoopI
70        Next LoopX
          
End Sub
    
Public Sub SaveRanking(ByVal Rank As eRanking)
       ' @ Guardamos el ranking
       
          Dim LoopI As Integer
          
10            For LoopI = 1 To MAX_TOP
20                Call WriteVar(DatPath & "Ranking.Dat", RenameRanking(Rank), _
                      "Top" & LoopI, Ranking(Rank).Nombre(LoopI) & "-" & Ranking(Rank).value(LoopI))
30            Next LoopI
End Sub

Public Sub CheckRankingUser(ByVal Userindex As Integer, ByVal Rank As eRanking)
          ' @ Desde aca nos hacemos la siguientes preguntas
          ' @ El personaje está en el ranking?
          ' @ El personaje puede ingresar al ranking?
          
          Dim LoopX As Integer
          Dim LoopY As Integer
          Dim loopZ As Integer
          Dim i As Integer
          Dim value As Long
          Dim Actualizacion As Byte
          Dim Auxiliar As String
          Dim PosRanking As Byte
          
10        With UserList(Userindex)
              
              ' @ Not gms
20            If EsGM(Userindex) Then Exit Sub
              
30            value = RenameValue(Userindex, Rank)
              
              ' @ Buscamos al personaje en el ranking
40            For i = 1 To MAX_TOP
50                If Ranking(Rank).Nombre(i) = UCase$(.Name) Then
60                    PosRanking = i
70                    Exit For
80                End If
90            Next i
              
              ' @ Si el personaje esta en el ranking actualizamos los valores.
100           If PosRanking <> 0 Then
                  ' ¿Si está actualizado pa que?
110               If value <> Ranking(Rank).value(PosRanking) Then
120                   Call ActualizarPosRanking(PosRanking, Rank, value)
                      
                      
                      ' ¿Es la pos 1? No hace falta ordenarlos
130                   If Not PosRanking = 1 Then
                          ' @ Chequeamos los datos para actualizar el ranking
140                       For LoopY = 1 To MAX_TOP
150                           For loopZ = 1 To MAX_TOP - LoopY
                                      
160                               If Ranking(Rank).value(loopZ) < Ranking(Rank).value(loopZ + 1) Then
                                      
                                      ' Actualizamos el valor
170                                   Auxiliar = Ranking(Rank).value(loopZ)
180                                   Ranking(Rank).value(loopZ) = Ranking(Rank).value(loopZ + 1)
190                                   Ranking(Rank).value(loopZ + 1) = Auxiliar
                                      
                                      ' Actualizamos el nombre
200                                   Auxiliar = Ranking(Rank).Nombre(loopZ)
210                                   Ranking(Rank).Nombre(loopZ) = Ranking(Rank).Nombre(loopZ + 1)
220                                   Ranking(Rank).Nombre(loopZ + 1) = Auxiliar
230                                   Actualizacion = 1
240                               End If
250                           Next loopZ
260                       Next LoopY
270                   End If
                          
280                   If Actualizacion <> 0 Then
290                       Call SaveRanking(Rank)
300                   End If
310               End If
                  
320               Exit Sub
330           End If
              
              ' @ Nos fijamos si podemos ingresar al ranking
340           For LoopX = 1 To MAX_TOP
350               If value > Ranking(Rank).value(LoopX) Then
360                   Call ActualizarRanking(LoopX, Rank, .Name, value)
370                   Exit For
380               End If
390           Next LoopX
              
400       End With
End Sub

Public Sub ActualizarPosRanking(ByVal Top As Byte, ByVal Rank As eRanking, ByVal value As Long)
          ' @ Actualizamos la pos indicada en caso de que el personaje esté en el ranking
          Dim LoopX As Integer

10        With Ranking(Rank)
              
20            .value(Top) = value
30        End With
End Sub
Public Sub ActualizarRanking(ByVal Top As Byte, ByVal Rank As eRanking, ByVal UserName As String, ByVal value As Long)
          
          '@ Actualizamos la lista de ranking
          
          Dim LoopC As Integer
          Dim i As Integer
          Dim j As Integer
          Dim valor(1 To MAX_TOP) As Long
          Dim Nombre(1 To MAX_TOP) As String
          
          ' @ Copia necesaria para evitar que se dupliquen repetidamente
10        For LoopC = 1 To MAX_TOP
20            valor(LoopC) = Ranking(Rank).value(LoopC)
30            Nombre(LoopC) = Ranking(Rank).Nombre(LoopC)
40        Next LoopC
          
          ' @ Corremos las pos, desde el "Top" que es la primera
50        For LoopC = Top To MAX_TOP - 1
60            Ranking(Rank).value(LoopC + 1) = valor(LoopC)
70            Ranking(Rank).Nombre(LoopC + 1) = Nombre(LoopC)
80        Next LoopC


          
90        Ranking(Rank).Nombre(Top) = UCase$(UserName)
100       Ranking(Rank).value(Top) = value
110       Call SaveRanking(Rank)
120       Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ranking de " & RenameRanking(Rank) & "»" & UserName & " ha subido al TOP " & Top & ".", FontTypeNames.FONTTYPE_GUILD))
End Sub



