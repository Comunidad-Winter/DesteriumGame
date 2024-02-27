Attribute VB_Name = "Retos1vs1"
'QUEDATE QUIETO PAJA

Option Explicit

Const RETOS_ARENAS As Byte = 9     'NUM DE ARENAS.
Const RETOS_VOLVER As Byte = 10    'TIEMPO PARA VOLVER DESP DE GANAR.
Const RETOS_CUENTA As Byte = 10     'SEGUNDOS DE CUENTA
Const RETOS_MAPA As Integer = 176  'NUMERO DE MAPA.

Type Datos
    Usuarios(1 To 2) As Integer      'UI DE LOS USUARIOS.
    Cuenta     As Byte         'CUENTA REGRESIVA.
    PorInventario As Boolean      'SI ES POR ITEMS.
    ApuestaOro As Long         'CANTIDAD DE ORO.
    SalaOcupada As Boolean      'PARA BUSCAR RINGS VACIOS.
    Ganador    As Integer      'UI DEL GANADOR DEL RETO.
End Type

Public Retos(1 To RETOS_ARENAS) As Datos
Public ElSlot  As Byte

Function PuedeEnviar(ByVal Userindex As Integer, ByVal otherUser As String, ByVal Oro As Long, ByRef error As String) As Boolean

' @ Checks si puede enviar reto

    PuedeEnviar = False

    Dim OtherUI As Integer

    With UserList(Userindex)


        ' If Not UserList(.Reto1vs1.RetoIndex).flags.Peleando = 1 Then
        '  error = "MO PODE REGTAR VIGOGT3TETETEETTETETETE."
        ' Exit Function
        'End If

        If Not .Pos.map = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡No puedes ingresar si no estas en Ullathorpe!!.", FontTypeNames.FONTTYPE_INFO)
            error = "Para enviar retos debes encontrarte en la Ciudad de Ullathorpe."
            Exit Function
        End If

        'Muerto.
        If .flags.Muerto <> 0 Then
            error = "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos."
            Exit Function
        End If

        'Preso.
        If .Counters.Pena <> 0 Then
            error = "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos."
            Exit Function
        End If

        'No tiene el oro.
        If .Stats.Gld < Oro Then
            error = "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos."
            Exit Function
        End If


        If Oro < 30000 Then
            WriteConsoleMsg Userindex, "La apuesta mínima en un reto son 30.000 monedas de oro.", FontTypeNames.FONTTYPE_INFO
            Exit Function
        End If

        If .Stats.ELV < 38 Then
            error = "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos."
            Exit Function
        End If

        'Ya en reto.
        If .Reto1vs1.RetoIndex <> 0 Then
            error = "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos."
            Exit Function
        End If

    End With

    OtherUI = NameIndex(otherUser)

    'No online.
    If Not OtherUI <> 0 Then
        error = "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos."
        Exit Function
    End If

    With UserList(OtherUI)

        'Muerto.
        If .flags.Muerto <> 0 Then
            error = "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos."
            Exit Function
        End If


        If Not .Pos.map = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡No puedes ingresar si no estas en Ullathorpe!!.", FontTypeNames.FONTTYPE_INFO)
            error = "Para enviar retos debes encontrarte en la Ciudad de Ullathorpe."
            Exit Function
        End If

        'Preso.
        If .Counters.Pena <> 0 Then
            error = "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos."
            Exit Function
        End If

        'amb hack

        If UserList(Userindex).Name = UserList(OtherUI).Name Then
            error = "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos."
            Exit Function
        End If


        If Oro < 30000 Then
            WriteConsoleMsg Userindex, "La apuesta mínima en un reto son 30.000 monedas de oro.", FontTypeNames.FONTTYPE_INFO
            Exit Function
        End If


        If .Stats.ELV < 38 Then
            error = "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos."
            Exit Function
        End If

        'No tiene el oro.
        If .Stats.Gld < Oro Then
            error = "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos."
            Exit Function
        End If

        'Ya en reto.
        If .Reto1vs1.RetoIndex <> 0 Then
            error = "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos."
            Exit Function
        End If

    End With

    'No hay salas.
    If Not SalaLibre <> 0 Then
        error = "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos."
        Exit Function
    End If

    ' If Not UserList(UserIndex).flags.EnEvento = 0 Then
    'error = "GORDO."
    ' Exit Function
    ' End If


    PuedeEnviar = True

End Function

Function DameX(ByVal Usuario As Byte, ByVal RetoIndex As Byte)

' @ Devuelve una posición X para un usuario y un reto.

Select Case RetoIndex

       Case 1   '<Arena 1.
            If Not Usuario <> 1 Then
                DameX = 13
            Else
                DameX = 27
            End If
            
       Case 2   '<Arena 2.
            If Not Usuario <> 1 Then
                DameX = 13
            Else
                DameX = 27
            End If
            
       Case 3   '<Arena 3.
            If Not Usuario <> 1 Then
                DameX = 13
            Else
                DameX = 27
            End If
            
       Case 4   '<Arena 4.
            If Not Usuario <> 1 Then
                DameX = 44
            Else
                DameX = 58
            End If
            
       Case 5   '<Arena 6.
            If Not Usuario <> 1 Then
                DameX = 44
            Else
                DameX = 50
            End If
            
       Case 6   '<Arena 6.
            If Not Usuario <> 1 Then
                DameX = 44
            Else
                DameX = 58
            End If
            
        Case 7 '<Arena 7.
            If Not Usuario <> 1 Then
                DameX = 74
            Else
                DameX = 88
            End If
            
        Case 8 '<Arena 7.
            If Not Usuario <> 1 Then
                DameX = 74
            Else
                DameX = 88
            End If
            
        Case 9 '<Arena 7.
            If Not Usuario <> 1 Then
                DameX = 74
            Else
                DameX = 88
            End If
End Select

End Function

Function DameY(ByVal Usuario As Byte, ByVal RetoIndex As Byte)

' @ Devuelve una posición Y para un usuario y un reto.

Select Case RetoIndex

       Case 1   '<Arena 1.
            If Not Usuario <> 1 Then
                DameY = 18
            Else
                DameY = 28
            End If
            
       Case 2   '<Arena 2.
            If Not Usuario <> 1 Then
                DameY = 46
            Else
                DameY = 56
            End If
            
       Case 3   '<Arena 3.
            If Not Usuario <> 1 Then
                DameY = 74
            Else
                DameY = 84
            End If
            
       Case 4   '<Arena 4.**************
            If Not Usuario <> 1 Then
                DameY = 18
            Else
                DameY = 28
            End If
            
       Case 5   '<Arena 6.
            If Not Usuario <> 1 Then
                DameY = 46
            Else
                DameY = 56
            End If
            
       Case 6   '<Arena 6.
            If Not Usuario <> 1 Then
                DameY = 74
            Else
                DameY = 84
            End If
        
         Case 7   '<Arena 6.*****************************
            If Not Usuario <> 1 Then
                DameY = 18
            Else
                DameY = 28
            End If
        
         Case 8   '<Arena 6.
            If Not Usuario <> 1 Then
                DameY = 46
            Else
                DameY = 58
            End If
            
         Case 9   '<Arena 6.
            If Not Usuario <> 1 Then
                DameY = 74
            Else
                DameY = 84
            End If
End Select

End Function

Function SalaLibre() As Byte

' @ Busca una arena que no esté usada.

    Dim LoopX  As Long

    For LoopX = 1 To RETOS_ARENAS
        If Not Retos(LoopX).SalaOcupada Then
            SalaLibre = CByte(LoopX)
            Exit Function
        End If
    Next LoopX

    SalaLibre = 0

End Function

Sub PasaSegundo()

' @ Pasa un segundo.

    Dim LoopX  As Long

    For LoopX = 1 To RETOS_ARENAS

        With Retos(LoopX)
            'Hay reto?
            If .SalaOcupada Then
                'Cuenta?
                If .Cuenta <> 0 Then
                    'Envia.
                    WriteConsoleMsg .Usuarios(1), "Reto> " & .Cuenta, FontTypeNames.FONTTYPE_INFO
                    WriteConsoleMsg .Usuarios(2), "Reto>" & .Cuenta, FontTypeNames.FONTTYPE_INFO
                    'Resta.
                    .Cuenta = .Cuenta - 1
                    'Llega a 0?
                    If Not .Cuenta <> 0 Then
                        'Despausea.
                        WritePauseToggle .Usuarios(1)
                        WritePauseToggle .Usuarios(2)
                        'Avisa
                        WriteConsoleMsg .Usuarios(1), "YA!", FontTypeNames.FONTTYPE_FIGHT
                        WriteConsoleMsg .Usuarios(2), "YA!", FontTypeNames.FONTTYPE_FIGHT
                    End If
                End If

                'Hay ganador?
                If .Ganador <> 0 Then
                    'Está logeado?
                    If UserList(.Ganador).ConnID <> -1 Then
                        'Pasa tiempo.
                        UserList(.Ganador).Reto1vs1.VolverSeg = UserList(.Ganador).Reto1vs1.VolverSeg - 1
                        'Se acabó el tiempo.
                        If Not UserList(.Ganador).Reto1vs1.VolverSeg <> 0 Then
                            'Devuelve a la posición.
                            Call WarpUserChar(.Ganador, UserList(.Ganador).flags.BeforeMap, UserList(.Ganador).flags.BeforeX, UserList(.Ganador).flags.BeforeY, False)
                            'Reset usuario y slot.
                            Call Retos1vs1.Limpiar(.Ganador)
                            Call Retos1vs1.LimpiarIndex(LoopX)
                        End If
                    End If
                End If
            End If

        End With

    Next LoopX

End Sub

Sub Enviar(ByVal Userindex As Integer, ByVal otherIndex As Integer, ByVal Apuesta As Long, ByVal Inventario As Boolean)

' @ Envia reto.

    Dim nextStr As String

    With UserList(Userindex)


        UserList(otherIndex).Reto1vs1.WaitingReto = .Name

        If UserList(Userindex).flags.Comerciando Then Exit Sub



        If Not .Pos.map = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡No puedes ingresar si no estas en Ullathorpe!!.", FontTypeNames.FONTTYPE_INFO)
            error = "Para enviar retos debes encontrarte en la Ciudad de Ullathorpe."
        End If

        'buffer para los datos.
        With .Reto1vs1
            '.ApuestaInv = Inventario
            .ApuestaOro = Apuesta
        End With

        'Prepara el mensaje
        If Apuesta <> 0 Then
            nextStr = "Apuesta " & Format$(Apuesta) & " monedas de oro"
        End If


'Y lo del limite de potas? aja, no tiene jaja, es  asihmhhple el sistema ya entendi jaja, pensé que no te andaba pero lo tenias
'A ver si podes ver esto, cambio de sentido de asociado

        'If Inventario Then
        'nextStr = "" & Format$(Apuesta) & " y los item del inventario"
        'End If

        'Avisa al usuario.

        WriteConsoleMsg otherIndex, .Name & "(" & UserList(otherIndex).Stats.ELV & ") te ha retado por la " & nextStr & ". si aceptas escribe /RETAR " & .Name & ".", FontTypeNames.FONTTYPE_GUILD

    End With

    'Datos del otro usuario.
    With UserList(otherIndex).Reto1vs1
        .MeEnvio = Userindex
    End With
    'Avisa
    Call WriteConsoleMsg(Userindex, "Le enviaste reto a " & UserList(otherIndex).Name & " (" & UserList(otherIndex).Stats.ELV & "). por la cantidad " & nextStr & "", FontTypeNames.FONTTYPE_GUILD)
End Sub

Sub Aceptar(ByVal Userindex As Integer, ByVal NameIngresed As String)


' @ Usuario acepta reto.

    Dim LibreSlot As Byte

    With UserList(Userindex)
        Dim targetUserIndex As Integer
        targetUserIndex = NameIndex(NameIngresed)



        If UserList(Userindex).flags.Comerciando Then Exit Sub

        Dim OtroUserIndex As Integer

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu

            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(Userindex, "¡¡Comercio cancelado, no puedes comerciar mientras aceptas un desafío!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)

                Call LimpiarComercioSeguro(Userindex)
                Call Protocol.FlushBuffer(OtroUserIndex)
            End If
        End If

        If .Reto1vs1.WaitingReto = vbNullString Then Exit Sub

        If UCase$(.Reto1vs1.WaitingReto) <> NameIngresed Then
            WriteConsoleMsg Userindex, "Ese usuario no te retó, debes ingresar /RETAR " & .Reto1vs1.WaitingReto, FontTypeNames.FONTTYPE_GUILD
            Exit Sub
        End If

        'Nadie lo reta.
        If Not .Reto1vs1.MeEnvio <> 0 Then Exit Sub



        If UserList(Userindex).Stats.Gld < UserList(.Reto1vs1.MeEnvio).Reto1vs1.ApuestaOro Then
            WriteConsoleMsg Userindex, "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If

        'Busca slot.
        LibreSlot = SalaLibre


        If Not .Pos.map = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡No puedes ingresar si no estas en Ullathorpe!!.", FontTypeNames.FONTTYPE_INFO)
            WriteConsoleMsg Userindex, "Para aceptar el la solicitud de retos, debes encontrarte en la Ciudad de Ullathorpe", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If


        'No hay sala.
        If Not LibreSlot <> 0 Then
            WriteConsoleMsg Userindex, "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If

        'No está online.
        If Not UserList(.Reto1vs1.MeEnvio).ConnID <> -1 Then
            WriteConsoleMsg Userindex, "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If

        If .Reto1vs1.RetoIndex <> 0 Then
            WriteConsoleMsg Userindex, "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If

        If .flags.Peleando <> 0 Then
            WriteConsoleMsg Userindex, "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If

        If .flags.EnEvento <> 0 Then
            WriteConsoleMsg Userindex, "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If

        If .reto2Data.reto_Index <> 0 Then
            WriteConsoleMsg Userindex, "Requisito para retar inválido. Verifica tu oro y el de tu oponente y la condición de ambos.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If

        If Not UserList(.Reto1vs1.MeEnvio).Pos.map = 1 Then
            WriteConsoleMsg Userindex, "Tu retador no se encuntra en la ciudad de Ullathorpe.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If

        If UserList(.Reto1vs1.MeEnvio).Reto1vs1.RetoIndex <> 0 Then
            WriteConsoleMsg Userindex, "Ya esta en un reto.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
        If UserList(.Reto1vs1.MeEnvio).Stats.Gld - 20000 < 0 Then
            WriteConsoleMsg Userindex, "El retador no posee el oro para la comisión del reto (20.000 monedas de Oro.)", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
        If UserList(Userindex).Stats.Gld - 20000 < 0 Then
            WriteConsoleMsg Userindex, "No posees el oro para la comisión del reto (20.000 monedas de Oro.)", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
        UserList(.Reto1vs1.MeEnvio).Stats.Gld = UserList(.Reto1vs1.MeEnvio).Stats.Gld - 20000
        UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld - 20000
        WriteUpdateUserStats (.Reto1vs1.MeEnvio)
        WriteUpdateUserStats (Userindex)

        'Que empieze el reto!
        Empezar Userindex, .Reto1vs1.MeEnvio, LibreSlot
        ElSlot = LibreSlot

    End With

End Sub

Sub Empezar(ByVal Userindex As Integer, ByVal EnviadorIndex As Integer, ByVal Slot As Byte)

' @ Empieza un nuevo reto.

'Llena los datos.
    On Error GoTo Err
    Dim LoopX  As Long

1   With Retos(Slot)

        'Setea los UI.
2       .Usuarios(1) = Userindex
3       .Usuarios(2) = EnviadorIndex

        'Guarda apuestas.
4       .ApuestaOro = UserList(EnviadorIndex).Reto1vs1.ApuestaOro
5       .PorInventario = UserList(EnviadorIndex).Reto1vs1.ApuestaInv

        'Setea cuenta regresiva.
6       .Cuenta = RETOS_CUENTA

        UserList(EnviadorIndex).flags.Inmovilizado = 0
        UserList(Userindex).flags.Paralizado = 0
        'Setea sala ocupada y resetea ganador UI
7       .SalaOcupada = True
8       .Ganador = 0


        For LoopX = 1 To 2
            'Setea anteriorPos
10          UserList(.Usuarios(LoopX)).Reto1vs1.Anteriorposition = UserList(.Usuarios(LoopX)).Pos
            UserList(.Usuarios(LoopX)).flags.Round = 0
            UserList(EnviadorIndex).flags.Inmovilizado = 0
            UserList(Userindex).flags.Paralizado = 0
            'Telep a los usuarios.
11          Call Usuarios.WarpUserChar(.Usuarios(LoopX), RETOS_MAPA, DameX(LoopX, Slot), DameY(LoopX, Slot), True)
            'Pause clientes.
12          Call Protocol.WritePauseToggle(.Usuarios(LoopX))
            UserList(.Usuarios(LoopX)).flags.Round = 0
            UserList(EnviadorIndex).flags.Inmovilizado = 0
            UserList(Userindex).flags.Paralizado = 0
            'Cuenta regresiva.
            WriteConsoleMsg .Usuarios(LoopX), UserList(.Usuarios(1)).Name & " vs " & UserList(.Usuarios(2)).Name & ".", FontTypeNames.FONTTYPE_TALK
            UserList(.Usuarios(LoopX)).flags.Round = 0
            UserList(EnviadorIndex).flags.Inmovilizado = 0
            UserList(Userindex).flags.Paralizado = 0
            ' Setear mapas
            UserList(.Usuarios(LoopX)).flags.BeforeMap = UserList(.Usuarios(LoopX)).Pos.map
            UserList(.Usuarios(LoopX)).flags.BeforeX = UserList(.Usuarios(LoopX)).Pos.X
            UserList(.Usuarios(LoopX)).flags.BeforeY = UserList(.Usuarios(LoopX)).Pos.Y
            'Setea retoIndex
14          UserList(.Usuarios(LoopX)).Reto1vs1.RetoIndex = Slot
            'Setea Round
            UserList(.Usuarios(LoopX)).flags.Round = 0
            UserList(EnviadorIndex).flags.Inmovilizado = 0
            UserList(Userindex).flags.Paralizado = 0
15      Next LoopX

        'Avistage to WORLD !!

        'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El reto de " & UserIndex & " vs " & EnviadorIndex & " ha dado inicio!", FontTypeNames.FONTTYPE_CITIZEN)
    End With
Err:
    Debug.Print "Linea " & Erl()

End Sub

Sub Muere(ByVal muertoIndex As Integer, Optional ByVal Desconexion As Boolean = False)
' @ Muere un usuario en reto
    Dim winnerIndex As Integer  'UI DEL GANADOR DEL RETO.
    Dim indexUser As Byte     'INDEX DE LOS USUARIOS DEL RETO.
    Dim indexReto As Byte
    indexReto = UserList(muertoIndex).Reto1vs1.RetoIndex
    indexUser = IIf(Retos(indexReto).Usuarios(1) = muertoIndex, 2, 1)
    'OBTENGO SU UI.
    winnerIndex = Retos(indexReto).Usuarios(indexUser)

    If Desconexion Then
        If UserList(winnerIndex).ConnID > 0 Then
            UserList(winnerIndex).flags.Round = 3
        ElseIf UserList(muertoIndex).ConnID > 0 Then
            UserList(muertoIndex).flags.Round = 3
        End If
    End If
    If UserList(winnerIndex).flags.Round <= 2 And UserList(winnerIndex).flags.Muerto = 0 Then
        UserList(winnerIndex).flags.Round = UserList(winnerIndex).flags.Round + 1
        With UserList(muertoIndex)
            With Retos(indexReto)
                WarpUserChar .Usuarios(1), UserList(.Usuarios(1)).flags.BeforeMap, UserList(.Usuarios(1)).flags.BeforeX, UserList(.Usuarios(1)).flags.BeforeY, True
                WarpUserChar .Usuarios(2), UserList(.Usuarios(2)).flags.BeforeMap, UserList(.Usuarios(2)).flags.BeforeX, UserList(.Usuarios(2)).flags.BeforeY, True
            End With
            WriteConsoleMsg muertoIndex, "Retos> Resultado parcial:" & vbNewLine & "Retos> " & .Name & " " & .flags.Round & " - " & UserList(winnerIndex).Name & " " & UserList(winnerIndex).flags.Round & "", FontTypeNames.FONTTYPE_TALK
            WriteConsoleMsg winnerIndex, "Retos> Resultado parcial:" & vbNewLine & "Retos> " & .Name & " " & .flags.Round & " - " & UserList(winnerIndex).Name & " " & UserList(winnerIndex).flags.Round & "", FontTypeNames.FONTTYPE_TALK
            WritePauseToggle muertoIndex
            WritePauseToggle winnerIndex
            Dim i As Long
            Dim pt As Integer
            Dim pt1 As Integer
            For i = 1 To RETOS_ARENAS
                If Retos(i).Usuarios(indexUser) > 0 Then
                    pt = Retos(i).Usuarios(1)
                    pt1 = Retos(i).Usuarios(2)
                    If (Retos(i).Cuenta = 0 And Retos(i).SalaOcupada = True) And (UserList(pt).flags.Muerto = 1 Or UserList(pt1).flags.Muerto = 1) Then
                        Retos(i).Cuenta = RETOS_CUENTA
                    End If
                End If
            Next i
            RevivirUsuario muertoIndex
            .Stats.MinHp = .Stats.MaxHp
            .Stats.MinMAN = .Stats.MaxMAN
            .Stats.MinSta = .Stats.MaxSta
            With UserList(winnerIndex)
                .Stats.MinHp = .Stats.MaxHp
                .Stats.MinMAN = .Stats.MaxMAN
                .Stats.MinSta = .Stats.MaxSta
            End With
        End With
    ElseIf UserList(muertoIndex).flags.Round <= 2 And UserList(muertoIndex).flags.Muerto = 0 Then
        UserList(muertoIndex).flags.Round = UserList(muertoIndex).flags.Round + 1
        With UserList(winnerIndex)
            .Stats.MinHp = .Stats.MaxHp
            .Stats.MinMAN = .Stats.MaxMAN
            .Stats.MinSta = .Stats.MaxSta
            WarpUserChar winnerIndex, .flags.BeforeMap, .flags.BeforeX, .flags.BeforeY, True
            WarpUserChar muertoIndex, UserList(muertoIndex).flags.BeforeMap, UserList(muertoIndex).flags.BeforeX, UserList(muertoIndex).flags.BeforeY, True
            WriteConsoleMsg muertoIndex, "Retos> Resultado parcial:" & vbNewLine & "Retos> " & .Name & " " & .flags.Round & " - " & UserList(winnerIndex).Name & " " & UserList(winnerIndex).flags.Round & "", FontTypeNames.FONTTYPE_TALK
            WriteConsoleMsg winnerIndex, "Retos> Resultado parcial:" & vbNewLine & "Retos> " & .Name & " " & .flags.Round & " - " & UserList(winnerIndex).Name & " " & UserList(winnerIndex).flags.Round & "", FontTypeNames.FONTTYPE_TALK
            WritePauseToggle muertoIndex
            WritePauseToggle winnerIndex


            For i = 1 To RETOS_ARENAS
                If Retos(i).Usuarios(indexUser) > 0 Then
                    pt = Retos(i).Usuarios(1)
                    pt1 = Retos(i).Usuarios(2)
                    If (Retos(i).Cuenta = 0 And Retos(i).SalaOcupada = True) And (UserList(pt).flags.Muerto = 1 Or UserList(pt1).flags.Muerto = 1) Then
                        Retos(i).Cuenta = RETOS_CUENTA
                    End If
                End If
            Next i
            RevivirUsuario winnerIndex
            .Stats.MinHp = .Stats.MaxHp
            .Stats.MinMAN = .Stats.MaxMAN
            .Stats.MinSta = .Stats.MaxSta
            With UserList(muertoIndex)
                .Stats.MinHp = .Stats.MaxHp
                .Stats.MinMAN = .Stats.MaxMAN
                .Stats.MinSta = .Stats.MaxSta
            End With
        End With
    End If
    If UserList(winnerIndex).flags.Round >= 2 Then
        WritePauseToggle muertoIndex
        WritePauseToggle winnerIndex

        For i = 1 To RETOS_ARENAS
            If Retos(i).Usuarios(indexUser) > 0 Then
                pt = Retos(i).Usuarios(1)
                pt1 = Retos(i).Usuarios(2)
                If (Retos(i).Cuenta >= 1 And Retos(i).SalaOcupada = True) And (UserList(pt).flags.Muerto = 0 Or UserList(pt1).flags.Muerto = 0) Then
                    Retos(i).Cuenta = RETOS_CUENTA
                End If
            End If
        Next i
        'ERA POR ORO
        'setea reto ganado.
        UserList(winnerIndex).Stats.RetosGanados = UserList(winnerIndex).Stats.RetosGanados + 1
        UserList(winnerIndex).Stats.OroGanado = UserList(winnerIndex).Stats.OroGanado + Retos(indexReto).ApuestaOro
        'setea reto perdi2.
        UserList(muertoIndex).Stats.RetosPerdidos = UserList(muertoIndex).Stats.RetosPerdidos + 1
        UserList(muertoIndex).Stats.OroPerdido = UserList(muertoIndex).Stats.OroPerdido + Retos(indexReto).ApuestaOro
        If Retos(indexReto).ApuestaOro <> 0 Then
            'Da el oro.
            UserList(winnerIndex).Stats.Gld = UserList(winnerIndex).Stats.Gld + Retos(indexReto).ApuestaOro
            UserList(muertoIndex).Stats.Gld = UserList(muertoIndex).Stats.Gld - Retos(indexReto).ApuestaOro
            'Update cliente.
            Call Protocol.WriteUpdateGold(winnerIndex)
            WriteUpdateGold muertoIndex
            'Has ganado blabla
            Call Protocol.WriteConsoleMsg(winnerIndex, "Felicitaciones has ganado el monto de " & Format$(Retos(indexReto).ApuestaOro, "") & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
        End If
        UserList(muertoIndex).flags.Inmovilizado = 0
        UserList(muertoIndex).flags.Paralizado = 0
        'ERA POR OBJETOS?
        If Retos(indexReto).PorInventario Then
            'Lo ejecuto.
            Call TirarTodosLosItems(muertoIndex)
            'Lo devuelvo a su posición..
            Call WarpUserChar(muertoIndex, 49, 76, 90, True)
            'Seteo el ganador.
            Retos(indexReto).Ganador = winnerIndex
            UserList(winnerIndex).Reto1vs1.VolverSeg = RETOS_VOLVER
            'Avisa.
            With UserList(winnerIndex)
                SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Retos> " & UserList(winnerIndex).Name & " vs " & UserList(muertoIndex).Name & ". Ganador " & UserList(winnerIndex).Name & ". Apuesta por " & Retos(ElSlot).ApuestaOro & " monedas de oro y los items.", FontTypeNames.FONTTYPE_INFO)
                WriteConsoleMsg winnerIndex, "Retos> Bienvenido en " & (RETOS_VOLVER) & " segundos podras agarrar todos los item ganados en este Reto, luego deslogea y seras enviado a tu posición anterior.", FontTypeNames.FONTTYPE_INFO
                WritePauseToggle muertoIndex
                WritePauseToggle winnerIndex
                .flags.Inmovilizado = 0
                .flags.Paralizado = 0
            End With

            'Limpia al usuario


            Limpiar muertoIndex
            Retos1vs1.Limpiar muertoIndex
            Retos1vs1.Limpiar winnerIndex
            If UserList(winnerIndex).Reto1vs1.VolverSeg >= RETOS_VOLVER Then
                Call WarpUserChar(winnerIndex, 49, 76, 90, True)
            End If
            'Cierra.
            Exit Sub
        End If
        'Los devuelvo a su posición..
        Call WarpUserChar(muertoIndex, 49, 76, 89, True)
        Call WarpUserChar(winnerIndex, 49, 76, 90, True)


        'Avisa al mundo.

        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Retos> " & UserList(winnerIndex).Name & " vs " & UserList(muertoIndex).Name & ". Ganador " & UserList(winnerIndex).Name & ". Apuesta por " & Retos(ElSlot).ApuestaOro & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)    'Limpia el index del reto
        LimpiarIndex indexReto

        'Limpia los usuarios
        Retos1vs1.Limpiar muertoIndex
        Retos1vs1.Limpiar winnerIndex
    ElseIf UserList(muertoIndex).flags.Round >= 2 Then
        WritePauseToggle muertoIndex
        WritePauseToggle winnerIndex
        For i = 1 To RETOS_ARENAS
            If Retos(i).Usuarios(indexUser) > 0 Then
                pt = Retos(i).Usuarios(1)
                pt1 = Retos(i).Usuarios(2)
                If (Retos(i).Cuenta >= 1 And Retos(i).SalaOcupada = True) And (UserList(pt).flags.Muerto = 1 Or UserList(pt1).flags.Muerto = 1) Then
                    Retos(i).Cuenta = RETOS_CUENTA
                End If
            End If
        Next i
        If Retos(indexReto).ApuestaOro <> 0 Then
            'Da el oro.
            UserList(winnerIndex).Stats.Gld = UserList(winnerIndex).Stats.Gld - Retos(indexReto).ApuestaOro
            UserList(muertoIndex).Stats.Gld = UserList(muertoIndex).Stats.Gld + Retos(indexReto).ApuestaOro
            'Update cliente.
            Call Protocol.WriteUpdateGold(winnerIndex)
            WriteUpdateGold muertoIndex
            'Has ganado blabla
            Call Protocol.WriteConsoleMsg(winnerIndex, "Felicitaciones has ganado el monto de " & Format$(Retos(indexReto).ApuestaOro, "") & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
        End If
        UserList(muertoIndex).flags.Inmovilizado = 0
        UserList(muertoIndex).flags.Paralizado = 0
        'ERA POR OBJETOS
        If Retos(indexReto).PorInventario Then
            'Lo ejecuto.
            Call TirarTodosLosItems(winnerIndex)
            'Lo devuelvo a su posición..
            Call WarpUserChar(muertoIndex, 49, 76, 90, True)
            'Seteo el ganador.
            Retos(indexReto).Ganador = muertoIndex
            UserList(muertoIndex).Reto1vs1.VolverSeg = RETOS_VOLVER
            'Avisa.
            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Retos> " & UserList(muertoIndex).Name & " vs " & UserList(muertoIndex).Name & ". Ganador " & UserList(winnerIndex).Name & ". Apuesta por " & Retos(ElSlot).ApuestaOro & " monedas de oro y los items.", FontTypeNames.FONTTYPE_INFO)
            WriteConsoleMsg muertoIndex, "Retos> Bienvenido a la sala de retos. Tienes " & (RETOS_VOLVER) & " segundos para agarrar los objetos antes de ser teletransportado a tu anterior posición.", FontTypeNames.FONTTYPE_INFO
            'Limpia al usuario
            Limpiar winnerIndex
            Retos1vs1.Limpiar muertoIndex
            Retos1vs1.Limpiar winnerIndex
            If UserList(muertoIndex).Reto1vs1.VolverSeg >= RETOS_VOLVER Then
                Call WarpUserChar(muertoIndex, 49, 76, 90, True)
            End If
            'Cierra.
            Exit Sub
        End If
        'Los devuelvo a su posición.
        Call WarpUserChar(muertoIndex, 49, 76, 89, True)
        Call WarpUserChar(winnerIndex, 49, 76, 90, True)
        'Avisa al mundo.
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Retos> " & UserList(winnerIndex).Name & " vs " & UserList(muertoIndex).Name & ". Ganador " & UserList(muertoIndex).Name & ". Apuesta por " & Retos(ElSlot).ApuestaOro & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)    'Limpia el index del reto
        LimpiarIndex indexReto
        'Limpia los usuarios
        Retos1vs1.Limpiar muertoIndex
        Retos1vs1.Limpiar winnerIndex
    End If
End Sub

Sub Limpiar(ByVal cleanIndex As Integer)

' @ Limpia el tipo de un usuario.

    With UserList(cleanIndex).Reto1vs1
        .MeEnvio = 0
        '.ApuestaInv = False
        .ApuestaOro = 0
        .VolverSeg = 0
        .RetoIndex = 0
    End With

End Sub

Sub LimpiarIndex(ByVal RetoIndex As Byte)

' @ Limpia un slot de un reto.

    With Retos(RetoIndex)

        .ApuestaOro = 0
        '.PorInventario = False
        .Cuenta = 0
        .Ganador = 0
        .SalaOcupada = False
        .Usuarios(1) = 0
        .Usuarios(2) = 0
    End With

End Sub







