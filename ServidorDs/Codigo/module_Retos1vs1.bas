Attribute VB_Name = "module_Retos1vs1"
' programado por maTih.-
' version original el dia : 26/04/2012

Option Explicit

Const RETOS_ARENAS      As Byte = 5     'NUM DE ARENAS.
Const RETOS_VOLVER      As Byte = 30    'TIEMPO PARA VOLVER DESP DE GANAR.
Const RETOS_CUENTA      As Byte = 5     'SEGUNDOS DE CUENTA
Const RETOS_MAPA        As Integer = 5  'NUMERO DE MAPA.

Type Datos
     Usuarios(1 To 2)   As Integer      'UI DE LOS USUARIOS.
     Cuenta             As Byte         'CUENTA REGRESIVA.
     PorInventario      As Boolean      'SI ES POR ITEMS.
     ApuestaOro         As Long         'CANTIDAD DE ORO.
     SalaOcupada        As Boolean      'PARA BUSCAR RINGS VACIOS.
     Ganador            As Integer      'UI DEL GANADOR DEL RETO.
End Type

Public Retos(1 To RETOS_ARENAS) As Datos

Function PuedeEnviar(ByVal UserIndex As Integer, ByVal otherUser As String, ByVal Oro As Long, ByRef error As String) As Boolean

' @ Checks si puede enviar reto

PuedeEnviar = False

Dim OtherUI As Integer

With UserList(UserIndex)

    'Muerto.
    If .flags.Muerto <> 0 Then
        error = "Estás muerto, no puedes usar los retos!"
        Exit Function
    End If

    'Preso.
    If .Counters.Pena <> 0 Then
       error = "Estás en la carcel, no puedes usar los retos!"
       Exit Function
    End If
    
    'No tiene el oro.
    If .Stats.GLD < Oro Then
       error = "No tienes el oro por el que quieres jugar el reto!"
       Exit Function
    End If
    
    'Ya en reto.
    If .Reto1vs1.RetoIndex <> 0 Then
       error = "Ya estás en un reto!"
       Exit Function
    End If
    
End With

OtherUI = NameIndex(otherUser)
    
    'No online.
    If Not OtherUI <> 0 Then
        error = "El usuario " & otherUser & " no se encuentra online!"
        Exit Function
    End If

With UserList(OtherUI)

    'Muerto.
    If .flags.Muerto <> 0 Then
        error = "Está muerto, no puede usar los retos!"
        Exit Function
    End If

    'Preso.
    If .Counters.Pena <> 0 Then
       error = "Está en la carcel, no puede usar los retos!"
       Exit Function
    End If
    
    'No tiene el oro.
    If .Stats.GLD < Oro Then
       error = "No tiene el oro por el que quieres jugar el reto!"
       Exit Function
    End If
    
    'Ya en reto.
    If .Reto1vs1.RetoIndex <> 0 Then
       error = "Ya está en un reto!"
       Exit Function
    End If
    
End With

    'No hay salas.
    If Not SalaLibre <> 0 Then
       error = "Todas las salas de retos están ocupadas!"
       Exit Function
    End If
    
PuedeEnviar = True

End Function

Function DameX(ByVal Usuario As Byte, ByVal RetoIndex As Byte)

' @ Devuelve una posición X para un usuario y un reto.

Select Case RetoIndex

       Case 1   '<Arena 1.
            If Not Usuario <> 1 Then
                DameX = 50
            Else
                DameX = 55
            End If
            
       Case 2   '<Arena 2.
            If Not Usuario <> 1 Then
                DameX = 50
            Else
                DameX = 55
            End If
            
       Case 3   '<Arena 3.
            If Not Usuario <> 1 Then
                DameX = 50
            Else
                DameX = 55
            End If
            
       Case 4   '<Arena 4.
            If Not Usuario <> 1 Then
                DameX = 50
            Else
                DameX = 55
            End If
            
       Case 5   '<Arena 6.
            If Not Usuario <> 1 Then
                DameX = 50
            Else
                DameX = 55
            End If
            
       Case 6   '<Arena 6.
            If Not Usuario <> 1 Then
                DameX = 50
            Else
                DameX = 55
            End If
End Select

End Function

Function DameY(ByVal Usuario As Byte, ByVal RetoIndex As Byte)

' @ Devuelve una posición Y para un usuario y un reto.

Select Case RetoIndex

       Case 1   '<Arena 1.
            If Not Usuario <> 1 Then
                DameY = 50
            Else
                DameY = 55
            End If
            
       Case 2   '<Arena 2.
            If Not Usuario <> 1 Then
                DameY = 50
            Else
                DameY = 55
            End If
            
       Case 3   '<Arena 3.
            If Not Usuario <> 1 Then
                DameY = 50
            Else
                DameY = 55
            End If
            
       Case 4   '<Arena 4.
            If Not Usuario <> 1 Then
                DameY = 50
            Else
                DameY = 55
            End If
            
       Case 5   '<Arena 6.
            If Not Usuario <> 1 Then
                DameY = 50
            Else
                DameY = 55
            End If
            
       Case 6   '<Arena 6.
            If Not Usuario <> 1 Then
                DameY = 50
            Else
                DameY = 55
            End If
End Select

End Function

Function SalaLibre() As Byte

' @ Busca una arena que no esté usada.

Dim loopX   As Long

For loopX = 1 To RETOS_ARENAS
    If Not Retos(loopX).SalaOcupada Then
       SalaLibre = CByte(loopX)
       Exit Function
    End If
Next loopX

SalaLibre = 0

End Function

Sub PasaSegundo()

' @ Pasa un segundo.

Dim loopX   As Long

For loopX = 1 To RETOS_ARENAS

    With Retos(loopX)
         'Hay reto?
         If .SalaOcupada Then
            'Cuenta?
            If .Cuenta <> 0 Then
               'Envia.
               WriteConsoleMsg .Usuarios(1), "Comienza en : " & .Cuenta, FontTypeNames.FONTTYPE_CITIZEN
               WriteConsoleMsg .Usuarios(2), "Comienza en : " & .Cuenta, FontTypeNames.FONTTYPE_CITIZEN
               'Resta.
               .Cuenta = .Cuenta - 1
               'Llega a 0?
               If Not .Cuenta <> 0 Then
                  'Despausea.
                  WritePauseToggle .Usuarios(1)
                  WritePauseToggle .Usuarios(2)
                  'Avisa
                  WriteConsoleMsg .Usuarios(1), "El reto ha comenzado!", FontTypeNames.FONTTYPE_CITIZEN
                  WriteConsoleMsg .Usuarios(2), "El reto ha comenzado!", FontTypeNames.FONTTYPE_CITIZEN
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
                     Call WarpUserChar(.Ganador, UserList(.Ganador).Reto1vs1.AnteriorPosition.map, UserList(.Ganador).Reto1vs1.AnteriorPosition.X, UserList(.Ganador).Reto1vs1.AnteriorPosition.Y, True)
                     'Reset usuario y slot.
                     Call module_Retos1vs1.Limpiar(.Ganador)
                     Call module_Retos1vs1.LimpiarIndex(loopX)
                  End If
               End If
            End If
         End If
         
    End With

Next loopX

End Sub

Sub Enviar(ByVal UserIndex As Integer, ByVal otherIndex As Integer, ByVal Apuesta As Long, ByVal Inventario As Boolean)

' @ Envia reto.

Dim nextStr As String

With UserList(UserIndex)
    
    'buffer para los datos.
    With .Reto1vs1
         .ApuestaInv = Inventario
         .ApuestaOro = Apuesta
    End With
    
    'Prepara el mensaje
    If Apuesta <> 0 Then
       nextStr = "Apostando " & Format$(Apuesta, "#,###") & " monedas de oro"
    End If
    
    If Inventario Then
       nextStr = " y los items del inventario"
    End If
    
    'Avisa al usuario.
    
    WriteConsoleMsg otherIndex, .name & " Te desafia en un reto 1vs1 " & nextStr & " tipea /ACEPTAR ó /RECHAZAR según tu desición", FontTypeNames.FONTTYPE_CITIZEN

End With

'Datos del otro usuario.
With UserList(otherIndex).Reto1vs1
     .MeEnvio = UserIndex
End With

'Avisa.
WriteConsoleMsg UserIndex, "El reto a sido enviado.", FontTypeNames.FONTTYPE_CITIZEN

End Sub

Sub Aceptar(ByVal UserIndex As Integer)

' @ Usuario acepta reto.

Dim LibreSlot   As Byte

With UserList(UserIndex)

     'Nadie lo reta.
     If Not .Reto1vs1.MeEnvio <> 0 Then Exit Sub
          
     'Busca slot.
     LibreSlot = SalaLibre
     
     'No hay sala.
     If Not LibreSlot <> 0 Then
        WriteConsoleMsg UserIndex, "No hay salas disponibles actualmente.", FontTypeNames.FONTTYPE_CITIZEN
        Exit Sub
     End If
     
     'No está online.
     If Not UserList(.Reto1vs1.MeEnvio).ConnID <> -1 Then
        WriteConsoleMsg UserIndex, "El usuario se ha desconectado.", FontTypeNames.FONTTYPE_CITIZEN
        Exit Sub
     End If
     
     'Que empieze el reto!
     Empezar UserIndex, .Reto1vs1.MeEnvio, LibreSlot
     
End With

End Sub

Sub Empezar(ByVal UserIndex As Integer, ByVal EnviadorIndex As Integer, ByVal Slot As Byte)

' @ Empieza un nuevo reto.

'Llena los datos.

Dim loopX   As Long

With Retos(Slot)
     
     'Setea los UI.
     .Usuarios(1) = EnviadorIndex
     .Usuarios(2) = UserIndex
     
     'Guarda apuestas.
     .ApuestaOro = UserList(EnviadorIndex).Reto1vs1.ApuestaOro
     .PorInventario = UserList(EnviadorIndex).Reto1vs1.ApuestaInv
     
     'Setea cuenta regresiva.
     .Cuenta = RETOS_CUENTA
     
     'Setea sala ocupada y resetea ganador UI
     .SalaOcupada = True
     .Ganador = 0
     
     For loopX = 1 To 2
         'Setea anteriorPos
         UserList(.Usuarios(loopX)).Reto1vs1.AnteriorPosition = UserList(.Usuarios(loopX)).Pos
         'Telep a los usuarios.
         Call Usuarios.WarpUserChar(.Usuarios(loopX), RETOS_MAPA, DameX(loopX, Slot), DameY(loopX, Slot), True)
         'Pause clientes.
         Call Protocol.WritePauseToggle(.Usuarios(loopX))
         'Cuenta regresiva.
         Call Protocol.WriteConsoleMsg(.Usuarios(loopX), "El reto iniciará en " & RETOS_CUENTA & " segundos.", FontTypeNames.FONTTYPE_CITIZEN)
         'Setea retoIndex
         UserList(.Usuarios(loopX)).Reto1vs1.RetoIndex = Slot
     Next loopX
     
     'Avistage to WORLD !!
     SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El reto de " & .Usuarios(1) & " vs " & .Usuarios(2) & " ha dado inicio!", FontTypeNames.FONTTYPE_CITIZEN)
     
End With

End Sub

Sub Muere(ByVal muertoIndex As Integer, Optional ByVal Desconexion As Boolean = False)

' @ Muere un usuario en reto

Dim winnerIndex As Integer  'UI DEL GANADOR DEL RETO.
Dim indexUser   As Byte     'INDEX DE LOS USUARIOS DEL RETO.
Dim indexReto   As Byte

indexReto = UserList(muertoIndex).Reto1vs1.RetoIndex

indexUser = IIf(Retos(indexReto).Usuarios(1) = muertoIndex, 2, 1)

'OBTENGO SU UI.
winnerIndex = Retos(indexReto).Usuarios(indexUser)

'setea reto ganado.
'UserList(winnerIndex).BalanceReto.Ganados = UserList(winnerIndex).BalanceReto.Ganados + 1

'setea reto perdi2.
'UserList(muertoIndex).BalanceReto.Perdidos = UserList(muertoIndex).BalanceReto.Perdidos + 1

'ERA POR ORO
If Retos(indexReto).ApuestaOro <> 0 Then
   'Da el oro.
   UserList(winnerIndex).Stats.GLD = UserList(winnerIndex).Stats.GLD + Retos(indexReto).ApuestaOro
   'Update cliente.
   Call Protocol.WriteUpdateGold(winnerIndex)
   'Has ganado blabla
   Call Protocol.WriteConsoleMsg(winnerIndex, "Has ganado " & Format$(Retos(indexReto).ApuestaOro, "#,###") & " monedas de oro.", FontTypeNames.FONTTYPE_CITIZEN)
End If

If Desconexion Then
   SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(muertoIndex).name & " Se desconectó en un reto.", FontTypeNames.FONTTYPE_CITIZEN)
End If

'ERA POR OBJETOS?
If Retos(indexReto).PorInventario Then
   'Lo ejecuto.
   Call TirarTodosLosItems(muertoIndex)
   'Lo devuelvo a su posición..
   Call WarpUserChar(muertoIndex, UserList(muertoIndex).Reto1vs1.AnteriorPosition.map, UserList(muertoIndex).Reto1vs1.AnteriorPosition.X, UserList(muertoIndex).Reto1vs1.AnteriorPosition.Y, True)
   'Seteo el ganador.
   Retos(indexReto).Ganador = winnerIndex
   UserList(winnerIndex).Reto1vs1.VolverSeg = RETOS_VOLVER
   'Avisa.
   SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(winnerIndex).name & " venció en el reto frente a " & UserList(muertoIndex).name & ", apostando los items del inventario y " & Format$(Retos(indexReto).ApuestaOro, "#,###") & " monedas de oro", FontTypeNames.FONTTYPE_CITIZEN)
   WriteConsoleMsg winnerIndex, "Tienes " & (RETOS_VOLVER) & " segundos para agarrar los objetos antes de ser teletransportado a tu anterior posición.", FontTypeNames.FONTTYPE_DIOS
   'Limpia al usuario
   Limpiar muertoIndex
   'Cierra.
   Exit Sub
End If

'Los devuelvo a su posición..
Call WarpUserChar(muertoIndex, UserList(muertoIndex).Reto1vs1.AnteriorPosition.map, UserList(muertoIndex).Reto1vs1.AnteriorPosition.X, UserList(muertoIndex).Reto1vs1.AnteriorPosition.Y, True)
Call WarpUserChar(winnerIndex, UserList(winnerIndex).Reto1vs1.AnteriorPosition.map, UserList(winnerIndex).Reto1vs1.AnteriorPosition.X, UserList(winnerIndex).Reto1vs1.AnteriorPosition.Y, True)
'Avisa al mundo.
SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(winnerIndex).name & " venció en el reto frente a " & UserList(muertoIndex).name & ", ganó " & Format$(Retos(indexReto).ApuestaOro, "#,###") & " monedas de oro", FontTypeNames.FONTTYPE_CITIZEN)
'Limpia el index del reto
LimpiarIndex indexReto
'Limpia los usuarios
Limpiar muertoIndex
Limpiar winnerIndex

End Sub

Sub Limpiar(ByVal cleanIndex As Integer)

' @ Limpia el tipo de un usuario.

Dim NoPos   As WorldPos

With UserList(cleanIndex).Reto1vs1
     .MeEnvio = 0
     .AnteriorPosition = NoPos
     .ApuestaInv = False
     .ApuestaOro = 0
     .VolverSeg = 0
     .RetoIndex = 0
End With

End Sub

Sub LimpiarIndex(ByVal RetoIndex As Byte)

' @ Limpia un slot de un reto.

With Retos(RetoIndex)

     .ApuestaOro = 0
     .PorInventario = False
     .Cuenta = 0
     .Ganador = 0
     .SalaOcupada = False
     .Usuarios(1) = 0
     .Usuarios(2) = 0

End With

End Sub

