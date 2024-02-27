Attribute VB_Name = "ModDeath"
Option Explicit
 
'Programado por JoaCo
 
Type DeathUser
     Userindex      As Integer      'UI del usuario.
     LastPosition   As WorldPos     'Pos que estaba antes de entrar.
     Esperando      As Byte         'Tiempo de espera para volver.
End Type
 
Type tDeath
     Cupos          As Byte         'Cantidad de cupos.
     Ingresaron     As Byte         'Cantidad que ingreso.
     Usuarios()     As DeathUser    'Tipo de usuarios
     Cuenta         As Byte         'Cuenta regresiva.
     Activo         As Boolean      'Hay deathmatch
     CaenObjs       As Boolean      'Caen objetos.
     AutoCancelTime As Byte         'Tiempo de auto-cancelamiento
     Ganador        As DeathUser    'Datos del ganador.
     BanqueroIndex  As Integer      'NPCindex del banquero..
End Type
 
Const CUENTA_NUM    As Byte = 5     'Segundos de cuenta.
Const ARENA_MAP     As Integer = 195 'Mapa de la arena.
Const ARENA_X       As Byte = 59    'X de la arena(se suma por usuario)
Const ARENA_Y       As Byte = 49    'Y de la arena.
Const BANCO_X       As Byte = 59   'X donde aparece el banquero.
Const BANCO_Y       As Byte = 48    'Y Donde aparece el banquro.
 
Const PREMIO_POR_CABEZA As Long = 17000 'Premio en oro , el cálculo es el de acá abajo.
Const TIEMPO_AUTOCANCEL As Byte = 180     '5 Minutos antes del auto-cancel.
Const TIEMPO_PARAVOLVER As Byte = 30     '2 Minutos para lukear objetos.
 
'Cálculo : PREMIO_POR_CABEZA * JUGADORES QUE PARTICIPARON
 
Public DeathMatch   As tDeath
 
Sub Limpiar()
 
' @ Limpia los datos anteriores.
 
Dim DumpPos     As WorldPos
Dim loopx       As Long
Dim LoopY       As Long
Dim esSalida    As Boolean
 
With DeathMatch
     .Cuenta = 0
     .Cupos = 0
     .Ingresaron = 0
     .Activo = False
     .CaenObjs = False
     
     'NPC Banquero invocado?
     If .BanqueroIndex <> 0 Then
        'Nos aseguramos de que esté invocado, con esto : P
        If Npclist(.BanqueroIndex).Numero <> 0 Then
           'Lo borramos.
           QuitarNPC .BanqueroIndex
        End If
     End If
     
     .BanqueroIndex = 0
     
     'Limpio el tipo de ganador.
     With .Ganador
          .Userindex = 0
          .LastPosition = DumpPos
          .Esperando = 0
     End With
     
     'Limpia los objetos que quedaron tira2.
     For loopx = 1 To 100
         For LoopY = 1 To 100
             With MapData(ARENA_MAP, loopx, LoopY)
                  'Hay objeto?
                  If .ObjInfo.objindex <> 0 Then
                     'Flag por si hay salida.
                     esSalida = (.TileExit.Map <> 0)
                     'No es del mapa.
                     If Not ItemNoEsDeMapa(.ObjInfo.objindex, esSalida) Then
                        'Erase :P
                        Call EraseObj(.ObjInfo.Amount, ARENA_MAP, loopx, LoopY)
                     End If
                  End If
             End With
         Next LoopY
     Next loopx
     
End With
 
End Sub
 
Sub Cancelar(ByRef CancelatedBy As String)
 
' @ Cancela el death.
 
Dim loopx   As Long
Dim UIndex  As Integer
Dim UPos    As WorldPos
 
'Aviso.
SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Deathmatch> Cancelado por : " & CancelatedBy & ".", FontTypeNames.FONTTYPE_ADMIN)
 
'Llevo los usuarios que entraron a ulla.
For loopx = 1 To UBound(DeathMatch.Usuarios())
    UIndex = DeathMatch.Usuarios(loopx).Userindex
    'Hay usuario?
    If UIndex <> -1 Then
       'Está logeado?
       If UserList(UIndex).ConnID <> -1 Then
          'Está en death?
          If UserList(UIndex).death Then
             'Telep to anterior posición.
             Call AnteriorPos(UIndex, UPos)
             WarpUserChar UIndex, UPos.Map, UPos.X, UPos.Y, True
             'Reset el flag.
             UserList(UIndex).death = False
          End If
       End If
    End If
Next loopx
 
'Limpia el tipo
Limpiar
 
End Sub
 
Sub ActivarNuevo(ByRef OrganizatedBy As String, ByVal Cupos As Byte, ByVal CaenObjetos As Boolean)
 
' @ Crea nuevo deathmatch.
 
Dim loopx   As Long
 
'Limpia el tipo.
Limpiar
 
'Llena los datos nuevos.
With DeathMatch
     .Cupos = Cupos
     .Activo = True
     .CaenObjs = CaenObjetos
     
     'Redim array.
     ReDim .Usuarios(1 To Cupos) As DeathUser
     
     'Lleno el array con -1s
     For loopx = 1 To Cupos
         .Usuarios(loopx).Userindex = -1
     Next loopx
     
     'Avisa al mundo.
     SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Deathmatch > Organizado : " & OrganizatedBy & " " & Cupos & " Cupos! para entrar /DEATH" & IIf(.CaenObjs, ", Cae el inventario!", "."), FontTypeNames.FONTTYPE_ADMIN)
     SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Deathmatch > Quedan 3 minutos antes del auto-cancelamiento si no se llena el cupo, el precio de inscripción es de 50.000 monedas de oro.", FontTypeNames.FONTTYPE_ADMIN)
     
     'Set el tiempo de auto-cancelación.
     .AutoCancelTime = TIEMPO_AUTOCANCEL
End With
 
End Sub
 
Sub Ingresar(ByVal Userindex As Integer)
 
' @ Usuario ingresa al death.
 
Dim LibreSlot   As Byte
Dim SumarCount  As Boolean
 
LibreSlot = ProximoSlot(SumarCount)
 
'No hay slot.
If Not LibreSlot <> 0 Then Exit Sub
 
With DeathMatch
     'Hay que sumar?
     If SumarCount Then .Ingresaron = .Ingresaron + 1
     
     'Lleno el usuario.
     .Usuarios(LibreSlot).LastPosition = UserList(Userindex).Pos
     .Usuarios(LibreSlot).Userindex = Userindex
     
     'Llevo a la arena.
     WarpUserChar Userindex, ARENA_MAP, ARENA_X, ARENA_Y, True
     
     'Aviso..
     WriteConsoleMsg Userindex, "Has ingresado al deathmatch, eres el participante nº" & LibreSlot & ".", FontTypeNames.FONTTYPE_ADMIN
     
     UserList(Userindex).death = True
     
     'Lleno el cupo?
     If .Ingresaron >= .Cupos Then
         'Quito el tiempo de auto-cancelación
         .AutoCancelTime = 0
         'Aviso que llenó el cupo
         SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("DeathMatch> El cupo ha sido completado!", FontTypeNames.FONTTYPE_ADMIN)
         'Doy inicio
         Iniciar
     End If
     
End With
 
End Sub
 
Sub Cuenta()
 
' @ Cuenta regresiva y auto-cancel acá.
 
Dim PacketToSend    As String
Dim CanSendPackage  As Boolean
 
With DeathMatch
     
    'Espera el ganador?
    If .Ganador.Userindex <> 0 Then
       'Tiempo de espera
       If .Ganador.Esperando <> 0 Then
          'resta.
          .Ganador.Esperando = .Ganador.Esperando - 1
          'Llego al fin el tiempo.
          If Not .Ganador.Esperando <> 0 Then
             'Telep to anterior pos.
             WarpUserChar .Ganador.Userindex, .Ganador.LastPosition.Map, .Ganador.LastPosition.X, .Ganador.LastPosition.Y, True
             'Aviso al usuario.
             WriteConsoleMsg .Ganador.Userindex, "El tiempo ha llegado a su fin, fuiste devuelto a tu posición anterior", FontTypeNames.FONTTYPE_ADMIN
             'Limpiar.
             Limpiar
          End If
        End If
    End If
   
    'Hay cuenta?
    If .Cuenta <> 0 Then
        'Resta el tiempo.
        .Cuenta = .Cuenta - 1
       
        If .Cuenta > 1 Then
            SendData SendTarget.toMap, ARENA_MAP, PrepareMessageConsoleMsg("El death iniciará en " & .Cuenta & " segundos.", FontTypeNames.FONTTYPE_ADMIN)
        ElseIf .Cuenta = 1 Then
            SendData SendTarget.toMap, ARENA_MAP, PrepareMessageConsoleMsg("El death iniciará en 1 segundo!", FontTypeNames.FONTTYPE_ADMIN)
        ElseIf .Cuenta <= 0 Then
            SendData SendTarget.toMap, ARENA_MAP, PrepareMessageConsoleMsg("El death inició! ¡Feliz matanza!", FontTypeNames.FONTTYPE_ADMIN)
            MapInfo(ARENA_MAP).Pk = True
        End If
    End If
   
    'Tiempo de auto-cancelamiento?
    If .AutoCancelTime <> 0 Then
       'Resto el contador
       If .AutoCancelTime <> 0 Then
           .AutoCancelTime = .AutoCancelTime - 1
       End If
             
       'Avisa cada 30 segundos.
       Select Case .AutoCancelTime
              Case 150      'Quedan 2:30.
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsg("Deathmatch> Se cancelará en 2:30 minutos", FontTypeNames.FONTTYPE_ADMIN)
              Case 120
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsg("Deathmatch> Se cancelará en 2 minutos", FontTypeNames.FONTTYPE_ADMIN)
              Case 90
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsg("Deathmatch> Se cancelará en 1:30 minutos", FontTypeNames.FONTTYPE_ADMIN)
              Case 60
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsg("Deathmatch> Se cancelará en 1 minuto", FontTypeNames.FONTTYPE_ADMIN)
              Case 30
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsg("Deathmatch> Se cancelará en 30 segundos", FontTypeNames.FONTTYPE_ADMIN)
              'Avisa a los 15
              Case 15
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsg("Deathmatch> Se cancelará en 15 segundos", FontTypeNames.FONTTYPE_ADMIN)
              'Avisa a los 10
              Case 10
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsg("Deathmatch> Se cancelará en 10 segundos", FontTypeNames.FONTTYPE_ADMIN)
              'Avisa a los 5
              Case 5
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsg("Deathmatch> Se cancelará en 5 segundos", FontTypeNames.FONTTYPE_ADMIN)
              'Avisa a los 3,2,1.
              Case 1, 2, 3
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsg("Deathmatch> Se cancelará en " & .AutoCancelTime & " segundo/s", FontTypeNames.FONTTYPE_ADMIN)
              Case 0
                   CanSendPackage = False
                   Call Cancelar("Falta de participantes.")
       End Select
       
       'Hay que enviar el mensaje?
       If CanSendPackage Then
          'Envia
          SendData SendTarget.ToAll, 0, PacketToSend
          'Reset el flag.
          CanSendPackage = False
       End If
       
    End If
   
End With
 
End Sub
 
Sub Iniciar()
 
' @ Inicia el evento.
 
Dim loopx   As Long
 
With DeathMatch
     
     'Set la cuenta.
     .Cuenta = CUENTA_NUM
     
     'Aviso a los usuarios.
     For loopx = 1 To UBound(.Usuarios())
         'Hay usuario?
         If .Usuarios(loopx).Userindex <> -1 Then
            'Está logeado?
            If UserList(.Usuarios(loopx).Userindex).ConnID <> -1 Then
               WriteConsoleMsg .Usuarios(loopx).Userindex, "Llenó el cupo! El Deathmatch iniciará en " & .Cuenta & " segundos!.", FontTypeNames.FONTTYPE_ADMIN
            Else    'No loged, limpio el tipo
               .Usuarios(loopx).Userindex = -1
            End If
         End If
     Next loopx
   
    'Por default el mapa es seguro..
    MapInfo(ARENA_MAP).Pk = False
     
End With
 
End Sub
 
Sub MuereUser(ByVal muertoIndex As Integer)
 
' @ Muere usuario en dm.
 
Dim MuertoPos       As WorldPos
Dim QuedanEnDeath   As Byte
 
'Obtengo la anterior posición del usuario
Call AnteriorPos(muertoIndex, MuertoPos)
 
'Si caen objetos pincho al usuario.
If DeathMatch.CaenObjs Then
   TirarTodosLosItems muertoIndex
End If
 
'Revivir usuario
RevivirUsuario muertoIndex
 
'Llenar vida.
UserList(muertoIndex).Stats.MinHp = UserList(muertoIndex).Stats.MaxHp
 
'Actualizar hp.
WriteUpdateHP muertoIndex
 
 
'Reset el flag.
UserList(muertoIndex).death = False
 
'Telep anterior pos.
WarpUserChar muertoIndex, MuertoPos.Map, MuertoPos.X, MuertoPos.Y, True
 
'Aviso al usuario
WriteConsoleMsg muertoIndex, "Has caido en el deathMatch, has sido revivido y llevado a tu posición anterior (Mapa : " & MapInfo(MuertoPos.Map).Name & ")", FontTypeNames.FONTTYPE_ADMIN
 
'Aviso al mapa.
SendData SendTarget.toMap, ARENA_MAP, PrepareMessageConsoleMsg(UserList(muertoIndex).Name & " ha sido derrotado.", FontTypeNames.FONTTYPE_ADMIN)
 
'Obtengo los usuarios que quedan..
QuedanEnDeath = QuedanVivos()
 
'Queda 1?
If Not QuedanEnDeath <> 1 Then
   'Ganó ese usuario!
   Terminar
End If
   
End Sub
 
Sub Terminar()
 
' @ Termina el death y gana un usuario.
 
Dim winnerIndex As Integer
Dim GoldPremio  As Long
 
winnerIndex = GanadorIndex
 
'No hay ganador!! TRAGEDIAA XDD
If Not winnerIndex <> -1 Then
   SendData SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("TRAGEDIA EN DEATHMATCHS!! WINNERINDEX = -1!!!!", FontTypeNames.FONTTYPE_GUILD)
   Limpiar
   Exit Sub
End If
 
'Hay ganador, le doi el premio..
GoldPremio = (PREMIO_POR_CABEZA * DeathMatch.Cupos)
UserList(winnerIndex).Stats.GLD = UserList(winnerIndex).Stats.GLD + GoldPremio
 
'Actualizo el oro
WriteUpdateGold winnerIndex
 
With UserList(winnerIndex)
    'Aviso al mundo.
    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Deathmatch> " & .Name & " - " & ListaClases(.clase) & " " & ListaRazas(.raza) & " Nivel " & .Stats.ELV & " Ganó " & Format$(GoldPremio, "#,###") & " monedas de oro, " & IIf(DeathMatch.CaenObjs, "y los objetos recaudados", "") & " por salir primero en el evento.", FontTypeNames.FONTTYPE_ADMIN)
End With
 
'Ganador a su anterior posición..
Dim ToPosition  As WorldPos
Call AnteriorPos(winnerIndex, ToPosition)
 UserList(winnerIndex).death = False
'Si era por objetos no lo llevo a la ciudad.
If DeathMatch.CaenObjs Then
   'Set los flags.
   DeathMatch.Ganador.LastPosition = ToPosition
   DeathMatch.Ganador.Userindex = winnerIndex
   DeathMatch.Ganador.Esperando = TIEMPO_PARAVOLVER
   'Le aviso al pibe que va a tener tiempo de lukear y depositar.
   WriteConsoleMsg winnerIndex, "Tienes " & (TIEMPO_PARAVOLVER / 60) & " minutos para agarrar los objetos que desees.", FontTypeNames.FONTTYPE_ADMIN
   WriteConsoleMsg winnerIndex, "Hay un banquero rondando este mapa, buscalo si lo necesitas.", FontTypeNames.FONTTYPE_ADMIN
   'Invoco un banquero y guardo su index : P
   DeathMatch.BanqueroIndex = SpawnNpc(24, GetBanqueroPos, True, False)
   Exit Sub
End If
 
'Warp.
WarpUserChar winnerIndex, ToPosition.Map, ToPosition.X, ToPosition.Y, True
 
SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Deathmatch> Finalizado", FontTypeNames.FONTTYPE_ADMIN)
Limpiar
 
End Sub
 
Sub AnteriorPos(ByVal Userindex As Integer, ByRef MuertoPosition As WorldPos)
 
' @ Devuelve la posición anterior del usuario.
 
Dim loopx   As Long
 
For loopx = 1 To UBound(DeathMatch.Usuarios())
    If DeathMatch.Usuarios(loopx).Userindex = Userindex Then
       MuertoPosition = DeathMatch.Usuarios(loopx).LastPosition
       Exit Sub
    End If
Next loopx
 
'Posición de ulla u.u
MuertoPosition = Ullathorpe
 
End Sub
 
Function AprobarIngreso(ByVal Userindex As Integer, ByRef MensajeError As String) As Boolean
 
' @ Checks si puede ingresar al death.
 
Dim DumpBoolean As Boolean
 
AprobarIngreso = False
 
'No hay death.
If Not DeathMatch.Activo Then
   MensajeError = "No hay deathmatch en curso."
   Exit Function
End If
 
'No hay cupos.
If Not ProximoSlot(DumpBoolean) <> 0 Then
   MensajeError = "Hay un deathmatch, pero las inscripciones están cerradas"
   Exit Function
End If
 
      If Not UserList(Userindex).Pos.Map = 1 Then
    'Call WriteConsoleMsg(UserIndex, "¡¡No puedes ingresar si no estas en Ullathorpe!!.", FontTypeNames.FONTTYPE_INFO)
    WriteConsoleMsg Userindex, "No puedes ingresar al DeathMatch si no estás en la Ciudad de Ullathorpe.", FontTypeNames.FONTTYPE_INFO
    Exit Function
     End If
 
'Ya inscripto?
If YaInscripto(Userindex) Then
   MensajeError = "Ya estás en el death!"
   Exit Function
End If
 
'Está muerto
If UserList(Userindex).flags.Muerto <> 0 Then
   MensajeError = "Muerto no puedes ingresar a un deathmatch, lo siento.."
   Exit Function
End If
 
'Está preso
If UserList(Userindex).Counters.Pena <> 0 Then
   MensajeError = "No puedes ingresar si estás preso."
   Exit Function
End If

If UserList(Userindex).Stats.GLD < 50000 Then
                    Call WriteConsoleMsg(Userindex, "No tienes suficientes monedas de oro, necesitas 50.000 monedas para ingresar al DeathMatch", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
               
                UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD - 50000
WriteUpdateUserStats (Userindex)
 
If UserList(Userindex).Pos.Map = 66 Then
            MensajeError = "No puedes ingresar si estás preso."
            Exit Function
            End If
 
AprobarIngreso = True
End Function
 
Function ProximoSlot(ByRef Sumar As Boolean) As Byte
 
' @ Posición para un usuario.
 
Dim loopx   As Long
 
Sumar = False
 
For loopx = 1 To UBound(DeathMatch.Usuarios())
    'No hay usuario.
    If Not DeathMatch.Usuarios(loopx).Userindex <> -1 Then
       'Slot encontrado.
       ProximoSlot = loopx
       'Hay que sumar el contador?
       If DeathMatch.Ingresaron < ProximoSlot Then Sumar = True
       Exit Function
    End If
Next loopx
 
ProximoSlot = 0
 
End Function
 
Function QuedanVivos() As Byte
 
' @ Devuelve la cantidad de usuarios vivos que quedan.
 
Dim loopx   As Long
Dim Counter As Byte
 
For loopx = 1 To UBound(DeathMatch.Usuarios())
    'Mientras halla usuario.
    If DeathMatch.Usuarios(loopx).Userindex <> -1 Then
       'Mientras esté logeado
       If UserList(DeathMatch.Usuarios(loopx).Userindex).ConnID <> -1 Then
          'Mientras esté en el mapa de death
          If Not UserList(DeathMatch.Usuarios(loopx).Userindex).Pos.Map <> ARENA_MAP Then
             'Sumo contador.
             Counter = Counter + 1
           End If
        End If
    End If
Next loopx
 
QuedanVivos = Counter
 
End Function
 
Function GanadorIndex() As Integer
 
' @ Busca el ganador..
 
Dim loopx   As Long
 
For loopx = 1 To UBound(DeathMatch.Usuarios())
    If DeathMatch.Usuarios(loopx).Userindex <> -1 Then
   
       If UserList(DeathMatch.Usuarios(loopx).Userindex).ConnID <> -1 Then
          If Not UserList(DeathMatch.Usuarios(loopx).Userindex).Pos.Map <> ARENA_MAP Then
         
             If Not UserList(DeathMatch.Usuarios(loopx).Userindex).flags.Muerto <> 0 Then
                GanadorIndex = DeathMatch.Usuarios(loopx).Userindex
                Exit Function
             End If
             
           End If
        End If
       
    End If
Next loopx
 
'No hay ganador! WTF!!!
GanadorIndex = -1
 
End Function
 
Function YaInscripto(ByVal Userindex As Integer) As Boolean
 
' @ Devuelve si ya está inscripto.
 
Dim loopx   As Long
 
For loopx = 1 To UBound(DeathMatch.Usuarios())
    If DeathMatch.Usuarios(loopx).Userindex = Userindex Then
       YaInscripto = True
       Exit Function
    End If
Next loopx
 
YaInscripto = False
 
End Function
 
Function GetBanqueroPos() As WorldPos
 
' @ Devuelve una posición para el banquero.
 
'No hay objeto.
If Not MapData(ARENA_MAP, BANCO_X, BANCO_Y).ObjInfo.objindex <> 0 Then
   'Si no hay usuario me quedo con esta pos.
   If Not MapData(ARENA_MAP, BANCO_X, BANCO_Y).Userindex <> 0 Then
      GetBanqueroPos.Map = ARENA_MAP
      GetBanqueroPos.X = BANCO_X
      GetBanqueroPos.Y = BANCO_Y
      Exit Function
   End If
End If
 
'Si no estaba libre el anterior tile, busco uno en un radio de 5 tiles.
Dim loopx   As Long
Dim LoopY   As Long
 
For loopx = (BANCO_X - 5) To (BANCO_X + 5)
    For LoopY = (BANCO_Y - 5) To (BANCO_Y + 5)
        With MapData(ARENA_MAP, loopx, LoopY)
             'No hay un objeto..
             If Not .ObjInfo.objindex <> 0 Then
                'No hay usuario.
                If Not .Userindex <> 0 Then
                   'Nos quedamos acá.
                   GetBanqueroPos.Map = ARENA_MAP
                   GetBanqueroPos.X = loopx
                   GetBanqueroPos.Y = LoopY
                   Exit Function
                End If
             End If
        End With
    Next LoopY
Next loopx
 
'Poco probable, pero bueno, si no hay una posición libre
'Devolvemos la posición "ORIGINAL", lo peor que puede pasar
'Es pisar un objeto : P
GetBanqueroPos.Map = ARENA_MAP
GetBanqueroPos.X = BANCO_X
GetBanqueroPos.Y = BANCO_Y
 
End Function

