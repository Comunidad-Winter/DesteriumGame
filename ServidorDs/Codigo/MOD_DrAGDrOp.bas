Attribute VB_Name = "MOD_DrAGDrOp"

Option Explicit
 
Sub DragToUser(ByVal UserIndex As Integer, ByVal tIndex As Integer, ByVal Slot As Byte, ByVal Amount As Integer, ByVal ACT As Boolean)

      ' @ Author : maTih.-
      '            Drag un slot a un usuario.

      Dim tobj    As Obj
      Dim tString As String
      Dim Espacio As Boolean
      Dim ObjIndex As Integer
      Dim errorfound As String

      'No quier el puto item

   On Error GoTo DragToUser_Error

10        If Not CanDragObj(UserList(UserIndex).Invent.Object(Slot).ObjIndex, errorfound) Then
20            WriteConsoleMsg UserIndex, errorfound, FontTypeNames.FONTTYPE_INFO

30            Exit Sub
40        End If


50        If UserList(UserIndex).flags.Comerciando Then Exit Sub

60        If UserList(tIndex).ACT = True Then
70            WriteConsoleMsg UserIndex, "El usuario no quiere tus items!", FontTypeNames.FONTTYPE_INFO
80            Exit Sub
90        End If

100       If UserList(UserIndex).flags.Muerto = 1 Then
110           WriteConsoleMsg UserIndex, "¡Estás Muerto!", FontTypeNames.FONTTYPE_INFO
120           Exit Sub
130       End If

140       If UserList(tIndex).flags.Muerto = 1 Then
150           WriteConsoleMsg UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO
160           Exit Sub
170       End If

          'Preparo el objeto.
180       tobj.Amount = Amount
190       tobj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
          
200       If Amount < 1 Then
210           Exit Sub
220       End If

230       If Not MeterItemEnInventario(tIndex, tobj) Then
240           WriteConsoleMsg UserIndex, "El usuario no tiene espacio en su inventario.", FontTypeNames.FONTTYPE_INFO
250           Exit Sub
260       End If
          
          'Quito el objeto.
270       QuitarUserInvItem UserIndex, Slot, Amount
          
          'Hago un update de su inventario.
280       UpdateUserInv False, UserIndex, Slot
          
          'Preparo el mensaje para userINdex (quien dragea)
          
290       tString = "Le has arrojado"
          
300       If tobj.Amount <> 1 Then
310          tString = tString & " " & tobj.Amount & " - " & ObjData(tobj.ObjIndex).Name
320       Else
330          tString = tString & " tu " & ObjData(tobj.ObjIndex).Name
340       End If
          
350       tString = tString & " a " & UserList(tIndex).Name
          
          'Envio el mensaje
360       WriteConsoleMsg UserIndex, tString, FontTypeNames.FONTTYPE_INFO
          
          'Preparo el mensaje para el otro usuario (quien recibe)
370       tString = UserList(UserIndex).Name & " te ha arrojado"
          
380       If tobj.Amount <> 1 Then
390          tString = tString & " " & tobj.Amount & " - " & ObjData(tobj.ObjIndex).Name
400       Else
410          tString = tString & " su " & ObjData(tobj.ObjIndex).Name
420       End If
          
          'Envio el mensaje al otro usuario
430       WriteConsoleMsg tIndex, tString, FontTypeNames.FONTTYPE_INFO

   On Error GoTo 0
   Exit Sub

DragToUser_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure DragToUser of Módulo MOD_DrAGDrOp in line " & Erl

End Sub
 
Public Sub DragToNPC(ByVal UserIndex As Integer, _
                     ByVal tNpc As Integer, _
                     ByVal Slot As Byte, _
                     ByVal Amount As Integer)
       
              ' @ Author : maTih.-
              '            Drag un slot a un npc.

   On Error GoTo DragToNPC_Error

10            On Error GoTo Errhandler
       
              Dim TeniaOro As Long
              Dim teniaObj As Integer
              Dim tmpIndex As Integer
       
20            tmpIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
30            TeniaOro = UserList(UserIndex).Stats.Gld
40            teniaObj = UserList(UserIndex).Invent.Object(Slot).Amount
       
              'Es un banquero?
50    If UserList(UserIndex).flags.Comerciando Then Exit Sub


      'If tmpIndex < 1 Then
       'WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
      'End If

      'If Amount < 1 Then
       'WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
      'End If

      'If Amount < tmpIndex Then
       'WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
      'End If
60            If Amount > teniaObj Then
70    WriteConsoleMsg UserIndex, "No tienes esa cantidad", FontTypeNames.FONTTYPE_INFO
80    Exit Sub
90    End If

100           If Npclist(tNpc).NPCtype = eNPCType.Banquero Then
110                   Call UserDejaObj(UserIndex, Slot, Amount)
                      'No tiene más el mismo amount que antes? entonces depositó.

120                   If teniaObj <> UserList(UserIndex).Invent.Object(Slot).Amount Then
130                           WriteConsoleMsg UserIndex, "Has depositado " & Amount & " - " & ObjData(tmpIndex).Name, FontTypeNames.FONTTYPE_INFO
140                           UpdateUserInv False, UserIndex, Slot
150                   End If

                      'Es un npc comerciante?
160           ElseIf Npclist(tNpc).Comercia = 1 Then
                      'El npc compra cualquier tipo de items?

170                   If Not Npclist(tNpc).TipoItems <> eOBJType.otCualquiera Or Npclist(tNpc).TipoItems = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).OBJType Then
180                           Call Comercio(eModoComercio.Venta, UserIndex, tNpc, Slot, Amount)
                              'Ganó oro? si es así es porque lo vendió.

190                           If TeniaOro <> UserList(UserIndex).Stats.Gld Then
200                                   WriteConsoleMsg UserIndex, "Le has vendido al " & Npclist(tNpc).Name & " " & Amount & " - " & ObjData(tmpIndex).Name, FontTypeNames.FONTTYPE_INFO
210                           End If

220                   Else
230                           WriteConsoleMsg UserIndex, "El npc no está interesado en comprar este tipo de objetos.", FontTypeNames.FONTTYPE_INFO
240                   End If
250           End If
       
260           Exit Sub
       
Errhandler:
       

   On Error GoTo 0
   Exit Sub

DragToNPC_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure DragToNPC of Módulo MOD_DrAGDrOp in line " & Erl
 
End Sub
 
Public Sub DragToPos(ByVal UserIndex As Integer, _
                     ByVal X As Byte, _
                     ByVal Y As Byte, _
                     ByVal Slot As Byte, _
                     ByVal Amount As Integer)
       
              ' @ Author : maTih.-
              '            Drag un slot a una posición.
       
              Dim errorfound As String
              Dim tobj       As Obj
              Dim tString    As String
       
              'No puede dragear en esa pos?

   On Error GoTo DragToPos_Error

10    If UserList(UserIndex).Pos.map = 200 Or UserList(UserIndex).Pos.map = 192 Or UserList(UserIndex).Pos.map = 195 Or UserList(UserIndex).Pos.map = 191 Or UserList(UserIndex).Pos.map = 176 Then Exit Sub

20    If UserList(UserIndex).flags.Muerto = 1 Then
30        WriteConsoleMsg UserIndex, "¡Estás Muerto!", FontTypeNames.FONTTYPE_INFO
40        Exit Sub
50    End If

      'If UserList(UserIndex).Invent.Object(Slot).ObjIndex < 1 Then
      ' WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
      'End If

      'If Amount < 1 Then
      ' WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
      'End If

      'If Amount < UserList(UserIndex).Invent.Object(Slot).ObjIndex Then
       'WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
      'End If

60             If Not CanDragObj(UserList(UserIndex).Invent.Object(Slot).ObjIndex, errorfound) Then
70                    WriteConsoleMsg UserIndex, errorfound, FontTypeNames.FONTTYPE_INFO

80                    Exit Sub

90            End If

100           If Not CanDragToPos(UserList(UserIndex).Pos.map, X, Y, errorfound) Then
110                   WriteConsoleMsg UserIndex, errorfound, FontTypeNames.FONTTYPE_INFO

120                   Exit Sub

130           End If
       
              'Creo el objeto.
140           tobj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
150           tobj.Amount = Amount
       
              'Agrego el objeto a la posición.
160           MakeObj tobj, UserList(UserIndex).Pos.map, CInt(X), CInt(Y)
       
              'Quito el objeto.
170           QuitarUserInvItem UserIndex, Slot, Amount
       
              'Actualizo el inventario
180           UpdateUserInv False, UserIndex, Slot
       
              'Preparo el mensaje.
190           tString = "¡Lanzas imprecisamente!"
       
              'If tobj.Amount <> 1 Then
                '      tString = tString & tobj.Amount & " - " & ObjData(tobj.ObjIndex).Name
              'Else
                      'tString = tString & "tu " & ObjData(tobj.ObjIndex).Name 'faltaba el tstring &
             ' End If
       
              'ENvio.
200           WriteConsoleMsg UserIndex, tString, FontTypeNames.FONTTYPE_INFO

   On Error GoTo 0
   Exit Sub

DragToPos_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure DragToPos of Módulo MOD_DrAGDrOp in line " & Erl
       
End Sub
 
Private Function CanDragToPos(ByVal map As Integer, _
                              ByVal X As Byte, _
                              ByVal Y As Byte, _
                              ByRef error As String) As Boolean
       
              ' @ Author : maTih.-
              '            Devuelve si se puede dragear un item a x posición.
       
   On Error GoTo CanDragToPos_Error

10            CanDragToPos = False
       
       

              'Zona segura?

20            If Not MapInfo(map).Pk Then
30                    error = "No está permitido arrojar objetos al suelo en zonas seguras."

40                    Exit Function

50            End If
       
              'Ya hay objeto?

60            If Not MapData(map, X, Y).ObjInfo.ObjIndex = 0 Then
70                    error = "Hay un objeto en esa posición!"

80                    Exit Function

90            End If
       
              'Tile bloqueado?

100           If Not MapData(map, X, Y).Blocked = 0 Then
110                   error = "No puedes arrojar objetos en esa posición"

120                   Exit Function

130           End If
              
140           If HayAgua(map, X, Y) Then
150                   error = "No puedes arrojar objetos al agua"
                      
160                   Exit Function

170           End If

180           CanDragToPos = True

   On Error GoTo 0
   Exit Function

CanDragToPos_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure CanDragToPos of Módulo MOD_DrAGDrOp in line " & Erl
       
End Function
 
Private Function CanDragObj(ByVal ObjIndex As Integer, _
                            ByRef error As String) As Boolean
       
              ' @ Author : maTih.-
              '            Devuelve si un objeto es drageable.
   On Error GoTo CanDragObj_Error

10            CanDragObj = False
       

       
20            If ObjIndex < 1 Or ObjIndex > UBound(ObjData()) Then Exit Function
       
              'Objeto newbie?

      'If ObjIndex < 1 Then
      'error = "No tienes esa cantidad de items."
       'Exit Function
      'End If

30            If ObjData(ObjIndex).Newbie <> 0 Then
40                    error = "No puedes arrojar objetos newbies!"

50                    Exit Function

60            End If
       
70             If ObjData(ObjIndex).VIP <> 0 Then
80                    error = "¡No puedes arrojar objetos tipo Oro, Plata o Bronce!"

90                    Exit Function

100           End If
              
110                   If ObjData(ObjIndex).VIPP <> 0 Then
120                   error = "¡No puedes arrojar objetos tipo Oro, Plata o Bronce!"

130                   Exit Function

140           End If
              
150                           If ObjData(ObjIndex).VIPB <> 0 Then
160                   error = "¡No puedes arrojar objetos tipo Oro, Plata o Bronce!"

170                   Exit Function

180           End If
              
              
190                   If ObjData(ObjIndex).Real <> 0 Then
200                   error = "¡No puedes arrojar tus objetos faccionarios!"

210                   Exit Function

220           End If
              
230                           If ObjData(ObjIndex).Caos <> 0 Then
240                   error = "¡No puedes arrojar tus objetos faccionarios!"

250                   Exit Function

260           End If
              
       
              'Está navgeando?

       
270           CanDragObj = True

   On Error GoTo 0
   Exit Function

CanDragObj_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure CanDragObj of Módulo MOD_DrAGDrOp in line " & Erl
       
End Function

Public Sub HandleDragInventory(ByVal UserIndex As Integer)

              ' @ Author : Amraphen.
              '            Drag&Drop de objetos en el inventario.

              Dim ObjSlot1   As Byte
              Dim ObjSlot2   As Byte

              Dim tmpUserObj As UserOBJ
       
   On Error GoTo HandleDragInventory_Error

10            If UserList(UserIndex).incomingData.length < 3 Then
20                    Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode

30                    Exit Sub

40            End If
       
50            With UserList(UserIndex)
              
                      'Leemos el paquete
60                    Call .incomingData.ReadByte
             
70                    ObjSlot1 = .incomingData.ReadByte
80                    ObjSlot2 = .incomingData.ReadByte
90           If UserList(UserIndex).flags.Comerciando Then Exit Sub
                      'Cambiamos si alguno es un anillo

100                   If .Invent.AnilloEqpSlot = ObjSlot1 Then
110                           .Invent.AnilloEqpSlot = ObjSlot2
120                   ElseIf .Invent.AnilloEqpSlot = ObjSlot2 Then
130                           .Invent.AnilloEqpSlot = ObjSlot1
140                   End If
             
                      'Cambiamos si alguno es un armor

150                   If .Invent.ArmourEqpSlot = ObjSlot1 Then
160                           .Invent.ArmourEqpSlot = ObjSlot2
170                   ElseIf .Invent.ArmourEqpSlot = ObjSlot2 Then
180                           .Invent.ArmourEqpSlot = ObjSlot1
190                   End If
             
                      'Cambiamos si alguno es un barco

200                   If .Invent.BarcoSlot = ObjSlot1 Then
210                           .Invent.BarcoSlot = ObjSlot2
220                   ElseIf .Invent.BarcoSlot = ObjSlot2 Then
230                           .Invent.BarcoSlot = ObjSlot1
240                   End If
             
                      'Cambiamos si alguno es un casco

250                   If .Invent.CascoEqpSlot = ObjSlot1 Then
260                           .Invent.CascoEqpSlot = ObjSlot2
270                   ElseIf .Invent.CascoEqpSlot = ObjSlot2 Then
280                           .Invent.CascoEqpSlot = ObjSlot1
290                   End If
             
                      'Cambiamos si alguno es un escudo

300                   If .Invent.EscudoEqpSlot = ObjSlot1 Then
310                           .Invent.EscudoEqpSlot = ObjSlot2
320                   ElseIf .Invent.EscudoEqpSlot = ObjSlot2 Then
330                           .Invent.EscudoEqpSlot = ObjSlot1
340                   End If
             
                      'Cambiamos si alguno es munición

350                   If .Invent.MunicionEqpSlot = ObjSlot1 Then
360                           .Invent.MunicionEqpSlot = ObjSlot2
370                   ElseIf .Invent.MunicionEqpSlot = ObjSlot2 Then
380                           .Invent.MunicionEqpSlot = ObjSlot1
390                   End If
             
                      'Cambiamos si alguno es un arma

400                   If .Invent.WeaponEqpSlot = ObjSlot1 Then
410                           .Invent.WeaponEqpSlot = ObjSlot2
420                   ElseIf .Invent.WeaponEqpSlot = ObjSlot2 Then
430                           .Invent.WeaponEqpSlot = ObjSlot1
440                   End If
             
                      'Hacemos el intercambio propiamente dicho
450                   tmpUserObj = .Invent.Object(ObjSlot1)
460                   .Invent.Object(ObjSlot1) = .Invent.Object(ObjSlot2)
470                   .Invent.Object(ObjSlot2) = tmpUserObj
       
                      'Actualizamos los 2 slots que cambiamos solamente
480                   Call UpdateUserInv(False, UserIndex, ObjSlot1)
490                   Call UpdateUserInv(False, UserIndex, ObjSlot2)
500           End With

   On Error GoTo 0
   Exit Sub

HandleDragInventory_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleDragInventory of Módulo MOD_DrAGDrOp in line " & Erl

End Sub

Public Sub HandleDragToPos(ByVal UserIndex As Integer)

              ' @ Author : maTih.-
              '            Drag&Drop de objetos en del inventario a una posición.

              Dim X      As Byte
              Dim Y      As Byte
              Dim Slot   As Byte
              Dim Amount As Integer
              Dim tUser  As Integer
              Dim tNpc   As Integer
1            If UserList(UserIndex).incomingData.length < 6 Then
2                    Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode

3                    Exit Sub

4            End If


      'Seria raro que llegue para acá
   On Error GoTo HandleDragToPos_Error

10            Call UserList(UserIndex).incomingData.ReadByte

20            X = UserList(UserIndex).incomingData.ReadByte()
30            Y = UserList(UserIndex).incomingData.ReadByte()
40            Slot = UserList(UserIndex).incomingData.ReadByte()
50            Amount = UserList(UserIndex).incomingData.ReadInteger()

60            tUser = MapData(UserList(UserIndex).Pos.map, X, Y).UserIndex
70            tNpc = MapData(UserList(UserIndex).Pos.map, X, Y).NpcIndex
          
80        If Amount > 0 Then
90            If MapData(UserList(UserIndex).Pos.map, X, Y).NpcIndex <> 0 Then
100                   MOD_DrAGDrOp.DragToNPC UserIndex, tNpc, Slot, Amount
110           Else
              
120                   MOD_DrAGDrOp.DragToPos UserIndex, X, Y, Slot, Amount
130           End If
140       Else


150           Call LogAntiCheat(UserList(UserIndex).Name & " intentó dupear en Drag To Pos")
160           Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " intentó dupear.", FontTypeNames.FONTTYPE_INFO))


170       End If

   On Error GoTo 0
   Exit Sub

HandleDragToPos_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleDragToPos of Módulo MOD_DrAGDrOp in line " & Erl

End Sub






