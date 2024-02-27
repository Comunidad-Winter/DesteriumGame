Attribute VB_Name = "MOD_DrAGDrOp"

Option Explicit
 
Sub DragToUser(ByVal Userindex As Integer, ByVal tIndex As Integer, ByVal Slot As Byte, ByVal Amount As Integer, ByVal ACT As Boolean)

' @ Author : maTih.-
'            Drag un slot a un usuario.

Dim tobj    As Obj
Dim tString As String
Dim Espacio As Boolean
Dim objindex As Integer
Dim errorfound As String

'No quier el puto item

         If Not CanDragObj(UserList(Userindex).Invent.Object(Slot).objindex, errorfound) Then
                WriteConsoleMsg Userindex, errorfound, FontTypeNames.FONTTYPE_INFO

                Exit Sub

        End If


         If Not CanDragObj(UserList(tIndex).Invent.Object(Slot).objindex, errorfound) Then
                WriteConsoleMsg Userindex, errorfound, FontTypeNames.FONTTYPE_INFO

                Exit Sub

        End If

If UserList(Userindex).flags.Comerciando Then Exit Sub

If UserList(tIndex).ACT = True Then
WriteConsoleMsg Userindex, "El usuario no quiere tus items!", FontTypeNames.FONTTYPE_INFO
Exit Sub
End If

If UserList(Userindex).flags.Muerto = 1 Then
    WriteConsoleMsg Userindex, "¡Estás Muerto!", FontTypeNames.FONTTYPE_INFO
    Exit Sub
End If

If UserList(tIndex).flags.Muerto = 1 Then
    WriteConsoleMsg Userindex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO
    Exit Sub
End If

'If tobj.ObjIndex > 0 Then
 'WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
'End If

'If tobj.Amount < 1 Then
 'WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
 
'End If

'If tobj.Amount < tobj.ObjIndex Then
 'WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
 'Exit Sub
'End If

'Preparo el objeto.
tobj.Amount = Amount
tobj.objindex = UserList(Userindex).Invent.Object(Slot).objindex

Espacio = MeterItemEnInventario(tIndex, tobj)

'No tiene espacio.
If Not Espacio Then
   WriteConsoleMsg Userindex, "El usuario no tiene espacio en su inventario.", FontTypeNames.FONTTYPE_INFO
   Exit Sub
End If

If Amount < 1 Then
WriteConsoleMsg Userindex, "No tienes espacio en el inventario.", FontTypeNames.FONTTYPE_INFO
Exit Sub
End If


'Quito el objeto.
QuitarUserInvItem Userindex, Slot, Amount

'Hago un update de su inventario.
UpdateUserInv False, Userindex, Slot

'Preparo el mensaje para userINdex (quien dragea)

tString = "Le has arrojado"

If tobj.Amount <> 1 Then
   tString = tString & " " & tobj.Amount & " - " & ObjData(tobj.objindex).Name
Else
   tString = tString & " tu " & ObjData(tobj.objindex).Name
End If

tString = tString & " a " & UserList(tIndex).Name

'Envio el mensaje
WriteConsoleMsg Userindex, tString, FontTypeNames.FONTTYPE_INFO

'Preparo el mensaje para el otro usuario (quien recibe)
tString = UserList(Userindex).Name & " te ha arrojado"

If tobj.Amount <> 1 Then
   tString = tString & " " & tobj.Amount & " - " & ObjData(tobj.objindex).Name
Else
   tString = tString & " su " & ObjData(tobj.objindex).Name
End If

'Envio el mensaje al otro usuario
WriteConsoleMsg tIndex, tString, FontTypeNames.FONTTYPE_INFO

End Sub
 
Public Sub DragToNPC(ByVal Userindex As Integer, _
                     ByVal tNpc As Integer, _
                     ByVal Slot As Byte, _
                     ByVal Amount As Integer)
 
        ' @ Author : maTih.-
        '            Drag un slot a un npc.

        On Error GoTo Errhandler
 
        Dim TeniaOro As Long
        Dim teniaObj As Integer
        Dim tmpIndex As Integer
 
        tmpIndex = UserList(Userindex).Invent.Object(Slot).objindex
        TeniaOro = UserList(Userindex).Stats.GLD
        teniaObj = UserList(Userindex).Invent.Object(Slot).Amount
 
        'Es un banquero?
If UserList(Userindex).flags.Comerciando Then Exit Sub



'If tmpIndex < 1 Then
 'WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
'End If

'If Amount < 1 Then
 'WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
'End If

'If Amount < tmpIndex Then
 'WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
'End If
                        If Amount > tmpIndex Then
WriteConsoleMsg Userindex, "No tienes esa cantidad", FontTypeNames.FONTTYPE_INFO
Exit Sub
End If

        If Npclist(tNpc).NPCtype = eNPCType.Banquero Then
                Call UserDejaObj(Userindex, Slot, Amount)
                'No tiene más el mismo amount que antes? entonces depositó.

                If teniaObj <> UserList(Userindex).Invent.Object(Slot).Amount Then
                        WriteConsoleMsg Userindex, "Has depositado " & Amount & " - " & ObjData(tmpIndex).Name, FontTypeNames.FONTTYPE_INFO
                        UpdateUserInv False, Userindex, Slot
                End If

                'Es un npc comerciante?
        ElseIf Npclist(tNpc).Comercia = 1 Then
                'El npc compra cualquier tipo de items?

                If Not Npclist(tNpc).TipoItems <> eOBJType.otCualquiera Or Npclist(tNpc).TipoItems = ObjData(UserList(Userindex).Invent.Object(Slot).objindex).OBJType Then
                        Call Comercio(eModoComercio.Venta, Userindex, tNpc, Slot, Amount)
                        'Ganó oro? si es así es porque lo vendió.

                        If TeniaOro <> UserList(Userindex).Stats.GLD Then
                                WriteConsoleMsg Userindex, "Le has vendido al " & Npclist(tNpc).Name & " " & Amount & " - " & ObjData(tmpIndex).Name, FontTypeNames.FONTTYPE_INFO
                        End If

                Else
                        WriteConsoleMsg Userindex, "El npc no está interesado en comprar este tipo de objetos.", FontTypeNames.FONTTYPE_INFO
                End If
        End If
 
        Exit Sub
 
Errhandler:
 
End Sub
 
Public Sub DragToPos(ByVal Userindex As Integer, _
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

If UserList(Userindex).Pos.Map = 200 Or UserList(Userindex).Pos.Map = 192 Or UserList(Userindex).Pos.Map = 195 Or UserList(Userindex).Pos.Map = 191 Or UserList(Userindex).Pos.Map = 176 Then Exit Sub

If UserList(Userindex).flags.Muerto = 1 Then
    WriteConsoleMsg Userindex, "¡Estás Muerto!", FontTypeNames.FONTTYPE_INFO
    Exit Sub
End If

'If UserList(UserIndex).Invent.Object(Slot).ObjIndex < 1 Then
' WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
'End If

'If Amount < 1 Then
' WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
'End If

'If Amount < UserList(UserIndex).Invent.Object(Slot).ObjIndex Then
 'WriteConsoleMsg UserIndex, "No tienes esa cantidad de ítems.", FontTypeNames.FONTTYPE_INFO
'End If

         If Not CanDragObj(UserList(Userindex).Invent.Object(Slot).objindex, errorfound) Then
                WriteConsoleMsg Userindex, errorfound, FontTypeNames.FONTTYPE_INFO

                Exit Sub

        End If

        If Not CanDragToPos(UserList(Userindex).Pos.Map, X, Y, errorfound) Then
                WriteConsoleMsg Userindex, errorfound, FontTypeNames.FONTTYPE_INFO

                Exit Sub

        End If
 
        'Creo el objeto.
        tobj.objindex = UserList(Userindex).Invent.Object(Slot).objindex
        tobj.Amount = Amount
 
        'Agrego el objeto a la posición.
        MakeObj tobj, UserList(Userindex).Pos.Map, CInt(X), CInt(Y)
 
        'Quito el objeto.
        QuitarUserInvItem Userindex, Slot, Amount
 
        'Actualizo el inventario
        UpdateUserInv False, Userindex, Slot
 
        'Preparo el mensaje.
        tString = "¡Lanzas imprecisamente!"
 
        'If tobj.Amount <> 1 Then
          '      tString = tString & tobj.Amount & " - " & ObjData(tobj.ObjIndex).Name
        'Else
                'tString = tString & "tu " & ObjData(tobj.ObjIndex).Name 'faltaba el tstring &
       ' End If
 
        'ENvio.
        WriteConsoleMsg Userindex, tString, FontTypeNames.FONTTYPE_INFO
 
End Sub
 
Private Function CanDragToPos(ByVal Map As Integer, _
                              ByVal X As Byte, _
                              ByVal Y As Byte, _
                              ByRef error As String) As Boolean
 
        ' @ Author : maTih.-
        '            Devuelve si se puede dragear un item a x posición.
 
        CanDragToPos = False
 
 

        'Zona segura?

        If Not MapInfo(Map).Pk Then
                error = "No está permitido arrojar objetos al suelo en zonas seguras."

                Exit Function

        End If
 
        'Ya hay objeto?

        If Not MapData(Map, X, Y).ObjInfo.objindex = 0 Then
                error = "Hay un objeto en esa posición!"

                Exit Function

        End If
 
        'Tile bloqueado?

        If Not MapData(Map, X, Y).Blocked = 0 Then
                error = "No puedes arrojar objetos en esa posición"

                Exit Function

        End If
        
        If HayAgua(Map, X, Y) Then
                error = "No puedes arrojar objetos al agua"
                
                Exit Function

        End If

        CanDragToPos = True
 
End Function
 
Private Function CanDragObj(ByVal objindex As Integer, _
                            ByRef error As String) As Boolean
 
        ' @ Author : maTih.-
        '            Devuelve si un objeto es drageable.
        CanDragObj = False
 

 
        If objindex < 1 Or objindex > UBound(ObjData()) Then Exit Function
 
        'Objeto newbie?

'If ObjIndex < 1 Then
'error = "No tienes esa cantidad de items."
 'Exit Function
'End If

        If ObjData(objindex).Newbie <> 0 Then
                error = "No puedes arrojar objetos newbies!"

                Exit Function

        End If
 
         If ObjData(objindex).VIP <> 0 Then
                error = "¡No puedes arrojar objetos tipo Oro, Plata o Bronce!"

                Exit Function

        End If
        
                If ObjData(objindex).VIPP <> 0 Then
                error = "¡No puedes arrojar objetos tipo Oro, Plata o Bronce!"

                Exit Function

        End If
        
                        If ObjData(objindex).VIPB <> 0 Then
                error = "¡No puedes arrojar objetos tipo Oro, Plata o Bronce!"

                Exit Function

        End If
        
        
                If ObjData(objindex).Real <> 0 Then
                error = "¡No puedes arrojar tus objetos faccionarios!"

                Exit Function

        End If
        
                        If ObjData(objindex).Caos <> 0 Then
                error = "¡No puedes arrojar tus objetos faccionarios!"

                Exit Function

        End If
        
 
        'Está navgeando?

 
        CanDragObj = True
 
End Function

Public Sub HandleDragInventory(ByVal Userindex As Integer)

        ' @ Author : Amraphen.
        '            Drag&Drop de objetos en el inventario.

        Dim ObjSlot1   As Byte
        Dim ObjSlot2   As Byte

        Dim tmpUserObj As UserOBJ
 
        If UserList(Userindex).incomingData.length < 3 Then
                Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode

                Exit Sub

        End If
 
        With UserList(Userindex)
        
                'Leemos el paquete
                Call .incomingData.ReadByte
       
                ObjSlot1 = .incomingData.ReadByte
                ObjSlot2 = .incomingData.ReadByte
       If UserList(Userindex).flags.Comerciando Then Exit Sub
                'Cambiamos si alguno es un anillo

                If .Invent.AnilloEqpSlot = ObjSlot1 Then
                        .Invent.AnilloEqpSlot = ObjSlot2
                ElseIf .Invent.AnilloEqpSlot = ObjSlot2 Then
                        .Invent.AnilloEqpSlot = ObjSlot1
                End If
       
                'Cambiamos si alguno es un armor

                If .Invent.ArmourEqpSlot = ObjSlot1 Then
                        .Invent.ArmourEqpSlot = ObjSlot2
                ElseIf .Invent.ArmourEqpSlot = ObjSlot2 Then
                        .Invent.ArmourEqpSlot = ObjSlot1
                End If
       
                'Cambiamos si alguno es un barco

                If .Invent.BarcoSlot = ObjSlot1 Then
                        .Invent.BarcoSlot = ObjSlot2
                ElseIf .Invent.BarcoSlot = ObjSlot2 Then
                        .Invent.BarcoSlot = ObjSlot1
                End If
       
                'Cambiamos si alguno es un casco

                If .Invent.CascoEqpSlot = ObjSlot1 Then
                        .Invent.CascoEqpSlot = ObjSlot2
                ElseIf .Invent.CascoEqpSlot = ObjSlot2 Then
                        .Invent.CascoEqpSlot = ObjSlot1
                End If
       
                'Cambiamos si alguno es un escudo

                If .Invent.EscudoEqpSlot = ObjSlot1 Then
                        .Invent.EscudoEqpSlot = ObjSlot2
                ElseIf .Invent.EscudoEqpSlot = ObjSlot2 Then
                        .Invent.EscudoEqpSlot = ObjSlot1
                End If
       
                'Cambiamos si alguno es munición

                If .Invent.MunicionEqpSlot = ObjSlot1 Then
                        .Invent.MunicionEqpSlot = ObjSlot2
                ElseIf .Invent.MunicionEqpSlot = ObjSlot2 Then
                        .Invent.MunicionEqpSlot = ObjSlot1
                End If
       
                'Cambiamos si alguno es un arma

                If .Invent.WeaponEqpSlot = ObjSlot1 Then
                        .Invent.WeaponEqpSlot = ObjSlot2
                ElseIf .Invent.WeaponEqpSlot = ObjSlot2 Then
                        .Invent.WeaponEqpSlot = ObjSlot1
                End If
       
                'Hacemos el intercambio propiamente dicho
                tmpUserObj = .Invent.Object(ObjSlot1)
                .Invent.Object(ObjSlot1) = .Invent.Object(ObjSlot2)
                .Invent.Object(ObjSlot2) = tmpUserObj
 
                'Actualizamos los 2 slots que cambiamos solamente
                Call UpdateUserInv(False, Userindex, ObjSlot1)
                Call UpdateUserInv(False, Userindex, ObjSlot2)
        End With

End Sub

Public Sub HandleDragToPos(ByVal Userindex As Integer)

        ' @ Author : maTih.-
        '            Drag&Drop de objetos en del inventario a una posición.

        Dim X      As Byte
        Dim Y      As Byte
        Dim Slot   As Byte
        Dim Amount As Integer
        Dim tUser  As Integer
        Dim tNpc   As Integer

        Call UserList(Userindex).incomingData.ReadByte

        X = UserList(Userindex).incomingData.ReadByte()
        Y = UserList(Userindex).incomingData.ReadByte()
        Slot = UserList(Userindex).incomingData.ReadByte()
        Amount = UserList(Userindex).incomingData.ReadInteger()

        tUser = MapData(UserList(Userindex).Pos.Map, X, Y).Userindex
        tNpc = MapData(UserList(Userindex).Pos.Map, X, Y).NpcIndex
        
        If MapData(UserList(Userindex).Pos.Map, X, Y).NpcIndex <> 0 Then
                MOD_DrAGDrOp.DragToNPC Userindex, tNpc, Slot, Amount
        Else
        
                MOD_DrAGDrOp.DragToPos Userindex, X, Y, Slot, Amount
        End If

End Sub






