Attribute VB_Name = "InvUsuario"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
' 22/05/2010: Los items newbies ya no son robables.
'***************************************************
 
'17/09/02
'Agregue que la función se asegure que el objeto no es un barco
 
On Error GoTo errhandler
 
    Dim i As Integer
    Dim objindex As Integer
   
    For i = 1 To UserList(UserIndex).CurrentInventorySlots
        objindex = UserList(UserIndex).Invent.Object(i).objindex
        If objindex > 0 Then
            If (ObjData(objindex).OBJType <> eOBJType.otLlaves And _
            ObjData(objindex).OBJType <> eOBJType.otMonturas And _
            ObjData(objindex).OBJType <> eOBJType.otMonturasDraco And _
                ObjData(objindex).OBJType <> eOBJType.otBarcos And _
                Not ItemNewbie(objindex)) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
        End If
    Next i
   
    Exit Function
 
errhandler:
    Call LogError("Error en TieneObjetosRobables. Error: " & Err.Number & " - " & Err.description)
End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal objindex As Integer, Optional ByRef sMotivo As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
'***************************************************

On Error GoTo manejador

    Dim flag As Boolean
    
    'Admins can use ANYTHING!
    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
        If ObjData(objindex).ClaseProhibida(1) <> 0 Then
            Dim i As Integer
            For i = 1 To NUMCLASES
                If ObjData(objindex).ClaseProhibida(i) = UserList(UserIndex).clase Then
                    ClasePuedeUsarItem = False
                    sMotivo = "Tu clase no puede usar este objeto."
                    Exit Function
                End If
            Next i
        End If
    End If
    
    ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim j As Integer

With UserList(UserIndex)
    For j = 1 To UserList(UserIndex).CurrentInventorySlots
        If .Invent.Object(j).objindex > 0 Then
             
             If ObjData(.Invent.Object(j).objindex).Newbie = 1 Then _
                    Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, UserIndex, j)
        
        End If
    Next j
    
    '[Barrin 17-12-03] Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
    'es transportado a su hogar de origen ;)
    If UCase$(MapInfo(.Pos.map).Restringir) = "NEWBIE" Then
        
        Dim DeDonde As WorldPos
        
        Select Case .Hogar
            Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                DeDonde = Lindos
            Case eCiudad.cUllathorpe
                DeDonde = Ullathorpe
            Case eCiudad.cBanderbill
                DeDonde = Banderbill
            Case Else
                DeDonde = Nix
        End Select
        
        Call WarpUserChar(UserIndex, DeDonde.map, DeDonde.X, DeDonde.Y, True)
    
    End If
    '[/Barrin]
End With

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim j As Integer

With UserList(UserIndex)
    For j = 1 To .CurrentInventorySlots
        .Invent.Object(j).objindex = 0
        .Invent.Object(j).Amount = 0
        .Invent.Object(j).Equipped = 0
    Next j
    
    .Invent.NroItems = 0
    
    .Invent.ArmourEqpObjIndex = 0
    .Invent.ArmourEqpSlot = 0
    
    .Invent.WeaponEqpObjIndex = 0
    .Invent.WeaponEqpSlot = 0
    
    .Invent.CascoEqpObjIndex = 0
    .Invent.CascoEqpSlot = 0
    
    .Invent.EscudoEqpObjIndex = 0
    .Invent.EscudoEqpSlot = 0
    
    .Invent.AnilloEqpObjIndex = 0
    .Invent.AnilloEqpSlot = 0
    
    .Invent.MunicionEqpObjIndex = 0
    .Invent.MunicionEqpSlot = 0
    
    .Invent.BarcoObjIndex = 0
    .Invent.BarcoSlot = 0
    
    .Invent.MonturaObjIndex = 0
    .Invent.MonturaSlot = 0
    
    .Invent.MochilaEqpObjIndex = 0
    .Invent.MochilaEqpSlot = 0
End With

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
'***************************************************
On Error GoTo errhandler


If Cantidad > 100000 Then Exit Sub

With UserList(UserIndex)
     If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then
        Exit Sub
    End If

    'SI EL Pjta TIENE ORO LO TIRAMOS
    If (Cantidad > 0) And (Cantidad <= .Stats.GLD) Then
            Dim i As Byte
            Dim MiObj As Obj
            'info debug
            Dim loops As Integer
            
            'Seguridad Alkon (guardo el oro tirado si supera los 50k)
            If Cantidad > 50000 Then
                Dim j As Integer
                Dim k As Integer
                Dim m As Integer
                Dim Cercanos As String
                m = .Pos.map
                For j = .Pos.X - 10 To .Pos.X + 10
                    For k = .Pos.Y - 10 To .Pos.Y + 10
                        If InMapBounds(m, j, k) Then
                            If MapData(m, j, k).UserIndex > 0 Then
                                Cercanos = Cercanos & UserList(MapData(m, j, k).UserIndex).name & ","
                            End If
                        End If
                    Next k
                Next j
                Call LogDesarrollo(.name & " tira oro. Cercanos: " & Cercanos)
            End If
            '/Seguridad
            Dim Extra As Long
            Dim TeniaOro As Long
            TeniaOro = .Stats.GLD
            If Cantidad > 500000 Then 'Para evitar explotar demasiado
                Extra = Cantidad - 500000
                Cantidad = 500000
            End If
            
            Do While (Cantidad > 0)
                
                If Cantidad > MAX_INVENTORY_OBJS And .Stats.GLD > MAX_INVENTORY_OBJS Then
                    MiObj.Amount = MAX_INVENTORY_OBJS
                    Cantidad = Cantidad - MiObj.Amount
                Else
                    MiObj.Amount = Cantidad
                    Cantidad = Cantidad - MiObj.Amount
                End If
    
                MiObj.objindex = iORO
                
                If EsGM(UserIndex) Then Call LogGM(.name, "Tiró cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.objindex).name)
                Dim AuxPos As WorldPos
                
                If .clase = eClass.Pirat And .Invent.BarcoObjIndex = 476 Then
                    AuxPos = TirarItemAlPiso(.Pos, MiObj, False)
                    If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                        .Stats.GLD = .Stats.GLD - MiObj.Amount
                    End If
                Else
                    AuxPos = TirarItemAlPiso(.Pos, MiObj, True)
                    If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                        .Stats.GLD = .Stats.GLD - MiObj.Amount
                    End If
                End If
                
                'info debug
                loops = loops + 1
                If loops > 100 Then
                    LogError ("Error en tiraroro")
                    Exit Sub
                End If
                
            Loop
            If TeniaOro = .Stats.GLD Then Extra = 0
            If Extra > 0 Then
                .Stats.GLD = .Stats.GLD - Extra
            End If
        
    End If
End With

Exit Sub

errhandler:
    Call LogError("Error en TirarOro. Error " & Err.Number & " : " & Err.description)
End Sub
Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    If Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots Then Exit Sub
    
    With UserList(UserIndex).Invent.Object(Slot)
        If .Amount <= Cantidad And .Equipped = 1 Then
            Call Desequipar(UserIndex, Slot)
        End If
        
        'Quita un objeto
        .Amount = .Amount - Cantidad
        '¿Quedan mas?
        If .Amount <= 0 Then
            UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
            .objindex = 0
            .Amount = 0
        End If
    End With

Exit Sub

errhandler:
    Call LogError("Error en QuitarUserInvItem. Error " & Err.Number & " : " & Err.description)
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

Dim NullObj As UserOBJ
Dim LoopC As Long

With UserList(UserIndex)
    'Actualiza un solo slot
    If Not UpdateAll Then
    
        'Actualiza el inventario
        If .Invent.Object(Slot).objindex > 0 Then
            Call ChangeUserInv(UserIndex, Slot, .Invent.Object(Slot))
        Else
            Call ChangeUserInv(UserIndex, Slot, NullObj)
        End If
    
    Else
    
    'Actualiza todos los slots
        For LoopC = 1 To .CurrentInventorySlots
            'Actualiza el inventario
            If .Invent.Object(LoopC).objindex > 0 Then
                Call ChangeUserInv(UserIndex, LoopC, .Invent.Object(LoopC))
            Else
                Call ChangeUserInv(UserIndex, LoopC, NullObj)
            End If
        Next LoopC
    End If
    
    Exit Sub
End With

errhandler:
    Call LogError("Error en UpdateUserInv. Error " & Err.Number & " : " & Err.description)

End Sub
Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 11/5/2010
'11/5/2010 - ZaMa: Arreglo bug que permitia apilar mas de 10k de items.
'***************************************************

Dim DropObj As Obj
Dim MapObj As Obj

With UserList(UserIndex)

    If Num > 0 Then
        
        DropObj.objindex = .Invent.Object(Slot).objindex
        
        If (ItemNewbie(DropObj.objindex) And (.flags.Privilegios And PlayerType.User)) And .flags.Muerto <> 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes tirar objetos newbie.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
                  If (ItemFaccionario(DropObj.objindex) And (.flags.Privilegios And PlayerType.User)) Then
            Call WriteConsoleMsg(UserIndex, "¡¡No puedes tirar tu armadura faccionaria!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
                         If (ItemVIP(DropObj.objindex) And (.flags.Privilegios And PlayerType.User)) Then
            Call WriteConsoleMsg(UserIndex, "Por seguridad no puedes arrojar tus objetos Oro, Plata o Bronce.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
                         If (ItemVIPB(DropObj.objindex) And (.flags.Privilegios And PlayerType.User)) Then
            Call WriteConsoleMsg(UserIndex, "Por seguridad no puedes arrojar tus objetos Oro, Plata o Bronce.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
                                 If (ItemVIPP(DropObj.objindex) And (.flags.Privilegios And PlayerType.User)) Then
            Call WriteConsoleMsg(UserIndex, "Por seguridad no puedes arrojar tus objetos Oro, Plata o Bronce.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        DropObj.Amount = MinimoInt(Num, .Invent.Object(Slot).Amount)

        'Check objeto en el suelo
        MapObj.objindex = MapData(.Pos.map, X, Y).ObjInfo.objindex
        MapObj.Amount = MapData(.Pos.map, X, Y).ObjInfo.Amount
        
        If MapObj.objindex = 0 Or MapObj.objindex = DropObj.objindex Then
        
            If MapObj.Amount = MAX_INVENTORY_OBJS Then
                Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
                        If ObjData(DropObj.objindex).Caos = 1 Or ObjData(DropObj.objindex).Real = 1 Then
            WriteConsoleMsg UserIndex, "¡¡ATENCIÓN!! ¡ACABAS DE TIRAR TU ARMADURA FACCIONARIA!", FontTypeNames.FONTTYPE_GUILD
            End If
            
             If ObjData(DropObj.objindex).Premium = 1 And (.flags.Privilegios = PlayerType.User) Then
        WriteConsoleMsg UserIndex, "No puedes tirar items PREMIUM!", FontTypeNames.FONTTYPE_INFO
        Exit Sub
        End If
            
            If DropObj.Amount + MapObj.Amount > MAX_INVENTORY_OBJS Then
                DropObj.Amount = MAX_INVENTORY_OBJS - MapObj.Amount
            End If
            
   
            Call MakeObj(DropObj, map, X, Y)

            Call QuitarUserInvItem(UserIndex, Slot, DropObj.Amount)
            Call UpdateUserInv(False, UserIndex, Slot)
            
            If ObjData(DropObj.objindex).OBJType = eOBJType.otBarcos Then
                Call WriteConsoleMsg(UserIndex, "¡¡ATENCIÓN!! ¡ACABAS DE TIRAR TU BARCA!", FontTypeNames.FONTTYPE_GUILD)
            End If
            
            If ObjData(DropObj.objindex).OBJType = eOBJType.otMonturas Then
            WriteConsoleMsg UserIndex, "¡¡ATENCIÓN!! ¡ACABAS DE TIRAR TU MONTURA!", FontTypeNames.FONTTYPE_GUILD
            End If
            
                         If ObjData(DropObj.objindex).OBJType = eOBJType.otMonturasDraco Then
                Call WriteConsoleMsg(UserIndex, "¡¡ATENCIÓN!! ¡ACABAS DE TIRAR TU MONTURA!", FontTypeNames.FONTTYPE_TALK)
            End If
            
            If ObjData(DropObj.objindex).OBJType = eOBJType.otLunar Then
Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha tirado una Gema Lunar. Se encuentra en el mapa " & UserList(UserIndex).Pos.map & ", " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_GUILD))
End If

            If ObjData(DropObj.objindex).OBJType = eOBJType.otvioleta Then
Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha tirado una Gema Violeta. Se encuentra en el mapa " & UserList(UserIndex).Pos.map & ", " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_GUILD))
End If
     
                 If ObjData(DropObj.objindex).OBJType = eOBJType.otAzul Then
Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha tirado una Gema Azul. Se encuentra en el mapa " & UserList(UserIndex).Pos.map & ", " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_GUILD))
End If

               If ObjData(DropObj.objindex).OBJType = eOBJType.otroja Then
Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha tirado una Gema Roja. Se encuentra en el mapa " & UserList(UserIndex).Pos.map & ", " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_GUILD))
End If

               If ObjData(DropObj.objindex).OBJType = eOBJType.otverde Then
Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha tirado una Gema Verde. Se encuentra en el mapa " & UserList(UserIndex).Pos.map & ", " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_GUILD))
End If

               If ObjData(DropObj.objindex).OBJType = eOBJType.otLila Then
Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha tirado una Gema Lila. Se encuentra en el mapa " & UserList(UserIndex).Pos.map & ", " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_GUILD))
End If
            
               If ObjData(DropObj.objindex).OBJType = eOBJType.otNaranja Then
Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha tirado una Gema Naranja. Se encuentra en el mapa " & UserList(UserIndex).Pos.map & ", " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_GUILD))
End If

               If ObjData(DropObj.objindex).OBJType = eOBJType.otCeleste Then
Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha tirado una Gema Celeste. Se encuentra en el mapa " & UserList(UserIndex).Pos.map & ", " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_GUILD))
End If

            
            If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.name, "Tiró cantidad:" & Num & " Objeto:" & ObjData(DropObj.objindex).name)
            
            'Log de Objetos que se tiran al piso. Pablo (ToxicWaste) 07/09/07
            'Es un Objeto que tenemos que loguear?
            If ObjData(DropObj.objindex).LOG = 1 Then
                Call LogDesarrollo(.name & " tiró al piso " & DropObj.Amount & " " & ObjData(DropObj.objindex).name & " Mapa: " & map & " X: " & X & " Y: " & Y)
            ElseIf DropObj.Amount > 5000 Then 'Es mucha cantidad? > Subí a 5000 el minimo porque si no se llenaba el log de cosas al pedo. (NicoNZ)
                'Si no es de los prohibidos de loguear, lo logueamos.
                If ObjData(DropObj.objindex).NoLog <> 1 Then
                    Call LogDesarrollo(.name & " tiró al piso " & DropObj.Amount & " " & ObjData(DropObj.objindex).name & " Mapa: " & map & " X: " & X & " Y: " & Y)
                End If
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
                       
End With


Call ActualizarAuras(UserIndex)

End Sub

Sub EraseObj(ByVal Num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

MapData(map, X, Y).ObjInfo.Amount = MapData(map, X, Y).ObjInfo.Amount - Num

If MapData(map, X, Y).ObjInfo.Amount <= 0 Then
    MapData(map, X, Y).ObjInfo.objindex = 0
    MapData(map, X, Y).ObjInfo.Amount = 0
    
    Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectDelete(X, Y))
End If

End Sub
Sub MakeObj(ByRef Obj As Obj, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

If Obj.objindex > 0 And Obj.objindex <= UBound(ObjData) Then

    If MapData(map, X, Y).ObjInfo.objindex = Obj.objindex Then
        MapData(map, X, Y).ObjInfo.Amount = MapData(map, X, Y).ObjInfo.Amount + Obj.Amount
    Else
        MapData(map, X, Y).ObjInfo = Obj
        
        Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(Obj.objindex).GrhIndex, X, Y))
    End If
End If

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As Obj) As Boolean
On Error GoTo errhandler

'Call LogTarea("MeterItemEnInventario")
 
Dim X As Integer
Dim Y As Integer
Dim Slot As Byte

'¿el user ya tiene un objeto del mismo tipo?
Slot = 1
Do Until UserList(UserIndex).Invent.Object(Slot).objindex = MiObj.objindex And _
         UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
   Slot = Slot + 1
   If Slot > MAX_INVENTORY_SLOTS Then
         Exit Do
   End If
Loop
    
'Sino busca un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(UserIndex).Invent.Object(Slot).objindex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
           Call WriteConsoleMsg(UserIndex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If
    
'Mete el objeto
If UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(UserIndex).Invent.Object(Slot).objindex = MiObj.objindex
   UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount
Else
   UserList(UserIndex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, UserIndex, Slot)


Exit Function
errhandler:

End Function

Sub GetObj(ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 18/12/2009
'30/08/2011: Shak - Ahora el oro va al inventario como los objetos.
'***************************************************
 
    Dim Obj As ObjData
    Dim MiObj As Obj
    Dim ObjPos As String
   
    With UserList(UserIndex)
        '¿Hay algun obj?
        If MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.objindex > 0 Then
            '¿Esta permitido agarrar este obj?
            If ObjData(MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.objindex).Agarrable <> 1 Then
                Dim X As Integer
                Dim Y As Integer
                Dim Slot As Byte
               
                X = .Pos.X
                Y = .Pos.Y
               
                Obj = ObjData(MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.objindex)
                MiObj.Amount = MapData(.Pos.map, X, Y).ObjInfo.Amount
                                 
                MiObj.objindex = MapData(.Pos.map, X, Y).ObjInfo.objindex
                
                ' @@ Cuicui
                If MiObj.Amount <= 0 Then
                    Call LogError(.name & " intentó agarrar un item con Cantidad:" & MiObj.Amount & ". Mapa:" & .Pos.map)
                    Exit Sub
                End If
                
                Dim RESULTADO As Boolean
                
                RESULTADO = (ObjData(MiObj.objindex).OBJType = 29)
                
                If RESULTADO Then Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha agarrado una " & ObjData(MiObj.objindex).name & ". Se encuentra en el mapa " & UserList(UserIndex).Pos.map, FontTypeNames.FONTTYPE_GUILD))
                
                ' El oro se va al inventario
                If ObjData(MiObj.objindex).OBJType = eOBJType.otGuita Then
                    
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + MiObj.Amount
                    If UserList(UserIndex).Stats.GLD > MaxOro Then UserList(UserIndex).Stats.GLD = MaxOro
                    Call WriteUpdateGold(UserIndex)
                    Call EraseObj(MapData(.Pos.map, X, Y).ObjInfo.Amount, .Pos.map, .Pos.X, .Pos.Y)
                    Exit Sub
                    
                End If
                
                If MeterItemEnInventario(UserIndex, MiObj) Then
                 'Lukea oro, actualizo ranking ***
               '     Dim targetRankPos As Byte
        
                '   targetRankPos = MOd_DunkanRankings.IngresaOro(UserIndex)
        
                '    If targetRankPos <> 0 Then Call MOd_DunkanRankings.ActualizarOros(UserIndex, targetRankPos)
                        'Quitamos el objeto
                        Call EraseObj(MapData(.Pos.map, X, Y).ObjInfo.Amount, .Pos.map, .Pos.X, .Pos.Y)
                        If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.objindex).name)
       
                        'Log de Objetos que se agarran del piso. Pablo (ToxicWaste) 07/09/07
                        'Es un Objeto que tenemos que loguear?
                        If ObjData(MiObj.objindex).LOG = 1 Then
                            ObjPos = " Mapa: " & .Pos.map & " X: " & .Pos.X & " Y: " & .Pos.Y
                            Call LogDesarrollo(.name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.objindex).name & ObjPos)
                        ElseIf MiObj.Amount > MAX_INVENTORY_OBJS - 1000 Then 'Es mucha cantidad?
                            'Si no es de los prohibidos de loguear, lo logueamos.
                            If ObjData(MiObj.objindex).NoLog <> 1 Then
                                ObjPos = " Mapa: " & .Pos.map & " X: " & .Pos.X & " Y: " & .Pos.Y
                                Call LogDesarrollo(.name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.objindex).name & ObjPos)
                            End If
                        End If
                    End If
               
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "No hay nada aquí.", FontTypeNames.FONTTYPE_INFO)
       End If
    End With
 
End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler

    'Desequipa el item slot del inventario
    Dim Obj As ObjData
    
    With UserList(UserIndex)
        With .Invent
            If (Slot < LBound(.Object)) Or (Slot > UBound(.Object)) Then
                Exit Sub
            ElseIf .Object(Slot).objindex = 0 Then
                Exit Sub
            End If
            
            Obj = ObjData(.Object(Slot).objindex)
        End With
        
        Select Case Obj.OBJType
            Case eOBJType.otWeapon
                With .Invent
                    .Object(Slot).Equipped = 0
                    .WeaponEqpObjIndex = 0
                    .WeaponEqpSlot = 0
                End With
                
                If Not .flags.Mimetizado = 1 Then
                    With .Char
                        .WeaponAnim = NingunArma
                        Call ChangeUserChar(UserIndex, .body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                    End With
                End If
            
            Case eOBJType.otFlechas
                With .Invent
                    .Object(Slot).Equipped = 0
                    .MunicionEqpObjIndex = 0
                    .MunicionEqpSlot = 0
                End With
            
            
            Case eOBJType.otManchas
                With .Invent
                    .Object(Slot).Equipped = 0
                    .MunicionEqpObjIndex = 0
                    .MunicionEqpSlot = 0
                End With
 
            
            Case eOBJType.otAnillo
                With .Invent
                    .Object(Slot).Equipped = 0
                    .AnilloEqpObjIndex = 0
                    .AnilloEqpSlot = 0
                End With
            
            Case eOBJType.otarmadura
                With .Invent
                    .Object(Slot).Equipped = 0
                    .ArmourEqpObjIndex = 0
                    .ArmourEqpSlot = 0
                End With
                
                Call DarCuerpoDesnudo(UserIndex, .flags.Mimetizado = 1)
                With .Char
                    Call ChangeUserChar(UserIndex, .body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                End With
                 
            Case eOBJType.otcasco
                With .Invent
                    .Object(Slot).Equipped = 0
                    .CascoEqpObjIndex = 0
                    .CascoEqpSlot = 0
                End With
                
                If Not .flags.Mimetizado = 1 Then
                    With .Char
                        .CascoAnim = NingunCasco
                        Call ChangeUserChar(UserIndex, .body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                    End With
                End If
            
            Case eOBJType.otescudo
                With .Invent
                    .Object(Slot).Equipped = 0
                    .EscudoEqpObjIndex = 0
                    .EscudoEqpSlot = 0
                End With
                
                If Not .flags.Mimetizado = 1 Then
                    With .Char
                        .ShieldAnim = NingunEscudo
                        Call ChangeUserChar(UserIndex, .body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                    End With
                End If
            
            Case eOBJType.otMochilas
                With .Invent
                    .Object(Slot).Equipped = 0
                    .MochilaEqpObjIndex = 0
                    .MochilaEqpSlot = 0
                End With
                
                Call InvUsuario.TirarTodosLosItemsEnMochila(UserIndex)
                .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
        End Select
    End With
    
    Call WriteUpdateUserStats(UserIndex)
    Call UpdateUserInv(False, UserIndex, Slot)
    
    ActualizarAuras UserIndex
    
    Exit Sub

errhandler:
    Call LogError("Error en Desquipar. Error " & Err.Number & " : " & Err.description)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal objindex As Integer, Optional ByRef sMotivo As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
'***************************************************

On Error GoTo errhandler
    
    If ObjData(objindex).Mujer = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Hombre
    ElseIf ObjData(objindex).Hombre = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Mujer
    Else
        SexoPuedeUsarItem = True
    End If
    
    If Not SexoPuedeUsarItem Then sMotivo = "Tu género no puede usar este objeto."
    
    Exit Function
errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal objindex As Integer, Optional ByRef sMotivo As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
'***************************************************

    If ObjData(objindex).Real = 1 Then
        If Not criminal(UserIndex) Then
            FaccionPuedeUsarItem = esArmada(UserIndex)
        Else
            FaccionPuedeUsarItem = False
        End If
    ElseIf ObjData(objindex).Caos = 1 Then
        If criminal(UserIndex) Then
            FaccionPuedeUsarItem = esCaos(UserIndex)
        Else
            FaccionPuedeUsarItem = False
        End If
    Else
        FaccionPuedeUsarItem = True
    End If
    
    If Not FaccionPuedeUsarItem Then sMotivo = "Tu alineación no puede usar este objeto."

End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'*************************************************
'Author: Unknown
'Last modified: 14/01/2010 (ZaMa)
'01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin
'14/01/2010: ZaMa - Agrego el motivo especifico por el que no puede equipar/usar el item.
'*************************************************

On Error GoTo errhandler

    'Equipa un item del inventario
    Dim Obj As ObjData
    Dim objindex As Integer
    Dim sMotivo As String
    
    With UserList(UserIndex)
        objindex = .Invent.Object(Slot).objindex
        Obj = ObjData(objindex)
        
        If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
             Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If Obj.Premium And Not EsPremium(UserIndex) Then
        WriteConsoleMsg UserIndex, "Sólo los PREMIUM pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO
        Exit Sub
        End If
        
 If Obj.OBJType = otWeapon Or Obj.OBJType = otarmadura Then
If .Stats.UserSkills(eSkill.Magia) < Obj.MagiaSkill Then
Call WriteConsoleMsg(UserIndex, "Para poder utilizar este ítem es necesario tener " & Obj.MagiaSkill & " skills en Mágia.", FontTypeNames.FONTTYPE_INFO)
Exit Sub
End If
End If
        
        If Obj.OBJType = otAnillo Then
If .Stats.UserSkills(eSkill.Resistencia) < Obj.RMSkill Then
Call WriteConsoleMsg(UserIndex, "Para poder utilizar este ítem es necesario tener " & Obj.RMSkill & " skills en Resistencia Mágica.", FontTypeNames.FONTTYPE_INFO)
Exit Sub
End If
End If

If Obj.OBJType = otWeapon Then
If .Stats.UserSkills(eSkill.Armas) < Obj.ArmaSkill Then
Call WriteConsoleMsg(UserIndex, "Para usar este ítem tienes que tener " & Obj.ArmaSkill & " skills en Combate con Armas.", FontTypeNames.FONTTYPE_INFO)
Exit Sub
End If
End If

If Obj.OBJType = otescudo Then
If .Stats.UserSkills(eSkill.Defensa) < Obj.EscudoSkill Then
Call WriteConsoleMsg(UserIndex, "Para usar este ítem tienes que tener " & Obj.EscudoSkill & " skills en Defensa con Escudos.", FontTypeNames.FONTTYPE_INFO)
Exit Sub
End If
End If

If Obj.OBJType = otcasco Or Obj.OBJType = otarmadura Then
If .Stats.UserSkills(eSkill.Tacticas) < Obj.ArmaduraSkill Then
Call WriteConsoleMsg(UserIndex, "Para usar este ítem tienes que tener " & Obj.ArmaduraSkill & " skills en Tácticas de Combate.", FontTypeNames.FONTTYPE_INFO)
Exit Sub
End If
End If

If Obj.OBJType = otWeapon Then
If .Stats.UserSkills(eSkill.Proyectiles) < Obj.ArcoSkill Then
Call WriteConsoleMsg(UserIndex, "Para usar este item tienes que tener " & Obj.ArcoSkill & " skills en Armas de Proyectiles.", FontTypeNames.FONTTYPE_INFO)
Exit Sub
End If
End If

If Obj.OBJType = otWeapon Then
If .Stats.UserSkills(eSkill.Apuñalar) < Obj.DagaSkill Then
Call WriteConsoleMsg(UserIndex, "Para utilizar este ítem necesitas " & Obj.DagaSkill & " skills en Apuñalar.", FontTypeNames.FONTTYPE_INFO)
Exit Sub
End If
End If

     If Obj.OBJType = otMonturas Then
If .Stats.UserSkills(eSkill.Equitacion) < Obj.Monturasskill Then
Call WriteConsoleMsg(UserIndex, "Para utilizar esta montura necesitas " & Obj.Monturasskill & " skills en Equitación.", FontTypeNames.FONTTYPE_INFO)
Exit Sub
End If
End If

     If Obj.OBJType = otMonturasDraco Then
If .Stats.UserSkills(eSkill.Equitacion) < Obj.MonturasDracoskill Then
Call WriteConsoleMsg(UserIndex, "Para utilizar esta montura necesitas " & Obj.MonturasDracoskill & " skills en Equitación.", FontTypeNames.FONTTYPE_INFO)
Exit Sub
End If
End If

If Obj.VIP = 1 And UserList(UserIndex).flags.Oro = 0 Then
WriteConsoleMsg UserIndex, "¡Sólo los usuarios Oro pueden ocupar estos ítems!", FontTypeNames.FONTTYPE_INFO
      Exit Sub
End If

If Obj.VIPP = 1 And UserList(UserIndex).flags.Plata = 0 Then
WriteConsoleMsg UserIndex, "¡Sólo los usuarios Plata pueden ocupar estos ítems!", FontTypeNames.FONTTYPE_INFO
      Exit Sub
End If

If Obj.VIPB = 1 And UserList(UserIndex).flags.Bronce = 0 Then
WriteConsoleMsg UserIndex, "¡Sólo los usuarios Bronce pueden ocupar estos ítems!", FontTypeNames.FONTTYPE_INFO
      Exit Sub
End If


        If Obj.Quince = 1 And Not EsQuinceM(UserIndex) Then
             Call WriteConsoleMsg(UserIndex, "Item restringido para nivel 15 o inferior.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
               If Obj.Treinta = 1 And Not EsTreintaM(UserIndex) Then
             Call WriteConsoleMsg(UserIndex, "Item restringido para nivel 13 o superior.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
                      If Obj.HM = 1 And Not EsHM(UserIndex) Then
             Call WriteConsoleMsg(UserIndex, "Item restringido para nivel 30 o Superior.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
                              If Obj.UM = 1 And Not EsUM(UserIndex) Then
             Call WriteConsoleMsg(UserIndex, "Item restringido para nivel 35 o superior.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
                      If Obj.MM = 1 And Not EsMM(UserIndex) Then
             Call WriteConsoleMsg(UserIndex, "Item restringido para nivel 45 o superior.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If

        Select Case Obj.OBJType
            Case eOBJType.otWeapon
               If ClasePuedeUsarItem(UserIndex, objindex, sMotivo) And _
                  FaccionPuedeUsarItem(UserIndex, objindex, sMotivo) Then
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)
                        'Animacion por defecto
                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.WeaponAnim = NingunArma
                        Else
                            .Char.WeaponAnim = NingunArma
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                        Exit Sub
                    End If
                    
                    'Quitamos el elemento anterior
                    If .Invent.WeaponEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                    End If
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.WeaponEqpObjIndex = objindex
                    .Invent.WeaponEqpSlot = Slot
                    
                    'El sonido solo se envia si no lo produce un admin invisible
                    If Not (.flags.AdminInvisible = 1) Then _
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.Y))
                    
If .flags.Mimetizado = 1 Then
                        .CharMimetizado.WeaponAnim = GetWeaponAnim(UserIndex, objindex)
                    Else
                        .Char.WeaponAnim = GetWeaponAnim(UserIndex, objindex)
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                    If .flags.Montando = 0 Then
                    .Char.WeaponAnim = GetWeaponAnim(UserIndex, objindex)
                    Else
                    .Char.WeaponAnim = NingunArma
                     Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                     End If
               Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
               End If
            
                           Case eOBJType.otBarcos
                 Dim Barco As ObjData
           Dim ModNave As Long
          Barco = ObjData(objindex)
      ModNave = ModNavegacion(.clase, UserIndex)
       If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
           WriteConsoleMsg UserIndex, "Necesitas " & Barco.MinSkill * 2 & " puntos en Navegación para equipar el barco.", FontTypeNames.FONTTYPE_INFO
           Exit Sub
           End If
               'Si esta equipado lo quita
                        If .Invent.Object(Slot).Equipped Then
                            'Quitamos del inv el item
                            Call Desequipar(UserIndex, Slot)
                            Exit Sub
                        End If
                        
                        'Quitamos el elemento anterior
                        If .Invent.BarcoObjIndex > 0 Then
                            Call Desequipar(UserIndex, .Invent.BarcoSlot)
                        End If
                
                        .Invent.Object(Slot).Equipped = 1
                        .Invent.BarcoObjIndex = objindex
                        .Invent.BarcoSlot = Slot
            
            Case eOBJType.otAnillo
               If ClasePuedeUsarItem(UserIndex, objindex, sMotivo) And _
                  FaccionPuedeUsarItem(UserIndex, objindex, sMotivo) Then
                        'Si esta equipado lo quita
                        If .Invent.Object(Slot).Equipped Then
                            'Quitamos del inv el item
                            Call Desequipar(UserIndex, Slot)
                            Exit Sub
                        End If
                        
                        'Quitamos el elemento anterior
                        If .Invent.AnilloEqpObjIndex > 0 Then
                            Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)
                        End If
                
                        .Invent.Object(Slot).Equipped = 1
                        .Invent.AnilloEqpObjIndex = objindex
                        .Invent.AnilloEqpSlot = Slot
                        
               Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
               End If
            
            Case eOBJType.otManchas
               If ClasePuedeUsarItem(UserIndex, objindex, sMotivo) And _
                  FaccionPuedeUsarItem(UserIndex, objindex, sMotivo) Then
                   
                        'Si esta equipado lo quita
                        If .Invent.Object(Slot).Equipped Then
                            'Quitamos del inv el item
                            Call Desequipar(UserIndex, Slot)
                            Exit Sub
                        End If
 
                        'Quitamos el elemento anterior
                        If .Invent.MunicionEqpObjIndex > 0 Then
                            Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
                        End If
                   
                        .Invent.Object(Slot).Equipped = 1
                        .Invent.MunicionEqpObjIndex = objindex
                        .Invent.MunicionEqpSlot = Slot
                   
               Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
               End If
 
            
            Case eOBJType.otFlechas
               If ClasePuedeUsarItem(UserIndex, objindex, sMotivo) And _
                  FaccionPuedeUsarItem(UserIndex, objindex, sMotivo) Then
                        
                        'Si esta equipado lo quita
                        If .Invent.Object(Slot).Equipped Then
                            'Quitamos del inv el item
                            Call Desequipar(UserIndex, Slot)
                            Exit Sub
                        End If
                        
                        'Quitamos el elemento anterior
                        If .Invent.MunicionEqpObjIndex > 0 Then
                            Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
                        End If
                
                        .Invent.Object(Slot).Equipped = 1
                        .Invent.MunicionEqpObjIndex = objindex
                        .Invent.MunicionEqpSlot = Slot
                        
               Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
               End If
            
            Case eOBJType.otarmadura
            If .flags.Montando = 1 Then Exit Sub
                If .flags.Navegando = 1 Then Exit Sub
                
                'Nos aseguramos que puede usarla
                If ClasePuedeUsarItem(UserIndex, objindex, sMotivo) And _
                   SexoPuedeUsarItem(UserIndex, objindex, sMotivo) And _
                   CheckRazaUsaRopa(UserIndex, objindex, sMotivo) And _
                   FaccionPuedeUsarItem(UserIndex, objindex, sMotivo) Then
                   
                   'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        Call DarCuerpoDesnudo(UserIndex, .flags.Mimetizado = 1)
                         If Not .flags.Mimetizado = 1 Or .flags.Montando Then
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                        Exit Sub
                    End If
            
                    'Quita el anterior
                    If .Invent.ArmourEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
                    End If
            
                    'Lo equipa
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.ArmourEqpObjIndex = objindex
                    .Invent.ArmourEqpSlot = Slot
                        
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.body = Obj.Ropaje
                    Else
                        .Char.body = Obj.Ropaje
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                    .flags.Desnudo = 0
                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eOBJType.otcasco
                If .flags.Navegando = 1 Then Exit Sub
                If ClasePuedeUsarItem(UserIndex, objindex, sMotivo) Then
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.CascoAnim = NingunCasco
                        Else
                            .Char.CascoAnim = NingunCasco
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                        Exit Sub
                    End If
            
                    'Quita el anterior
                    If .Invent.CascoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
                    End If
            
                    'Lo equipa
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.CascoEqpObjIndex = objindex
                    .Invent.CascoEqpSlot = Slot
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.CascoAnim = Obj.CascoAnim
                    Else
                        .Char.CascoAnim = Obj.CascoAnim
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eOBJType.otescudo
                If .flags.Navegando = 1 Then Exit Sub
                If .flags.Montando = 1 Then Exit Sub
                
                 If ClasePuedeUsarItem(UserIndex, objindex, sMotivo) And _
                     FaccionPuedeUsarItem(UserIndex, objindex, sMotivo) Then
        
                     'Si esta equipado lo quita
                     If .Invent.Object(Slot).Equipped Then
                         Call Desequipar(UserIndex, Slot)
                         If .flags.Mimetizado = 1 Then
                             .CharMimetizado.ShieldAnim = NingunEscudo
                         Else
                             .Char.ShieldAnim = NingunEscudo
                             Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                         End If
                         Exit Sub
                     End If
             
                     'Quita el anterior
                     If .Invent.EscudoEqpObjIndex > 0 Then
                         Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
                     End If
             
                     'Lo equipa
                     
                     .Invent.Object(Slot).Equipped = 1
                     .Invent.EscudoEqpObjIndex = objindex
                     .Invent.EscudoEqpSlot = Slot
                     
                     If .flags.Mimetizado = 1 Then
                         .CharMimetizado.ShieldAnim = Obj.ShieldAnim
                     Else
                         .Char.ShieldAnim = Obj.ShieldAnim
                         
                         Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                     End If
                 Else
                     Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                 End If
                 
            Case eOBJType.otMochilas
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If .Invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                If .Invent.MochilaEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, .Invent.MochilaEqpSlot)
                End If
                .Invent.Object(Slot).Equipped = 1
                .Invent.MochilaEqpObjIndex = objindex
                .Invent.MochilaEqpSlot = Slot
                .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + Obj.MochilaType * 5
                'Call WriteAddSlots(UserIndex, Obj.MochilaType)
        End Select
    End With
    
    'Actualiza
    ActualizarAuras UserIndex
    
    Call UpdateUserInv(False, UserIndex, Slot)
    
    Exit Sub
    
errhandler:
    Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.description)
End Sub

Private Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer, Optional ByRef sMotivo As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
'***************************************************

On Error GoTo errhandler

    With UserList(UserIndex)
        'Verifica si la raza puede usar la ropa
        If .Raza = eRaza.Humano Or _
           .Raza = eRaza.Elfo Or _
           .Raza = eRaza.Drow Then
                CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
        Else
                CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
        End If
        
        'Solo se habilita la ropa exclusiva para Drows por ahora. Pablo (ToxicWaste)
        If (.Raza <> eRaza.Drow) And ObjData(ItemIndex).RazaDrow Then
            CheckRazaUsaRopa = False
        End If
    End With
    
    If Not CheckRazaUsaRopa Then sMotivo = "Tu raza no puede usar este objeto."
    
    Exit Function
    
errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvPotion(ByVal UserIndex As Integer, _
                 ByVal Slot As Byte, _
                 ByVal SecondaryClick As Byte)

    Dim Obj As ObjData
    
    With UserList(UserIndex)

        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo. ", _
                    FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If .Invent.Object(Slot).Amount = 0 Then Exit Sub
        
        Obj = ObjData(.Invent.Object(Slot).objindex)
        
        If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If SecondaryClick Then
            If Not IntervaloPermiteUsarClick(UserIndex) Then Exit Sub
        Else

            If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
            
        End If
          
        If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then
            Call WriteConsoleMsg(UserIndex, "¡¡Debes esperar unos momentos para tomar otra poción!!", _
                    FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
                        
        Select Case Obj.OBJType
            
            Case eOBJType.otPociones
                
                .flags.TomoPocion = True
                .flags.TipoPocion = Obj.TipoPocion
                        
                Select Case .flags.TipoPocion
                
                    Case 1 'Modif la agilidad
                        .flags.DuracionEfecto = Obj.DuracionEfecto
                
                        'Usa el item
                        .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + _
                                RandomNumber(Obj.MinModificador, Obj.MaxModificador)

                        If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then
                            .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS

                        End If

                        If .Stats.UserAtributos(eAtributos.Agilidad) > 2 * .Stats.UserAtributosBackUP(Agilidad) Then
                            .Stats.UserAtributos(eAtributos.Agilidad) = 2 * .Stats.UserAtributosBackUP(Agilidad)

                        End If
                        
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, _
                                    .Pos.Y))

                        End If

                        Call WriteUpdateDexterity(UserIndex)
                        
                    Case 2 'Modif la fuerza
                        .flags.DuracionEfecto = Obj.DuracionEfecto
                
                        'Usa el item
                        .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + _
                                RandomNumber(Obj.MinModificador, Obj.MaxModificador)

                        If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then
                            .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS

                        End If

                        If .Stats.UserAtributos(eAtributos.Fuerza) > 2 * .Stats.UserAtributosBackUP(Fuerza) Then
                            .Stats.UserAtributos(eAtributos.Fuerza) = 2 * .Stats.UserAtributosBackUP(Fuerza)

                        End If

                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, _
                                    .Pos.Y))

                        End If

                        Call WriteUpdateStrenght(UserIndex)
                        
                    Case 3 'Pocion roja, restaura HP
                        'Usa el item
                        .Stats.MinHp = .Stats.MinHp + RandomNumber(Obj.MinModificador, Obj.MaxModificador)

                        If .Stats.MinHp > .Stats.MaxHp Then
                            .Stats.MinHp = .Stats.MaxHp

                        End If

                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, _
                                    .Pos.Y))

                        End If

                        Call WriteUpdateHP(UserIndex)
                    
                    Case 4 'Pocion azul, restaura MANA
                        'Usa el item
                        'nuevo calculo para recargar mana
                        .Stats.MinMAN = .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, 4) + .Stats.ELV \ 2 + 40 / .Stats.ELV

                        If .Stats.MinMAN > .Stats.MaxMAN Then
                            .Stats.MinMAN = .Stats.MaxMAN

                        End If

                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, _
                                    .Pos.Y))

                        End If
                        
                        Call WriteUpdateMana(UserIndex)
                        
                    Case 5 ' Pocion violeta

                        If .flags.Envenenado = 1 Then
                            .flags.Envenenado = 0
                            Call WriteConsoleMsg(UserIndex, "Te has curado del envenenamiento.", _
                                    FontTypeNames.FONTTYPE_INFO)

                        End If

                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, _
                                    .Pos.Y))

                        End If

                        Call WriteUpdateUserStats(UserIndex)
                        
                    Case 6  ' Pocion Negra

                        If .flags.Privilegios And PlayerType.User Then
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Call UserDie(UserIndex)
                            Call WriteConsoleMsg(UserIndex, "Sientes un gran mareo y pierdes el conocimiento.", _
                                    FontTypeNames.FONTTYPE_FIGHT)

                        End If
                        
                    Case 7 'pocion energia
                    
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                          ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, _
                                    .Pos.Y))

                        End If
                        
                        .Stats.MinSta = .Stats.MinSta + (.Stats.MaxSta * 0.1)
    
                        If .Stats.MinSta > .Stats.MaxSta Then
                            .Stats.MinSta = .Stats.MaxSta

                        End If
                        
                        'Call AddtoVar(UserList(UserIndex).Stats.MinSta, UserList(UserIndex).Stats.MaxSta * 0.1, UserList(UserIndex).Stats.MaxSta)
                        'If UserList(UserIndex).Stats.MinSta > UserList(UserIndex).Stats.MaxSta Then UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
                        Call WriteUpdateSta(UserIndex)

                End Select
                
                Call UpdateUserInv(False, UserIndex, Slot)
                
        End Select
    
    End With

End Sub

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'*************************************************
'Author: Unknown
'Last modified: 10/12/2009
'Handels the usage of items from inventory box.
'24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legión.
'24/01/2007 Pablo (ToxicWaste) - Utilización nueva de Barco en lvl 20 por clase Pirata y Pescador.
'01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin, except to its own client
'17/11/2009: ZaMa - Ahora se envia una orientacion de la posicion hacia donde esta el que uso el cuerno.
'27/11/2009: Budi - Se envia indivualmente cuando se modifica a la Agilidad o la Fuerza del personaje.
'08/12/2009: ZaMa - Agrego el uso de hacha de madera elfica.
'10/12/2009: ZaMa - Arreglos y validaciones en todos las herramientas de trabajo.
'*************************************************

    Dim Obj As ObjData
    Dim objindex As Integer
    Dim TargObj As ObjData
    Dim MiObj As Obj
    
    With UserList(UserIndex)
    
        If .Invent.Object(Slot).Amount = 0 Then Exit Sub
        
        Obj = ObjData(.Invent.Object(Slot).objindex)
        
        If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Obj.OBJType = eOBJType.otWeapon Then
            If Obj.proyectil = 1 Then
                If Not .flags.ModoCombate Then
             Call WriteConsoleMsg(UserIndex, "Para realizar esta accion debes activar el modo combate, puedes hacerlo con la tecla ""C""", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
             End If
                'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
                If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
            Else
                'dagas
                If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
            End If
        Else
            If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
        End If
        
        objindex = .Invent.Object(Slot).objindex
        .flags.TargetObjInvIndex = objindex
        .flags.TargetObjInvSlot = Slot
        
        Select Case Obj.OBJType
            Case eOBJType.otUseOnce
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
        
                'Usa el item
                .Stats.MinHam = .Stats.MinHam + Obj.MinHam
                If .Stats.MinHam > .Stats.MaxHam Then _
                    .Stats.MinHam = .Stats.MaxHam
                .flags.Hambre = 0
                Call WriteUpdateHungerAndThirst(UserIndex)
                'Sonido
                
                If objindex = e_ObjetosCriticos.Manzana Or objindex = e_ObjetosCriticos.Manzana2 Or objindex = e_ObjetosCriticos.ManzanaNewbie Then
                    Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MORFAR_MANZANA)
                Else
                    Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.SOUND_COMIDA)
                End If
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                Call UpdateUserInv(False, UserIndex, Slot)
        
            Case eOBJType.otGuita
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                .Stats.GLD = .Stats.GLD + .Invent.Object(Slot).Amount
                .Invent.Object(Slot).Amount = 0
                .Invent.Object(Slot).objindex = 0
                .Invent.NroItems = .Invent.NroItems - 1
                
                Call UpdateUserInv(False, UserIndex, Slot)
                Call WriteUpdateGold(UserIndex)
                
            Case eOBJType.otWeapon
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Not .Stats.MinSta > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Estás muy cansad" & _
                                IIf(.Genero = eGenero.Hombre, "o", "a") & ".", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If ObjData(objindex).proyectil = 1 Then
                    If .Invent.Object(Slot).Equipped = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberías equipartela.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
            If Not .flags.ModoCombate Then
                Call WriteConsoleMsg(UserIndex, "¡¡No puedes lanzar flechas si no estas en modo combate!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
                    Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Proyectiles)  'Call WriteWorkRequestTarget(UserIndex, Proyectiles)
                ElseIf .flags.TargetObj = Leña Then
                    If .Invent.Object(Slot).objindex = DAGA Then
                        If .Invent.Object(Slot).Equipped = 0 Then
                            Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberías equipartela.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                            
                        Call TratarDeHacerFogata(.flags.TargetObjMap, _
                            .flags.TargetObjX, .flags.TargetObjY, UserIndex)
                    End If
                Else
                
                    Select Case objindex
                        Case CAÑA_PESCA, RED_PESCA
                            If .Invent.WeaponEqpObjIndex = CAÑA_PESCA Or .Invent.WeaponEqpObjIndex = RED_PESCA Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Pesca)  'Call WriteWorkRequestTarget(UserIndex, eSkill.Pesca)
                            Else
                                 Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                            
                        Case HACHA_LEÑADOR, HACHA_DORADA
                            If .Invent.WeaponEqpObjIndex = HACHA_LEÑADOR Or .Invent.WeaponEqpObjIndex = HACHA_DORADA Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.talar)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                            
                        Case PIQUETE_MINERO
                            If .Invent.WeaponEqpObjIndex = PIQUETE_MINERO Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Mineria)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                            
                              Case PIQUETE_ORO
                            If .Invent.WeaponEqpObjIndex = PIQUETE_ORO Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Mineria)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                            
                        Case MARTILLO_HERRERO
                            If .Invent.WeaponEqpObjIndex = MARTILLO_HERRERO Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.herreria)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                            
                             Case SERRUCHO_CARPINTERO
                            If .Invent.WeaponEqpObjIndex = SERRUCHO_CARPINTERO Then
                                Call EnivarObjConstruibles(UserIndex)
                                Call WriteShowCarpenterForm(UserIndex)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                    End Select
                End If
                            
                        
      
             Case eOBJType.otBebidas
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed
                If .Stats.MinAGU > .Stats.MaxAGU Then _
                    .Stats.MinAGU = .Stats.MaxAGU
                .flags.Sed = 0
                Call WriteUpdateHungerAndThirst(UserIndex)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                End If
                
                Call UpdateUserInv(False, UserIndex, Slot)
            
            Case eOBJType.otLlaves
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .flags.TargetObj = 0 Then Exit Sub
                TargObj = ObjData(.flags.TargetObj)
                '¿El objeto clickeado es una puerta?
                If TargObj.OBJType = eOBJType.otPuertas Then
                    '¿Esta cerrada?
                    If TargObj.Cerrada = 1 Then
                          '¿Cerrada con llave?
                          If TargObj.Llave > 0 Then
                             If TargObj.clave = Obj.clave Then
                 
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objindex _
                                = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objindex).IndexCerrada
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objindex
                                Call WriteConsoleMsg(UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                             Else
                                Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                             End If
                          Else
                             If TargObj.clave = Obj.clave Then
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objindex _
                                = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objindex).IndexCerradaLlave
                                Call WriteConsoleMsg(UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objindex
                                Exit Sub
                             Else
                                Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                             End If
                          End If
                    Else
                          Call WriteConsoleMsg(UserIndex, "No está cerrada.", FontTypeNames.FONTTYPE_INFO)
                          Exit Sub
                    End If
                End If
            
            Case eOBJType.otBotellaVacia
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If Not HayAgua(.Pos.map, .flags.TargetX, .flags.TargetY) Then
                    Call WriteConsoleMsg(UserIndex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                MiObj.Amount = 1
                MiObj.objindex = ObjData(.Invent.Object(Slot).objindex).IndexAbierta
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)
                End If
                
                Call UpdateUserInv(False, UserIndex, Slot)
            
            Case eOBJType.otBotellaLlena
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed
                If .Stats.MinAGU > .Stats.MaxAGU Then _
                    .Stats.MinAGU = .Stats.MaxAGU
                .flags.Sed = 0
                Call WriteUpdateHungerAndThirst(UserIndex)
                MiObj.Amount = 1
                MiObj.objindex = ObjData(.Invent.Object(Slot).objindex).IndexCerrada
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)
                End If
                
                Call UpdateUserInv(False, UserIndex, Slot)
            
            Case eOBJType.otPergaminos
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .Stats.MaxMAN > 0 Then
                    If .flags.Hambre = 0 And _
                        .flags.Sed = 0 Then
                        Call AgregarHechizo(UserIndex, Slot)
                        Call UpdateUserInv(False, UserIndex, Slot)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_INFO)
                End If
            Case eOBJType.otMinerales
                If .flags.Muerto = 1 Then
                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub
                End If
                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, FundirMetal) 'Call WriteWorkRequestTarget(UserIndex, FundirMetal)
               
            Case eOBJType.otInstrumentos
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Obj.Real Then '¿Es el Cuerno Real?
                    If FaccionPuedeUsarItem(UserIndex, objindex) Then
                        If MapInfo(.Pos.map).Pk = False Then
                            Call WriteConsoleMsg(UserIndex, "No hay peligro aquí. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                        Else
                            Call AlertarFaccionarios(UserIndex)
                            Call SendData(SendTarget.toMap, .Pos.map, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                        End If
                        
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(UserIndex, "Sólo miembros del ejército real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                ElseIf Obj.Caos Then '¿Es el Cuerno Legión?
                    If FaccionPuedeUsarItem(UserIndex, objindex) Then
                        If MapInfo(.Pos.map).Pk = False Then
                            Call WriteConsoleMsg(UserIndex, "No hay peligro aquí. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                        Else
                            Call AlertarFaccionarios(UserIndex)
                            Call SendData(SendTarget.toMap, .Pos.map, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                        End If
                        
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(UserIndex, "Sólo miembros de la legión oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
                'Si llega aca es porque es o Laud o Tambor o Flauta
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                End If
               
                       Case eOBJType.otBarcos
                'Verifica si esta aproximado al agua antes de permitirle navegar
                If .Stats.ELV < 25 Then
                    ' Solo pirata y trabajador pueden navegar antes
                    If .clase <> eClass.Worker And .clase <> eClass.Pirat Then
                        Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    Else
                        ' Pero a partir de 20
                        If .Stats.ELV < 20 Then
                            
                            If .clase = eClass.Worker And .Stats.UserSkills(eSkill.Pesca) <> 100 Then
                                Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 y además tu skill en pesca debe ser 100.", FontTypeNames.FONTTYPE_INFO)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)
                            End If
                            
                            Exit Sub
                        Else
                            ' Esta entre 20 y 25, si es trabajador necesita tener 100 en pesca
                            If .clase = eClass.Worker Then
                                If .Stats.UserSkills(eSkill.Pesca) <> 100 Then
                                    Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 o superior y además tu skill en pesca debe ser 100.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Sub
                                End If
                            End If

                        End If
                    End If
                End If
                
                If ((LegalPos(.Pos.map, .Pos.X - 1, .Pos.Y, True, False) _
                        Or LegalPos(.Pos.map, .Pos.X, .Pos.Y - 1, True, False) _
                        Or LegalPos(.Pos.map, .Pos.X + 1, .Pos.Y, True, False) _
                        Or LegalPos(.Pos.map, .Pos.X, .Pos.Y + 1, True, False)) _
                        And .flags.Navegando = 0) _
                        Or .flags.Navegando = 1 Then
                    Call DoNavega(UserIndex, Obj, Slot)
                Else
                    Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)
                End If
                
                
              Case eOBJType.otMonturas
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                End If
                If ((LegalPos(.Pos.map, .Pos.X, .Pos.Y, True, False) _
                        Or LegalPos(.Pos.map, .Pos.X, .Pos.Y, True, False) _
                        Or LegalPos(.Pos.map, .Pos.X, .Pos.Y, True, False) _
                        Or LegalPos(.Pos.map, .Pos.X, .Pos.Y, True, False)) _
                        And .flags.Navegando = 0) _
                        Or .flags.Navegando = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡No puedes montar en el agua!", FontTypeNames.FONTTYPE_INFO)
                Else
                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARMONTURA, .Pos.X, .Pos.Y))
                Call DoEquita(UserIndex, Obj, Slot)
                End If
                
                   Case eOBJType.otMonturasDraco
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                End If
                If ((LegalPos(.Pos.map, .Pos.X, .Pos.Y, True, False) _
                        Or LegalPos(.Pos.map, .Pos.X, .Pos.Y, True, False) _
                        Or LegalPos(.Pos.map, .Pos.X, .Pos.Y, True, False) _
                        Or LegalPos(.Pos.map, .Pos.X, .Pos.Y, True, False)) _
                        And .flags.Navegando = 0) _
                        Or .flags.Navegando = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡No puedes montar en el agua!", FontTypeNames.FONTTYPE_INFO)
                Else
                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARMONTURADRACO, .Pos.X, .Pos.Y))
                Call DoEquita(UserIndex, Obj, Slot)
                End If
                
        End Select
    
    End With

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Call WriteBlacksmithWeapons(UserIndex)
End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Call WriteCarpenterObjects(UserIndex)
End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Call WriteBlacksmithArmors(UserIndex)
End Sub

Sub TirarTodo(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

    With UserList(UserIndex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        Call TirarTodosLosItems(UserIndex)
        
    End With

End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With ObjData(index)
        ItemSeCae = (.Real <> 1 Or .NoSeCae = 0) And _
                    (.Caos <> 1 Or .NoSeCae = 0) And _
                    .OBJType <> eOBJType.otLlaves And _
                    .OBJType <> eOBJType.otBarcos And _
                    .OBJType <> eOBJType.otMonturas And _
                    .OBJType <> eOBJType.otMonturasDraco And _
                    .NoSeCae = 0
    End With

End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2010 (ZaMa)
'12/01/2010: ZaMa - Ahora los piratas no explotan items solo si estan entre 20 y 25
'***************************************************

    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    Dim DropAgua As Boolean
    
    With UserList(UserIndex)
        For i = 1 To .CurrentInventorySlots
            ItemIndex = .Invent.Object(i).objindex
            If ItemIndex > 0 Then
                 If ItemSeCae(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo el Obj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.objindex = ItemIndex

                    DropAgua = True
                    ' Es pirata?
                    If .clase = eClass.Pirat Then
                        ' Si tiene galeon equipado
                        If .Invent.BarcoObjIndex = 476 Then
                            ' Limitación por nivel, después dropea normalmente
                            If .Stats.ELV >= 20 And .Stats.ELV <= 25 Then
                                ' No dropea en agua
                                DropAgua = False
                            End If
                        End If
                    End If
                    
                    Call Tilelibre(.Pos, NuevaPos, MiObj, DropAgua, True)
                    
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
                    End If
                 End If
            End If
        Next i
    End With
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
    
    ItemNewbie = ObjData(ItemIndex).Newbie = 1
End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 23/11/2009
'07/11/09: Pato - Fix bug #2819911
'23/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    
    With UserList(UserIndex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        For i = 1 To UserList(UserIndex).CurrentInventorySlots
            ItemIndex = .Invent.Object(i).objindex
            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo MiObj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.objindex = ItemIndex
                    'Pablo (ToxicWaste) 24/01/2007
                    'Tira los Items no newbies en todos lados.
                    Tilelibre .Pos, NuevaPos, MiObj, True, True
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
                    End If
                End If
            End If
        Next i
    End With

End Sub

Sub TirarTodosLosItemsEnMochila(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/09 (Budi)
'***************************************************
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    
    With UserList(UserIndex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        For i = MAX_NORMAL_INVENTORY_SLOTS + 1 To .CurrentInventorySlots
            ItemIndex = .Invent.Object(i).objindex
            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo MiObj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.objindex = ItemIndex
                    Tilelibre .Pos, NuevaPos, MiObj, True, True
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
                    End If
                End If
            End If
        Next i
    End With

End Sub

Public Function getObjType(ByVal objindex As Integer) As eOBJType
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If objindex > 0 Then
        getObjType = ObjData(objindex).OBJType
    End If
    
End Function
Function ItemFaccionario(ByVal ItemIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
    
    ItemFaccionario = ObjData(ItemIndex).Caos Or ObjData(ItemIndex).Real = 1
End Function
Function ItemVIP(ByVal ItemIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function

     ItemVIP = ObjData(ItemIndex).VIP = 1
End Function
Function ItemVIPB(ByVal ItemIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function

     ItemVIPB = ObjData(ItemIndex).VIPB = 1
End Function
Function ItemVIPP(ByVal ItemIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function

     ItemVIPP = ObjData(ItemIndex).VIPP = 1
End Function
Function ItemQuince(ByVal ItemIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
    
    ItemQuince = ObjData(ItemIndex).Quince = 1
End Function
Function ItemTreinta(ByVal ItemIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
    
    ItemTreinta = ObjData(ItemIndex).Treinta = 1
End Function
Function ItemHM(ByVal ItemIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
    
    ItemHM = ObjData(ItemIndex).HM = 1
End Function
Function ItemUM(ByVal ItemIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
    
    ItemUM = ObjData(ItemIndex).UM = 1
End Function
Public Sub moveItem(ByVal UserIndex As Integer, ByVal originalSlot As Integer, ByVal newSlot As Integer)
 
Dim tmpObj As UserOBJ
Dim newObjIndex As Integer, originalObjIndex As Integer
If (originalSlot <= 0) Or (newSlot <= 0) Then Exit Sub
 
With UserList(UserIndex)
    If (originalSlot > .CurrentInventorySlots) Or (newSlot > .CurrentInventorySlots) Then Exit Sub
   
    tmpObj = .Invent.Object(originalSlot)
    .Invent.Object(originalSlot) = .Invent.Object(newSlot)
    .Invent.Object(newSlot) = tmpObj
   
    'Viva VB6 y sus putas deficiencias.
    If .Invent.AnilloEqpSlot = originalSlot Then
        .Invent.AnilloEqpSlot = newSlot
    ElseIf .Invent.AnilloEqpSlot = newSlot Then
        .Invent.AnilloEqpSlot = originalSlot
    End If
   
    If .Invent.ArmourEqpSlot = originalSlot Then
        .Invent.ArmourEqpSlot = newSlot
    ElseIf .Invent.ArmourEqpSlot = newSlot Then
        .Invent.ArmourEqpSlot = originalSlot
    End If
   
    If .Invent.BarcoSlot = originalSlot Then
        .Invent.BarcoSlot = newSlot
    ElseIf .Invent.BarcoSlot = newSlot Then
        .Invent.BarcoSlot = originalSlot
    End If
   
    If .Invent.CascoEqpSlot = originalSlot Then
         .Invent.CascoEqpSlot = newSlot
    ElseIf .Invent.CascoEqpSlot = newSlot Then
         .Invent.CascoEqpSlot = originalSlot
    End If
   
    If .Invent.EscudoEqpSlot = originalSlot Then
        .Invent.EscudoEqpSlot = newSlot
    ElseIf .Invent.EscudoEqpSlot = newSlot Then
        .Invent.EscudoEqpSlot = originalSlot
    End If
   
    If .Invent.MochilaEqpSlot = originalSlot Then
        .Invent.MochilaEqpSlot = newSlot
    ElseIf .Invent.MochilaEqpSlot = newSlot Then
        .Invent.MochilaEqpSlot = originalSlot
    End If
   
    If .Invent.MunicionEqpSlot = originalSlot Then
        .Invent.MunicionEqpSlot = newSlot
    ElseIf .Invent.MunicionEqpSlot = newSlot Then
        .Invent.MunicionEqpSlot = originalSlot
    End If
   
    If .Invent.WeaponEqpSlot = originalSlot Then
        .Invent.WeaponEqpSlot = newSlot
    ElseIf .Invent.WeaponEqpSlot = newSlot Then
        .Invent.WeaponEqpSlot = originalSlot
    End If
 
    Call UpdateUserInv(False, UserIndex, originalSlot)
    Call UpdateUserInv(False, UserIndex, newSlot)
End With
End Sub
 
