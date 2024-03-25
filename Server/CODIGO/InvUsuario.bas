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

    On Error GoTo ErrHandler

    Dim i        As Integer

    Dim ObjIndex As Integer
    
    For i = 1 To UserList(UserIndex).CurrentInventorySlots
        ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex

        If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And ObjData(ObjIndex).OBJType <> eOBJType.otBarcos And ObjData(ObjIndex).OBJType <> eOBJType.otMonturas And ObjData(ObjIndex).Bronce <> 1 And Not ItemNewbie(ObjIndex)) Then
                TieneObjetosRobables = True

                Exit Function

            End If

        End If

    Next i
    
    Exit Function

ErrHandler:
    Call LogError("Error en TieneObjetosRobables. Error: " & Err.number & " - " & Err.description)

End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, _
                            ByVal ObjIndex As Integer, _
                            Optional ByRef sMotivo As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '***************************************************

    On Error GoTo manejador
    
    'Admins can use ANYTHING!
    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
        If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then

            Dim i As Integer

            For i = 1 To NUMCLASES

                If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).Clase Then
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

Function ClasePuedeItem(ByVal Clase As Integer, ByVal ObjIndex As Integer) As Boolean

    On Error GoTo manejador
    
    If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then

        Dim i As Integer

        For i = 1 To NUMCLASES

            If ObjData(ObjIndex).ClaseProhibida(i) = Clase Then
                ClasePuedeItem = False
                Exit Function

            End If

        Next i

    End If
    
    ClasePuedeItem = True

    Exit Function

manejador:
    LogError ("Error en ClasePuedeItem")

End Function

' Comprueba si tiene objetos que para su level no está permitido usar más...
Sub QuitarLevelObj(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo QuitarLevelObj_Err

    '</EhHeader>
    
    Dim j As Long
    
    With UserList(UserIndex)

        For j = 1 To .CurrentInventorySlots

            If .Invent.Object(j).ObjIndex > 0 Then
                If ObjData(.Invent.Object(j).ObjIndex).LvlMax > 0 Then
                    If .Stats.Elv >= ObjData(.Invent.Object(j).ObjIndex).LvlMax Then
                        Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                        Call UpdateUserInv(False, UserIndex, j)

                    End If

                End If

            End If

        Next j
    
    End With

    '<EhFooter>
    Exit Sub

QuitarLevelObj_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.QuitarLevelObj " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub QuitarNewbieObj(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo QuitarNewbieObj_Err

    '</EhHeader>

    Dim j As Integer

    With UserList(UserIndex)

        For j = 1 To UserList(UserIndex).CurrentInventorySlots

            If .Invent.Object(j).ObjIndex > 0 Then
                If ObjData(.Invent.Object(j).ObjIndex).Newbie = 1 Then
                    Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, UserIndex, j)

                End If

            End If

        Next j

        'If MapInfo(.Pos.Map).Restringir = eRestrict.restrict_newbie Then
        
        'Call WarpUserChar(UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, True)
    
        'End If
    End With

    '<EhFooter>
    Exit Sub

QuitarNewbieObj_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.QuitarNewbieObj " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo LimpiarInventario_Err

    '</EhHeader>

    Dim j As Integer

    With UserList(UserIndex)

        For j = 1 To .CurrentInventorySlots
            .Invent.Object(j).ObjIndex = 0
            .Invent.Object(j).Amount = 0
            .Invent.Object(j).Equipped = 0
        Next j
    
        .Invent.NroItems = 0
    
        .Invent.ArmourEqpObjIndex = 0
        .Invent.ArmourEqpSlot = 0
    
        .Invent.WeaponEqpObjIndex = 0
        .Invent.WeaponEqpSlot = 0
    
        .Invent.AuraEqpObjIndex = 0
        .Invent.AuraEqpSlot = 0
    
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
    
        .Invent.MochilaEqpObjIndex = 0
        .Invent.MochilaEqpSlot = 0
    
        .Invent.MonturaObjIndex = 0
        .Invent.MochilaEqpSlot = 0
    
        .Invent.MagicObjIndex = 0
        .Invent.MagicSlot = 0

    End With

    '<EhFooter>
    Exit Sub

LimpiarInventario_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.LimpiarInventario " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, _
                      ByVal Slot As Byte, _
                      ByVal cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler
    
    Dim A As Long
    
    If Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots Then Exit Sub
    
    With UserList(UserIndex).Invent.Object(Slot)
    
        ' En eventos de cambio de clase,raza,level los objetos no se consumen. Excepto las Pociones
        A = UserList(UserIndex).flags.SlotEvent

        If A > 0 Then
            If Events(A).ChangeClass > 0 Or Events(A).ChangeRaze > 0 Or Events(A).ChangeLevel > 0 Then
                If (ObjData(.ObjIndex).OBJType = otFlechas) Then Exit Sub

            End If

        End If

        If .Amount <= cantidad And .Equipped = 1 Then
            Call Desequipar(UserIndex, Slot)

        End If
        
        'Quita un objeto
        .Amount = .Amount - cantidad

        '¿Quedan mas?
        If .Amount <= 0 Then
            UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
            
            .ObjIndex = 0
            .Amount = 0

        End If

    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en QuitarUserInvItem. Error " & Err.number & " : " & Err.description)
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, _
                  ByVal UserIndex As Integer, _
                  ByVal Slot As Byte)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim NullObj As UserOBJ

    Dim LoopC   As Long

    With UserList(UserIndex)

        'Actualiza un solo slot
        If Not UpdateAll Then
    
            'Actualiza el inventario
            If .Invent.Object(Slot).ObjIndex > 0 Then
                Call ChangeUserInv(UserIndex, Slot, .Invent.Object(Slot))
            Else
                Call ChangeUserInv(UserIndex, Slot, NullObj)

            End If
    
        Else
    
            'Actualiza todos los slots
            For LoopC = 1 To .CurrentInventorySlots

                'Actualiza el inventario
                If .Invent.Object(LoopC).ObjIndex > 0 Then
                    Call ChangeUserInv(UserIndex, LoopC, .Invent.Object(LoopC))
                Else
                    Call ChangeUserInv(UserIndex, LoopC, NullObj)

                End If
            
            Next LoopC

        End If
    
        Exit Sub

    End With

ErrHandler:
    Call LogError("Error en UpdateUserInv. Error " & Err.number & " : " & Err.description)

End Sub

Sub DropObj(ByVal UserIndex As Integer, _
            ByVal Slot As Byte, _
            ByVal Num As Integer, _
            ByVal Map As Integer, _
            ByVal X As Integer, _
            ByVal Y As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 11/5/2010
    '11/5/2010 - ZaMa: Arreglo bug que permitia apilar mas de 10k de items.
    '***************************************************
    '<EhHeader>
    On Error GoTo DropObj_Err

    '</EhHeader>

    Dim DropObj  As Obj

    Dim MapObj   As Obj

    Dim TempTick As Long
        
    With UserList(UserIndex)
        TempTick = GetTime

        If Num > 0 Then
        
            DropObj.ObjIndex = .Invent.Object(Slot).ObjIndex

            If ObjData(DropObj.ObjIndex).OBJType = eOBJType.otMonturas Then
                Call WriteConsoleMsg(UserIndex, "No puedes tirar la montura.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            If Not EsGmDios(UserIndex) Then
                If ObjData(DropObj.ObjIndex).NoNada = 1 Then
                    If ObjData(DropObj.ObjIndex).LvlMax > UserList(UserIndex).Stats.Elv Then
                        Call QuitarUserInvItem(UserIndex, Slot, DropObj.Amount)
                        Call UpdateUserInv(False, UserIndex, Slot)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes realizar ninguna acción con este objeto. ¡Podría ser de uso personal!", FontTypeNames.FONTTYPE_INFO)

                    End If
                        
                    Exit Sub
    
                End If

            End If
            
            If ObjData(DropObj.ObjIndex).NoDrop = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes dropear este objeto.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            If Not EsGm(UserIndex) Then
                If ObjData(DropObj.ObjIndex).Premium = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡No puedes tirar los objetos Premium!!", FontTypeNames.FONTTYPE_TALK)
                    
                    Exit Sub

                End If
            
                If ObjData(DropObj.ObjIndex).Oro = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡No puedes tirar los objetos Oro!!", FontTypeNames.FONTTYPE_TALK)
                    
                    Exit Sub

                End If
                
                If ObjData(DropObj.ObjIndex).Plata = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡No puedes tirar los objetos Plata!!", FontTypeNames.FONTTYPE_TALK)
                    
                    Exit Sub

                End If
                
                If ObjData(DropObj.ObjIndex).OBJType = otTransformVIP Then
                    Call WriteConsoleMsg(UserIndex, "¡¡No puedes tirar los skins!!", FontTypeNames.FONTTYPE_TALK)

                    Exit Sub

                End If
                
                If ObjData(DropObj.ObjIndex).Caos = 1 Or ObjData(DropObj.ObjIndex).Real = 1 Then
                    Call WriteConsoleMsg(UserIndex, "No puedes tirar objetos faccionarios.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If

            End If
            
            ' // NUEVO
            If .flags.SlotReto > 0 Then
                If Retos(.flags.SlotReto).config(eRetoConfig.eItems) = 1 Then
                    Call WriteConsoleMsg(UserIndex, "No puedes dropear objetos si estas luchando por los mismos", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If

            End If
            
            If .flags.SlotEvent > 0 Then
                If Events(.flags.SlotEvent).LimitRed > 0 Then
                    Call WriteConsoleMsg(UserIndex, "No puedes dropear objetos si estas luchando por límite de pociones rojas.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If

            End If
                
            If ObjData(DropObj.ObjIndex).OBJType = eOBJType.otBarcos Then
                Call WriteConsoleMsg(UserIndex, "¡¡ATENCIÓN!! ¡NO puedes tirar los barcos al suelo!", FontTypeNames.FONTTYPE_TALK)

                'Exit Sub
            End If
            
            DropObj.Amount = MinimoInt(Num, .Invent.Object(Slot).Amount)

            'Check objeto en el suelo
            MapObj.ObjIndex = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
            MapObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount
        
            If MapObj.ObjIndex = 0 Or MapObj.ObjIndex = DropObj.ObjIndex Then
        
                If MapObj.Amount = MAX_INVENTORY_OBJS Then
                    Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
            
                If DropObj.Amount + MapObj.Amount > MAX_INVENTORY_OBJS Then
                    DropObj.Amount = MAX_INVENTORY_OBJS - MapObj.Amount

                End If
            
                If Not ItemNewbie(DropObj.ObjIndex) Then Call MakeObj(DropObj, Map, X, Y)
                Call QuitarUserInvItem(UserIndex, Slot, DropObj.Amount)
                Call UpdateUserInv(False, UserIndex, Slot)
            
                If ObjData(DropObj.ObjIndex).OBJType = eOBJType.otGemas Then
                    If TempTick - .Counters.SpamMessage > 60000 Then
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El personaje '" & .Name & "' ha tirado la " & ObjData(DropObj.ObjIndex).Name & " en " & MapInfo(.Pos.Map).Name & "(Mapa: " & .Pos.Map & " " & .Pos.X & " " & .Pos.Y & ")", FontTypeNames.FONTTYPE_GUILD))
                        .Counters.SpamMessage = TempTick

                    End If

                End If
            
                If Not .flags.Privilegios And PlayerType.User Then
                    Call Logs_User(.Name, eGm, eDropObj, "tiró al piso " & DropObj.Amount & " " & ObjData(DropObj.ObjIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)

                End If

                If ObjData(DropObj.ObjIndex).Log = 1 Then
                    Call Logs_User(.Name, eLog.eUser, eDropObj, "tiró al piso " & DropObj.Amount & " " & ObjData(DropObj.ObjIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)
                
                ElseIf DropObj.Amount > 100 Then

                    If ObjData(DropObj.ObjIndex).NoLog <> 1 Then
                        Call Logs_User(.Name, eLog.eUser, eDropObj, "tiró al piso " & DropObj.Amount & " " & ObjData(DropObj.ObjIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)

                    End If

                End If

            Else
                Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

    '<EhFooter>
    Exit Sub

DropObj_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.DropObj " & "at line " & Erl

    '</EhFooter>
End Sub

Sub EraseObj(ByVal Num As Integer, _
             ByVal Map As Integer, _
             ByVal X As Integer, _
             ByVal Y As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo EraseObj_Err

    '</EhHeader>

    With MapData(Map, X, Y)
        .ObjInfo.Amount = .ObjInfo.Amount - Num
    
        If .ObjInfo.Amount <= 0 Then
            .ObjInfo.ObjIndex = 0
            .ObjInfo.Amount = 0
            .ObjEvent = 0
            
            Call ModAreas.DeleteEntity(ModAreas.Pack(Map, X, Y), ENTITY_TYPE_OBJECT)

        End If

    End With

    '<EhFooter>
    Exit Sub

EraseObj_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.EraseObj " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub MakeObj(ByRef Obj As Obj, _
            ByVal Map As Integer, _
            ByVal X As Integer, _
            ByVal Y As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo MakeObj_Err

    '</EhHeader>
    
    If Obj.ObjIndex > 0 And Obj.ObjIndex <= UBound(ObjData) Then
    
        With MapData(Map, X, Y)
            
            If .ObjInfo.ObjIndex = Obj.ObjIndex Then
                .ObjInfo.Amount = .ObjInfo.Amount + Obj.Amount
            Else
                .Protect = GetTime
                .ObjInfo = Obj
                
                If .trigger <> eTrigger.zonaOscura Then

                    Dim Coordinates As WorldPos

                    Coordinates.Map = Map
                    Coordinates.X = X
                    Coordinates.Y = Y
                
                    Call ModAreas.CreateEntity(ModAreas.Pack(Map, X, Y), ENTITY_TYPE_OBJECT, Coordinates, ObjData(.ObjInfo.ObjIndex).SizeWidth, ObjData(.ObjInfo.ObjIndex).SizeHeight)

                End If

            End If

        End With

    End If

    '<EhFooter>
    Exit Sub

MakeObj_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.MakeObj " & "at line " & Erl
        
    '</EhFooter>
End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, _
                               ByRef MiObj As Obj, _
                               Optional ByVal ShowMessage As Boolean = True) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim Slot As Byte
    
    With UserList(UserIndex)
        '¿el user ya tiene un objeto del mismo tipo?
        Slot = 1
        
        Do Until .Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
            Slot = Slot + 1

            If Slot > .CurrentInventorySlots Then
                Exit Do

            End If

        Loop
            
        'Sino busca un slot vacio
        If Slot > .CurrentInventorySlots Then
            Slot = 1

            Do Until .Invent.Object(Slot).ObjIndex = 0
                Slot = Slot + 1

                If Slot > .CurrentInventorySlots Then
                    If ShowMessage Then Call WriteConsoleMsg(UserIndex, "No puedes cargar más objetos.", FontTypeNames.FONTTYPE_FIGHT)
                    MeterItemEnInventario = False
                    Exit Function

                End If

            Loop

            .Invent.NroItems = .Invent.NroItems + 1

        End If
    
        If Slot > MAX_NORMAL_INVENTORY_SLOTS And Slot <= MAX_INVENTORY_SLOTS Then
            If Not ItemSeCae(MiObj.ObjIndex) Then
                If ShowMessage Then Call WriteConsoleMsg(UserIndex, "No puedes contener objetos especiales en tu " & ObjData(.Invent.MochilaEqpObjIndex).Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                MeterItemEnInventario = False
                Exit Function

            End If

        End If

        'Mete el objeto
        If .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
            .Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
            .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + MiObj.Amount
        Else
            .Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS

        End If

    End With
    
    MeterItemEnInventario = True
           
    Call UpdateUserInv(False, UserIndex, Slot)
    
    Exit Function

ErrHandler:
    Call LogError("Error en MeterItemEnInventario. Error " & Err.number & " : " & Err.description)

End Function

Sub GetObj(ByVal UserIndex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 18/12/2009
    '18/12/2009: ZaMa - Oro directo a la billetera.
    '***************************************************
    '<EhHeader>
    On Error GoTo GetObj_Err

    '</EhHeader>

    Dim Obj    As ObjData

    Dim MiObj  As Obj

    Dim ObjPos As String

    With UserList(UserIndex)

        '¿Hay algun obj?
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex > 0 Then
            If Not EsGm(UserIndex) Then
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 1 Then Exit Sub

            End If
            
            '¿Esta permitido agarrar este obj?
            If ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then

                Dim X As Integer

                Dim Y As Integer
                
                X = .Pos.X
                Y = .Pos.Y
                
                Obj = ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex)
                MiObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount
                MiObj.ObjIndex = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                
                MapData(.Pos.Map, .Pos.X, .Pos.Y).Protect = 0
                
                ' Oro directo a la billetera!
                'If Obj.OBJType = otGuita Then
                ' .Stats.Gld = .Stats.Gld + MiObj.Amount
                'Quitamos el objeto
                'Call EraseObj(MapData(.Pos.Map, X, Y).ObjInfo.Amount, .Pos.Map, .Pos.X, .Pos.Y)
                        
                ' Call WriteUpdateGold(UserIndex)
                'Else
                If MeterItemEnInventario(UserIndex, MiObj) Then
                         
                    'Quitamos el objeto
                    Call EraseObj(MapData(.Pos.Map, X, Y).ObjInfo.Amount, .Pos.Map, .Pos.X, .Pos.Y)

                    ' Comprobamos si esta en una misión
                    Call Quests_Check_Objs(UserIndex, MiObj.ObjIndex, MiObj.Amount)
                          
                    If Not .flags.Privilegios And PlayerType.User Then
                        Call Logs_User(.Name, eGm, eGetObj, .Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)

                    End If
        
                    If ObjData(MiObj.ObjIndex).Log = 1 Then
                        ObjPos = " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y
                        Call Logs_User(.Name, eLog.eUser, eGetObj, .Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)
                            
                    ElseIf MiObj.Amount > 100 Then

                        If ObjData(MiObj.ObjIndex).NoLog <> 1 Then
                            ObjPos = " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y
                            Call Logs_User(.Name, eLog.eUser, eGetObj, .Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)

                        End If

                    End If

                End If

                'End If
            End If

        Else
            Call WriteConsoleMsg(UserIndex, "No hay nada aquí.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    '<EhFooter>
    Exit Sub

GetObj_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.GetObj " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
    '***************************************************
    'Author: Unknown
    'Last Modification: 26/05/2011
    '26/05/2011: Amraphen - Agregadas armaduras faccionarias de segunda jerarquía.
    '***************************************************

    On Error GoTo ErrHandler

    'Desequipa el item slot del inventario
    Dim Obj As ObjData
    
    With UserList(UserIndex)
        With .Invent

            If (Slot < LBound(.Object)) Or (Slot > UBound(.Object)) Then

                Exit Sub

            ElseIf .Object(Slot).ObjIndex = 0 Then

                Exit Sub

            End If
            
            Obj = ObjData(.Object(Slot).ObjIndex)

        End With
        
        If Obj.SkillNum > 0 Or Obj.SkillsEspecialNum > 0 Then
            Call UserStats_UpdateEffectAll(UserIndex, Obj, False)

        End If
        
        Select Case Obj.OBJType
            
            Case eOBJType.otMonturas
                Call DoEquita(UserIndex, Obj, Slot)
                
                With .Invent
                    .Object(Slot).Equipped = 0
                    .MonturaObjIndex = 0
                    .MonturaSlot = 0

                End With
                
            Case eOBJType.otReliquias
                ' mEffect.Effect_UpdateUser UserIndex, True
                
                With .Invent
                    .Object(Slot).Equipped = 0
                    .ReliquiaObjIndex = 0
                    .ReliquiaSlot = 0

                End With
                
            Case eOBJType.otPendienteParty

                ' # Actualizar los porcentajes
                If .GroupIndex > 0 Then
                    UpdatePorcentaje .GroupIndex

                End If
                
                With .Invent
                    .Object(Slot).Equipped = 0
                    .PendientePartyObjIndex = 0
                    .PendientePartySlot = 0

                End With
                
            Case eOBJType.otMagic

                With .Invent
                    .Object(Slot).Equipped = 0
                    .MagicObjIndex = 0
                    .MagicSlot = 0

                End With
                
            Case eOBJType.otWeapon

                With .Invent
                    .Object(Slot).Equipped = 0
                    .WeaponEqpObjIndex = 0
                    .WeaponEqpSlot = 0

                End With
                
                '.Skins.WeaponIndex = 0
                 
                If Not .flags.Mimetizado = 1 Then

                    With .Char
                        .AuraIndex(2) = 0
                        .WeaponAnim = NingunArma
                        Call ChangeUserChar(UserIndex, .Body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraIndex)

                    End With

                End If
                
            Case eOBJType.otAuras

                With .Invent
                    .Object(Slot).Equipped = 0
                    .AuraEqpObjIndex = 0
                    .AuraEqpSlot = 0

                End With
            
                With .Char
                    .AuraIndex(5) = 0
                    Call ChangeUserChar(UserIndex, .Body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraIndex)

                End With
                
            Case eOBJType.otFlechas

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
                
                If .flags.TransformVIP > 0 Then
                    Call TransformVIP_User(UserIndex, 0)

                End If

                With .Invent

                    'Si tiene armadura faccionaria de segunda jerarquía equipada la sacamos:
                    If .FactionArmourEqpObjIndex Then
                        Call Desequipar(UserIndex, .FactionArmourEqpSlot)

                    End If
                    
                    .Object(Slot).Equipped = 0
                    .ArmourEqpObjIndex = 0
                    .ArmourEqpSlot = 0
                    
                End With
                
                '.Skins.ArmourIndex = 0
                
                If .flags.Navegando = 0 Then
                    Call DarCuerpoDesnudo(UserIndex, .flags.Mimetizado = 1)

                End If
                
                With .Char
                    .AuraIndex(1) = 0
                    Call ChangeUserChar(UserIndex, .Body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraIndex)

                End With
                 
            Case eOBJType.otcasco

                With .Invent
                    .Object(Slot).Equipped = 0
                    .CascoEqpObjIndex = 0
                    .CascoEqpSlot = 0

                End With
                
                ' .Skins.HelmIndex = 0
                
                If Not .flags.Mimetizado = 1 Then

                    With .Char
                        .AuraIndex(3) = 0
                        .CascoAnim = NingunCasco
                        Call ChangeUserChar(UserIndex, .Body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraIndex)

                    End With

                End If
            
            Case eOBJType.otescudo

                With .Invent
                    .Object(Slot).Equipped = 0
                    .EscudoEqpObjIndex = 0
                    .EscudoEqpSlot = 0

                End With
                
                ' .Skins.ShieldIndex = 0
                 
                If Not .flags.Mimetizado = 1 Then

                    With .Char
                        .AuraIndex(4) = 0
                        .ShieldAnim = NingunEscudo
                        Call ChangeUserChar(UserIndex, .Body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraIndex)

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
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Desquipar. Error " & Err.number & " : " & Err.description)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, _
                           ByVal ObjIndex As Integer, _
                           Optional ByRef sMotivo As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '***************************************************

    On Error GoTo ErrHandler
    
    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
        If ObjData(ObjIndex).Mujer = 1 Then
            SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Hombre
        ElseIf ObjData(ObjIndex).Hombre = 1 Then
            SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Mujer
        Else
            SexoPuedeUsarItem = True

        End If
        
    Else
        SexoPuedeUsarItem = True

    End If
    
    If Not SexoPuedeUsarItem Then sMotivo = "Tu género no puede usar este objeto."

    Exit Function

ErrHandler:
    Call LogError("SexoPuedeUsarItem")

End Function

Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, _
                              ByVal ObjIndex As Integer, _
                              Optional ByRef sMotivo As String) As Boolean

    '<EhHeader>
    On Error GoTo FaccionPuedeUsarItem_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: 26/05/2011 (Amraphen)
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '26/05/2011: Amraphen - Agrego validación para armaduras faccionarias de segunda jerarquía.
    '***************************************************
    Dim ArmourIndex           As Integer

    Dim FaltaPrimeraJerarquia As Boolean

    If ObjData(ObjIndex).Real Then
        If Not Escriminal(UserIndex) And esArmada(UserIndex) Then
            If ObjData(ObjIndex).Real = 2 Then
                ArmourIndex = UserList(UserIndex).Invent.ArmourEqpObjIndex
                
                If ArmourIndex > 0 And ObjData(ArmourIndex).Real = 1 Then
                    FaccionPuedeUsarItem = True
                Else
                    FaccionPuedeUsarItem = False
                    FaltaPrimeraJerarquia = True

                End If

            Else 'Es item faccionario común
                FaccionPuedeUsarItem = True

            End If

        Else
            FaccionPuedeUsarItem = False

        End If

    ElseIf ObjData(ObjIndex).Caos Then

        If Escriminal(UserIndex) And esCaos(UserIndex) Then
            If ObjData(ObjIndex).Caos = 2 Then
                ArmourIndex = UserList(UserIndex).Invent.ArmourEqpObjIndex
                
                If ArmourIndex > 0 And ObjData(ArmourIndex).Caos = 1 Then
                    FaccionPuedeUsarItem = True
                Else
                    FaccionPuedeUsarItem = False
                    FaltaPrimeraJerarquia = True

                End If

            Else 'Es item faccionario común
                FaccionPuedeUsarItem = True

            End If

        Else
            FaccionPuedeUsarItem = False

        End If

    Else
        FaccionPuedeUsarItem = True

    End If
    
    If Not FaccionPuedeUsarItem Then
        If FaltaPrimeraJerarquia Then
            sMotivo = "Debes tener equipada una armadura faccionaria."
        Else
            sMotivo = "Tu alinación no puede usar este objeto."

        End If

    End If

    '<EhFooter>
    Exit Function

FaccionPuedeUsarItem_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.FaccionPuedeUsarItem " & "at line " & Erl
        
    '</EhFooter>
End Function

Private Function CheckUserSkill(ByVal UserIndex As Integer, _
                                ByRef Obj As ObjData) As Boolean

    '<EhHeader>
    On Error GoTo CheckUserSkill_Err

    '</EhHeader>

    With UserList(UserIndex)

        If .Stats.UserSkills(eSkill.Magia) < Obj.MagiaSkill Then
            Call WriteConsoleMsg(UserIndex, "Para poder utilizar este ítem es necesario tener " & Obj.MagiaSkill & " skills en Mágia.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If

        If .Stats.UserSkills(eSkill.Resistencia) < Obj.RMSkill Then
            Call WriteConsoleMsg(UserIndex, "Para poder utilizar este ítem es necesario tener " & Obj.RMSkill & " skills en Resistencia Mágica.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If

        If Obj.OBJType = otWeapon Then
            If .Stats.UserSkills(eSkill.Armas) < Obj.ArmaSkill Then
                Call WriteConsoleMsg(UserIndex, "Para usar este ítem tienes que tener " & Obj.ArmaSkill & " skills en Combate con Armas.", FontTypeNames.FONTTYPE_INFO)

                Exit Function

            End If

        End If

        If .Stats.UserSkills(eSkill.Defensa) < Obj.EscudoSkill Then
            Call WriteConsoleMsg(UserIndex, "Para usar este ítem tienes que tener " & Obj.EscudoSkill & " skills en Defensa con Escudos.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If

        If .Stats.UserSkills(eSkill.Tacticas) < Obj.ArmaduraSkill Then
            Call WriteConsoleMsg(UserIndex, "Para usar este ítem tienes que tener " & Obj.ArmaduraSkill & " skills en Evasión.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If

        If Obj.OBJType = otWeapon Then
            If .Stats.UserSkills(eSkill.Proyectiles) < Obj.ArcoSkill Then
                Call WriteConsoleMsg(UserIndex, "Para usar este item tienes que tener " & Obj.ArcoSkill & " skills en Armas de Proyectiles.", FontTypeNames.FONTTYPE_INFO)

                Exit Function

            End If

        End If

        If .Stats.UserSkills(eSkill.Apuñalar) < Obj.DagaSkill Then
            Call WriteConsoleMsg(UserIndex, "Para utilizar este ítem necesitas " & Obj.DagaSkill & " skills en Apuñalar.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If
            
        If .Stats.UserSkills(eSkill.Magia) < Obj.MagiaSkill Then
            Call WriteConsoleMsg(UserIndex, "Para usar este item tienes que tener " & Obj.MagiaSkill & " skills en Magia.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If
        
        CheckUserSkill = True
    
    End With

    '<EhFooter>
    Exit Function

CheckUserSkill_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.CheckUserSkill " & "at line " & Erl

    '</EhFooter>
End Function

' @ Aplicamos /Des aplicamops los atributos del objeto
Sub UserStats_UpdateEffectAll(ByVal UserIndex As Integer, _
                              ByRef Obj As ObjData, _
                              ByVal Equipped As Boolean)

    '<EhHeader>
    On Error GoTo UserStats_UpdateEffectAll_Err

    '</EhHeader>
    
    Dim A          As Long

    Dim SkillIndex As Integer

    Dim Amount     As Integer
    
    With Obj
    
        If .SkillNum > 0 Then

            For A = 1 To .SkillNum
                SkillIndex = .Skill(A).Selected
                Amount = IIf(Equipped, .Skill(A).Amount, -.Skill(A).Amount)
                UserList(UserIndex).Stats.UserSkills(.Skill(A).Selected) = UserList(UserIndex).Stats.UserSkills(.Skill(A).Selected) + Amount
            Next A

        End If
        
        If .SkillsEspecialNum > 0 Then

            For A = 1 To .SkillsEspecialNum
                SkillIndex = .SkillsEspecial(A).Selected
                Amount = IIf(Equipped, .SkillsEspecial(A).Amount, -.SkillsEspecial(A).Amount)
                UserList(UserIndex).Stats.UserSkillsEspecial(SkillIndex) = UserList(UserIndex).Stats.UserSkillsEspecial(SkillIndex) + Amount
                Call UserStats_UpdateEffectUser(UserIndex, SkillIndex, Amount)
            Next A
        
        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

UserStats_UpdateEffectAll_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.UserStats_UpdateEffectAll " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Sub UserStats_UpdateEffectUser(ByVal UserIndex As Integer, _
                               ByVal SkillNum As Byte, _
                               ByVal Amount As Integer)

    '<EhHeader>
    On Error GoTo UserStats_UpdateEffectUser_Err

    '</EhHeader>
    
    With UserList(UserIndex)

        Select Case SkillNum
        
            Case 1 ' Vida
                .Stats.MaxHp = .Stats.MaxHp + Amount
                .Stats.MinHp = .Stats.MaxHp
                Call WriteUpdateUserStats(UserIndex)
                    
            Case 2 ' Maná
                .Stats.MaxMan = .Stats.MaxMan + Amount
                .Stats.MinMan = .Stats.MaxMan
                Call WriteUpdateUserStats(UserIndex)
                    
            Case 3 'Curación : Skill que define un porcentaje (1 a 100)
            
            Case 4 'Escudo Mágico : Skill que define un porcentaje (1 a 100)
            
            Case 5 'Veneno : Skill que define un porcentaje (1 a 100)
            
            Case 6 'Fuerza
                .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + Amount

                If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                If .Stats.UserAtributos(eAtributos.Fuerza) > 2 * .Stats.UserAtributosBackUP(Fuerza) Then .Stats.UserAtributos(eAtributos.Fuerza) = 2 * .Stats.UserAtributosBackUP(Fuerza)
                Call WriteUpdateStrenght(UserIndex)
                        
            Case 7 'Agilidad
                .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + Amount

                If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                If .Stats.UserAtributos(eAtributos.Agilidad) > 2 * .Stats.UserAtributosBackUP(Agilidad) Then .Stats.UserAtributos(eAtributos.Agilidad) = 2 * .Stats.UserAtributosBackUP(Fuerza)
                Call WriteUpdateDexterity(UserIndex)

        End Select
    
    End With

    '<EhFooter>
    Exit Sub

UserStats_UpdateEffectUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.UserStats_UpdateEffectUser " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
    '*************************************************
    'Author: Unknown
    'Last modified: 26/05/2011 (Amraphen)
    '01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin
    '14/01/2010: ZaMa - Agrego el motivo especifico por el que no puede equipar/usar el item.
    '26/05/2011: Amraphen - Agregadas armaduras faccionarias de segunda jerarquía.
    '*************************************************

    On Error GoTo ErrHandler

    'Equipa un item del inventario
    Dim Obj      As ObjData

    Dim ObjIndex As Integer

    Dim sMotivo  As String
    
    With UserList(UserIndex)
        ObjIndex = .Invent.Object(Slot).ObjIndex
        Obj = ObjData(ObjIndex)
        
        If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
                
        If Obj.Bronce = 1 And Not .flags.Bronce = 1 Then
            Call WriteConsoleMsg(UserIndex, "Sólo los usuarios [AVENTURERO] pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If Obj.Plata = 1 And Not .flags.Plata = 1 Then
            Call WriteConsoleMsg(UserIndex, "Sólo los usuarios [HEROE] pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If Obj.Oro = 1 And Not .flags.Oro = 1 Then
            Call WriteConsoleMsg(UserIndex, "Sólo los usuarios [LEYENDA] pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If Obj.Premium = 1 And Not .flags.Premium = 1 Then
            Call WriteConsoleMsg(UserIndex, "Sólo los usuarios PREMIUM pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If Obj.Navidad = 1 And ModoNavidad = 0 Then
            Call WriteConsoleMsg(UserIndex, "Ya supera navidad wei", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
                
        If Obj.LvlMax <> 0 And Obj.LvlMax < .Stats.Elv Then
            Call WriteConsoleMsg(UserIndex, "Sólo puedes usar este objeto hasta Nivel '" & Obj.LvlMax & "'.", FontTypeNames.FONTTYPE_USERPREMIUM)
            Exit Sub

        End If
        
        If Obj.LvlMin <> 0 And Obj.LvlMin > .Stats.Elv Then
            Call WriteConsoleMsg(UserIndex, "Sólo puedes usar este objeto a partir del Nivel '" & Obj.LvlMax & "'.", FontTypeNames.FONTTYPE_USERPREMIUM)
            Exit Sub

        End If
        
        ' Skill requerido para el objeto
        If Not CheckUserSkill(UserIndex, Obj) Then Exit Sub
        
        If .flags.SlotReto > 0 Then
        
            ' Uso de Escudos/Cascos
            If (Retos(.flags.SlotReto).config(eRetoConfig.eEscudos) = 0 And Obj.OBJType = otescudo) Or (Retos(.flags.SlotReto).config(eRetoConfig.eCascos) = 0 And Obj.OBJType = otcasco) Then
                Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite el uso de este objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            ' Uso de Objetos [BRONCE] [PLATA] [ORO] [PREMIUM]
            'If (Retos(.flags.SlotReto).config(eRetoConfig.eBronce) = 0 And .flags.Bronce = 0) Or (Retos(.flags.SlotReto).config(eRetoConfig.ePlata) = 0 And .flags.Plata = 0) Or (Retos(.flags.SlotReto).config(eRetoConfig.eOro) = 0 And .flags.Oro = 0) Or (Retos(.flags.SlotReto).config(eRetoConfig.ePremium) = 0 And .flags.Premium = 0) Then
                
            'Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite el uso de este objeto.", FontTypeNames.FONTTYPE_INFO)

            ' Exit Sub

            ' End If
            
        End If
        
        If .flags.SlotEvent > 0 Then
            If Events(.flags.SlotEvent).Modality = eModalityEvent.DagaRusa Then
                Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite el uso de objetos.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If

        End If

        Select Case Obj.OBJType

            Case eOBJType.otMagic

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                    
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)

                        Exit Sub

                    End If
                    
                    'Quitamos el elemento anterior
                    If .Invent.MagicObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.MagicSlot)

                    End If
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.MagicObjIndex = ObjIndex
                    .Invent.MagicSlot = Slot
                    
                Else
                    
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If

            Case eOBJType.otReliquias

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                    
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)

                        Exit Sub

                    End If
                    
                    'Quitamos el elemento anterior
                    If .Invent.ReliquiaObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.ReliquiaSlot)

                    End If
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.ReliquiaObjIndex = ObjIndex
                    .Invent.ReliquiaSlot = Slot
                    
                    'Call mEffect.Effect_UpdateUser(UserIndex, False)
                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
                
            Case eOBJType.otPendienteParty

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                    
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)

                        Exit Sub

                    End If
                    
                    'Quitamos el elemento anterior
                    If .Invent.PendientePartyObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.PendientePartySlot)

                    End If
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.PendientePartyObjIndex = ObjIndex
                    .Invent.PendientePartySlot = Slot
                    
                    Call WriteConsoleMsg(UserIndex, "En caso de que seas líder de un grupo podrás cambiar el porcentaje hasta " & ObjData(ObjIndex).Porc & "%", FontTypeNames.FONTTYPE_INFOGREEN)
                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
                
            Case eOBJType.otAuras

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                     
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)

                        'Animacion por defecto
                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.AuraIndex(5) = NingunAura
                        Else
                            .Char.AuraIndex(5) = NingunAura
                            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                        End If

                        Exit Sub

                    End If
                     
                    'Quitamos el elemento anterior
                    If .Invent.AuraEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.AuraEqpSlot)

                    End If
                     
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.AuraEqpObjIndex = ObjIndex
                    .Invent.AuraEqpSlot = Slot
                     
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.WeaponAnim = ObjData(ObjIndex).AuraIndex
                        .CharMimetizado.AuraIndex(5) = ObjData(ObjIndex).AuraIndex
                    Else
                        .Char.AuraIndex(5) = ObjData(ObjIndex).AuraIndex
                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
                
            Case eOBJType.otWeapon

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                    
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)

                        'Animacion por defecto
                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.WeaponAnim = NingunArma
                        Else
                            .Char.WeaponAnim = NingunArma
                            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                        End If

                        Exit Sub
                    Else

                        ' Quiere equipar un arma dos manos y tiene escudo.
                        If .Invent.EscudoEqpObjIndex > 0 Then
                            If ObjData(ObjIndex).DosManos = 1 Then
                                Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)

                            End If

                        End If

                    End If
                    
                    'Quitamos el elemento anterior
                    If .Invent.WeaponEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)

                    End If
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.WeaponEqpObjIndex = ObjIndex
                    .Invent.WeaponEqpSlot = Slot
                    
                    'El sonido solo se envia si no lo produce un admin invisible
                    If Not (.flags.AdminInvisible = 1) Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_SACARARMA, .Pos.X, .Pos.Y, .Char.charindex))
                    
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.WeaponAnim = GetWeaponAnim(UserIndex, .Raza, ObjIndex)
                        .CharMimetizado.AuraIndex(2) = ObjData(ObjIndex).AuraIndex(2)
                    Else
                        .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Raza, ObjIndex)
                        .Char.AuraIndex(2) = ObjData(ObjIndex).AuraIndex(2)
                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
            
                'Call Skins_CheckObj(UserIndex, ObjIndex)

            Case eOBJType.otAnillo

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then

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
                    .Invent.AnilloEqpObjIndex = ObjIndex
                    .Invent.AnilloEqpSlot = Slot
                        
                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eOBJType.otFlechas

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                        
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
                    .Invent.MunicionEqpObjIndex = ObjIndex
                    .Invent.MunicionEqpSlot = Slot
                        
                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eOBJType.otarmadura

                If .flags.Navegando = 1 Then Exit Sub
                If .flags.Montando = 1 Then Exit Sub
                
                'Nos aseguramos que puede usarla
                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And SexoPuedeUsarItem(UserIndex, ObjIndex, sMotivo) And CheckRazaUsaRopa(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                    
                    'Nos fijamos si es armadura de segunda jerarquia
                    If Obj.Real = 2 Or Obj.Caos = 2 Then

                        'Si esta equipado lo quita
                        If .Invent.Object(Slot).Equipped Then
                            Call Desequipar(UserIndex, Slot)
                            
                            If Not .flags.Mimetizado = 1 Then
                                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                            End If
                            
                            Exit Sub

                        End If
                        
                        'Quita el anterior
                        If .Invent.FactionArmourEqpObjIndex > 0 Then
                            Call Desequipar(UserIndex, .Invent.FactionArmourEqpSlot)

                        End If
                        
                        'Lo equipa
                        .Invent.Object(Slot).Equipped = 1
                        .Invent.FactionArmourEqpObjIndex = ObjIndex
                        .Invent.FactionArmourEqpSlot = Slot
                        
                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.Body = GetArmourAnim(UserIndex, ObjIndex)
                            
                        End If

                    Else

                        'Si esta equipado lo quita
                        If .Invent.Object(Slot).Equipped Then
                            Call Desequipar(UserIndex, Slot)
                            
                            'Esto está de más:
                            'Call DarCuerpoDesnudo(UserIndex, .flags.Mimetizado = 1)
                            If Not .flags.Mimetizado = 1 Then
                                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                            End If
                            
                            Exit Sub

                        End If
                
                        'Quita el anterior
                        If .Invent.ArmourEqpObjIndex > 0 Then
                            Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)

                        End If
                
                        'Lo equipa
                        .Invent.Object(Slot).Equipped = 1
                        .Invent.ArmourEqpObjIndex = ObjIndex
                        .Invent.ArmourEqpSlot = Slot
                            
                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.Body = GetArmourAnim(UserIndex, ObjIndex)
                            .CharMimetizado.AuraIndex(1) = ObjData(ObjIndex).AuraIndex(1)
                        Else
                            .Char.Body = GetArmourAnim(UserIndex, ObjIndex)
                            .Char.AuraIndex(1) = ObjData(ObjIndex).AuraIndex(1)
                            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                        End If

                        .flags.Desnudo = 0

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
                
                'Call Skins_CheckObj(UserIndex, ObjIndex)
                
            Case eOBJType.otcasco

                If .flags.Navegando = 1 Then Exit Sub
                
                If .flags.SlotEvent > 0 Then
                    If Events(.flags.SlotEvent).config(eConfigEvent.eCascoEscudo) = 0 Then Exit Sub

                End If
        
                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then

                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)

                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.CascoAnim = NingunCasco
                        Else
                            .Char.CascoAnim = NingunCasco
                            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                        End If

                        Exit Sub

                    End If
            
                    'Quita el anterior
                    If .Invent.CascoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.CascoEqpSlot)

                    End If
            
                    'Lo equipa
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.CascoEqpObjIndex = ObjIndex
                    .Invent.CascoEqpSlot = Slot

                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.CascoAnim = GetHelmAnim(UserIndex, .Invent.CascoEqpObjIndex)
                        .CharMimetizado.AuraIndex(3) = ObjData(ObjIndex).AuraIndex(3)
                    Else
                        .Char.CascoAnim = GetHelmAnim(UserIndex, .Invent.CascoEqpObjIndex)
                        .Char.AuraIndex(3) = ObjData(ObjIndex).AuraIndex(3)
                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
                
                'Call Skins_CheckObj(UserIndex, ObjIndex)

            Case eOBJType.otescudo

                If .flags.Navegando = 1 Then Exit Sub
                If .flags.SlotEvent > 0 Then
                    If Events(.flags.SlotEvent).config(eConfigEvent.eCascoEscudo) = 0 Then Exit Sub

                End If
                
                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
        
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)

                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.ShieldAnim = NingunEscudo
                        Else
                            .Char.ShieldAnim = NingunEscudo
                            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                        End If

                        Exit Sub
                        
                    Else

                        ' Quiere equipar un escudo y tiene arma dos manos
                        If .Invent.WeaponEqpObjIndex > 0 Then
                            If ObjData(.Invent.WeaponEqpObjIndex).DosManos = 1 Then
                                Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)

                            End If

                        End If

                    End If
             
                    'Quita el anterior
                    If .Invent.EscudoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)

                    End If
             
                    'Lo equipa
                     
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.EscudoEqpObjIndex = ObjIndex
                    .Invent.EscudoEqpSlot = Slot
                     
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.ShieldAnim = GetShieldAnim(UserIndex, .Invent.EscudoEqpObjIndex)
                        .CharMimetizado.AuraIndex(4) = ObjData(ObjIndex).AuraIndex(4)
                    Else
                        .Char.ShieldAnim = GetShieldAnim(UserIndex, ObjIndex)
                        .Char.AuraIndex(4) = ObjData(ObjIndex).AuraIndex(4)
                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
                 
                'Call Skins_CheckObj(UserIndex, ObjIndex)
    
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
                .Invent.MochilaEqpObjIndex = ObjIndex
                .Invent.MochilaEqpSlot = Slot
                .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + Obj.MochilaType * 5
                Call WriteAddSlots(UserIndex, Obj.MochilaType)

        End Select
    
    End With

    ' Agrega los Atributos necesarios segun los skills del objeto.
    If Obj.SkillNum > 0 Or Obj.SkillsEspecialNum > 0 Then
        Call UserStats_UpdateEffectAll(UserIndex, Obj, True)

    End If
        
    'Actualiza
    Call UpdateUserInv(False, UserIndex, Slot)
    
    Exit Sub
    
ErrHandler:
    Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.number & " - Error Description : " & Err.description)

End Sub

Public Function CheckRazaUsaRopa(ByVal UserIndex As Integer, _
                                 ItemIndex As Integer, _
                                 Optional ByRef sMotivo As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        'Verifica si la raza puede usar la ropa
        If .Raza = eRaza.Humano Or .Raza = eRaza.Elfo Or .Raza = eRaza.Drow Then
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
        Else

            If (ObjData(ItemIndex).RopajeEnano <> 0) Then
                CheckRazaUsaRopa = True
            Else
                CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)

            End If

        End If
        
        'Solo se habilita la ropa exclusiva para Drows por ahora. Pablo (ToxicWaste)
        If (.Raza <> eRaza.Drow) And ObjData(ItemIndex).RazaDrow Then
            CheckRazaUsaRopa = False

        End If

    End With
    
    If EsGm(UserIndex) Then CheckRazaUsaRopa = True
    
    If Not CheckRazaUsaRopa Then sMotivo = "Tu raza no puede usar este objeto."
    
    Exit Function
    
ErrHandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Private Sub Potion_SimulatePotion(ByVal UserIndex As Integer, ByVal Slot As Byte)

    '<EhHeader>
    On Error GoTo Potion_SimulatePotion_Err

    '</EhHeader>
                                  
    Dim TempTick As Long
    
    With UserList(UserIndex)
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        Call UpdateUserInv(False, UserIndex, Slot)

        ' Los admin invisibles solo producen sonidos a si mismos
        If .flags.AdminInvisible = 1 Then
            Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
        Else

            If TempTick - .Counters.RuidoPocion > 1000 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                .Counters.RuidoPocion = TempTick

            End If

        End If

    End With

    '<EhFooter>
    Exit Sub

Potion_SimulatePotion_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.Potion_SimulatePotion " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub UseInvItem(ByVal UserIndex As Integer, _
               ByVal Slot As Byte, _
               ByVal SecondaryClick As Byte, _
               ByVal Value As Long)

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
    '<EhHeader>
    On Error GoTo UseInvItem_Err

    '</EhHeader>

    Dim Obj      As ObjData

    Dim ObjIndex As Integer

    Dim TargObj  As ObjData

    Dim MiObj    As Obj

    Dim sMotivo  As String
    
    With UserList(UserIndex)
    
        If .Invent.Object(Slot).Amount = 0 Then Exit Sub
        
        Obj = ObjData(.Invent.Object(Slot).ObjIndex)

        If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFORED)

            Exit Sub

        End If
        
        If Not ClasePuedeUsarItem(UserIndex, .Invent.Object(Slot).ObjIndex, sMotivo) Then
            Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFORED)

            Exit Sub

        End If
        
        If Obj.OBJType = otTransformVIP Then
            If Not CheckRazaUsaRopa(UserIndex, .Invent.Object(Slot).ObjIndex, sMotivo) Then
                Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFORED)
    
                Exit Sub
    
            End If

        End If
        
        If Obj.Bronce = 1 And Not .flags.Bronce = 1 Then
            Call WriteConsoleMsg(UserIndex, "Sólo los usuarios [AVENTURERO] pueden usar estos objetos.", FontTypeNames.FONTTYPE_USERBRONCE)

            Exit Sub

        End If
        
        If Obj.Plata = 1 And Not .flags.Plata = 1 Then
            Call WriteConsoleMsg(UserIndex, "Sólo los usuarios [HEROE] pueden usar estos objetos.", FontTypeNames.FONTTYPE_USERPLATA)

            Exit Sub

        End If
        
        If Obj.Oro = 1 And Not .flags.Oro = 1 Then
            Call WriteConsoleMsg(UserIndex, "Sólo los usuarios [LEYENDA] pueden usar estos objetos.", FontTypeNames.FONTTYPE_USERGOLD)

            Exit Sub

        End If
        
        If Obj.Premium = 1 And Not .flags.Premium = 1 Then
            Call WriteConsoleMsg(UserIndex, "Sólo los usuarios [PREMIUM] pueden usar estos objetos.", FontTypeNames.FONTTYPE_USERPREMIUM)

            Exit Sub

        End If
        
        If Obj.LvlMax <> 0 And Obj.LvlMax < .Stats.Elv Then
            Call WriteConsoleMsg(UserIndex, "Sólo puedes usar este objeto hasta Nivel '" & Obj.LvlMax & "'.", FontTypeNames.FONTTYPE_USERPREMIUM)
            Exit Sub

        End If
        
        If Obj.LvlMin <> 0 And Obj.LvlMin > .Stats.Elv Then
            Call WriteConsoleMsg(UserIndex, "Sólo puedes usar este objeto a partir del Nivel '" & Obj.LvlMax & "'.", FontTypeNames.FONTTYPE_USERPREMIUM)
            Exit Sub

        End If
            
        If Obj.OBJType = eOBJType.otWeapon Then
            If Obj.proyectil = 1 Then
                
                'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
                If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
                
                If Obj.Municion = 1 Then
                    If .Invent.MunicionEqpObjIndex = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Debes equipar las municiones antes de usar el arma de proyectil.", FontTypeNames.FONTTYPE_USERPREMIUM)

                        Exit Sub

                    End If

                End If

            Else

                'dagas
                If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub

            End If

        Else

            If SecondaryClick Then
                If Not IntervaloPermiteUsarClick(UserIndex) Then Exit Sub
            Else

                If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub

            End If
           
        End If
        
        If .flags.Meditando Then
            .flags.Meditando = False
            .Char.FX = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))

        End If
            
        ObjIndex = .Invent.Object(Slot).ObjIndex
        .flags.TargetObjInvIndex = ObjIndex
        .flags.TargetObjInvSlot = Slot
        
        Select Case Obj.OBJType
                
            Case eOBJType.otItemRandom
                    
                Call Chest_AbreFortuna(UserIndex, ObjIndex)
                    
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                Call UpdateUserInv(False, UserIndex, Slot)
                    
            Case eOBJType.otcofre
                
                Call Chest_DropObj(UserIndex, ObjIndex, 0, 0, 0, True)
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                Call UpdateUserInv(False, UserIndex, Slot)

            Case eOBJType.otPociones

                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo. ", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                
                If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then Exit Sub
                
                Dim TempTick As Long, CanUse As Boolean

                .flags.TomoPocion = True
                .flags.TipoPocion = Obj.TipoPocion
                
                TempTick = GetTime
                CanUse = True
                
                Select Case .flags.TipoPocion
                
                    Case 1 'Modif la agilidad
                        .flags.DuracionEfecto = Obj.DuracionEfecto
                
                        'Usa el item
                        .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)

                        If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS

                        If .Stats.UserAtributos(eAtributos.Agilidad) > 2 * .Stats.UserAtributosBackUP(Agilidad) Then .Stats.UserAtributos(eAtributos.Agilidad) = 2 * .Stats.UserAtributosBackUP(Agilidad)
                        
                        'Quitamos del inv el item
                        If Obj.Ilimitado = 0 Then
                            Call QuitarUserInvItem(UserIndex, Slot, 1)

                        End If
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                            
                        Else

                            If TempTick - .Counters.RuidoPocion > 1000 Then
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                                .Counters.RuidoPocion = TempTick

                            End If

                        End If

                        Call WriteUpdateDexterity(UserIndex)
                        
                    Case 2 'Modif la fuerza
                        .flags.DuracionEfecto = Obj.DuracionEfecto
                
                        'Usa el item
                        .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)

                        If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS

                        If .Stats.UserAtributos(eAtributos.Fuerza) > 2 * .Stats.UserAtributosBackUP(Fuerza) Then .Stats.UserAtributos(eAtributos.Fuerza) = 2 * .Stats.UserAtributosBackUP(Fuerza)
                        
                        'Quitamos del inv el item
                        If Obj.Ilimitado = 0 Then
                            Call QuitarUserInvItem(UserIndex, Slot, 1)

                        End If
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                        Else

                            If TempTick - .Counters.RuidoPocion > 1000 Then
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                                .Counters.RuidoPocion = TempTick

                            End If

                        End If

                        Call WriteUpdateStrenght(UserIndex)
                        
                    Case 3 'Pocion roja, restaura HP
                            
                        ' # Está en un evento que cuenta las rojas
                        If .flags.RedValid Then
                            .flags.RedUsage = .flags.RedUsage + 1
                                
                            If .flags.RedUsage > .flags.RedLimit Then
                                Call WriteConsoleMsg(UserIndex, "Parece ser que el evento tiene limite de pociones rojas a configurado a un máximo de: " & .flags.RedLimit, FontTypeNames.FONTTYPE_INFORED)
                                Exit Sub

                            End If

                        End If
                            
                        If CanUse Then
                            .Stats.MinHp = .Stats.MinHp + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
    
                            If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp

                        End If
                        
                        'Quitamos del inv el item
                        If Obj.Ilimitado = 0 Then
                            Call QuitarUserInvItem(UserIndex, Slot, 1)

                        End If
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                        Else

                            If TempTick - .Counters.RuidoPocion > 1000 Then
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                                .Counters.RuidoPocion = TempTick

                            End If

                        End If
                        
                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageUpdateControlPotas(.Char.charindex, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMan, .Stats.MaxMan))
                        
                        Call WriteUpdateHP(UserIndex)
                            
                    Case 4 'Pocion azul, restaura MANA
                        
                        If CanUse Then
                            .Stats.MinMan = .Stats.MinMan + (Porcentaje(.Stats.MaxMan, 3) + .Stats.Elv \ 2 + 40 / .Stats.Elv)
                            
                            If .Stats.MinMan > .Stats.MaxMan Then .Stats.MinMan = .Stats.MaxMan

                        End If
                        
                        'Quitamos del inv el item
                        If Obj.Ilimitado = 0 Then
                            Call QuitarUserInvItem(UserIndex, Slot, 1)

                        End If
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                        Else

                            If TempTick - .Counters.RuidoPocion > 1000 Then
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                                .Counters.RuidoPocion = TempTick

                            End If

                        End If
                        
                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageUpdateControlPotas(.Char.charindex, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMan, .Stats.MaxMan))
                        
                        Call WriteUpdateMana(UserIndex)

                    Case 5 ' Pocion violeta

                        If .flags.Envenenado = 1 Then
                            .flags.Envenenado = 0
                            Call WriteUpdateEffect(UserIndex)
                            Call WriteConsoleMsg(UserIndex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)

                        End If

                        'Quitamos del inv el item
                        If Obj.Ilimitado = 0 Then
                            Call QuitarUserInvItem(UserIndex, Slot, 1)

                        End If
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                        Else

                            If TempTick - .Counters.RuidoPocion > 1000 Then
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                                .Counters.RuidoPocion = TempTick

                            End If

                        End If
                        
                        Call WriteUpdateUserStats(UserIndex)
                        
                    Case 6  ' Pocion Negra

                        If .flags.SlotEvent > 0 Or .flags.SlotReto > 0 Then Exit Sub
                        If .flags.Comerciando Then Exit Sub
                        
                        If .flags.Privilegios And PlayerType.User Then
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Call UserDie(UserIndex)
                            Call WriteConsoleMsg(UserIndex, "Sientes un gran mareo y pierdes el conocimiento.", FontTypeNames.FONTTYPE_FIGHT)

                        End If
                        
                        Call WriteUpdateUserStats(UserIndex)
                        
                    Case 7 ' Poción de energía

                        If .flags.Transform = 1 Then
                            Call WriteConsoleMsg(UserIndex, "No puedes utilizar esta poción estando transformado.", FontTypeNames.FONTTYPE_INFORED)

                            Exit Sub

                        End If
                        
                        If Obj.Ilimitado = 0 Then
                            Call QuitarUserInvItem(UserIndex, Slot, 1)

                        End If
                              
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                        Else

                            If TempTick - .Counters.RuidoPocion > 1000 Then
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                                .Counters.RuidoPocion = TempTick

                            End If
        
                        End If
                              
                        .Stats.MinSta = .Stats.MinSta + (.Stats.MaxSta * 0.1)

                        If .Stats.MinSta > .Stats.MaxSta Then .Stats.MinSta = .Stats.MaxSta
                              
                        Call WriteUpdateSta(UserIndex)

                End Select
               
                Call UpdateUserInv(False, UserIndex, Slot)
                
            Case eOBJType.oteffect
                Call mEffect.Effect_Add(UserIndex, Slot, .Invent.Object(Slot).ObjIndex)
            
            Case eOBJType.otUseOnce

                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
        
                'Usa el item
                .Stats.MinHam = .Stats.MinHam + Obj.MinHam

                If .Stats.MinHam > .Stats.MaxHam Then .Stats.MinHam = .Stats.MaxHam
                .flags.Hambre = 0
                Call WriteUpdateHungerAndThirst(UserIndex)
                'Sonido
                
                If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
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
                
                .Stats.Gld = .Stats.Gld + .Invent.Object(Slot).Amount
                .Invent.Object(Slot).Amount = 0
                .Invent.Object(Slot).ObjIndex = 0
                .Invent.NroItems = .Invent.NroItems - 1
                
                Call UpdateUserInv(False, UserIndex, Slot)
                Call WriteUpdateGold(UserIndex)
                
            Case eOBJType.otGuitaDsp

                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                
                .Stats.Eldhir = .Stats.Eldhir + .Invent.Object(Slot).Amount
                .Invent.Object(Slot).Amount = 0
                .Invent.Object(Slot).ObjIndex = 0
                .Invent.NroItems = .Invent.NroItems - 1
                
                Call UpdateUserInv(False, UserIndex, Slot)
                Call WriteUpdateDsp(UserIndex)

            Case eOBJType.otGemasEffect

                If .flags.Muerto Then
                    Call WriteConsoleMsg(UserIndex, "No puedes usar bonificaciones estando muerto.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                
                If .flags.SelectedBono > 0 Then
                    WriteConsoleMsg UserIndex, "Ya tienes un efecto activado.", FontTypeNames.FONTTYPE_INFO

                    Exit Sub

                End If
                
                .flags.SelectedBono = .Invent.Object(Slot).ObjIndex
                .Counters.TimeBono = ObjData(.Invent.Object(Slot).ObjIndex).BonoTime * 60
                
                WriteConsoleMsg UserIndex, "Has activado el efecto de la gema. El mismo desaparecerá en " & Int(.Counters.TimeBono / 60) & " minutos. Utiliza /EST para saber cuando tiempo te queda.", FontTypeNames.FONTTYPE_INFO

            Case eOBJType.otGemaTelep

                If Obj.TelepMap = 0 Or Obj.TelepX = 0 Or Obj.TelepY = 0 Then Exit Sub
                If .flags.Muerto Then Exit Sub
                If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Or .flags.Desafiando > 0 Or .Counters.Pena > 0 Then Exit Sub
                If MapInfo(.Pos.Map).Pk Then Exit Sub
                
                If .flags.Plata = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Necesitas ser usuario Plata para utilizar este scroll.", FontTypeNames.FONTTYPE_INFORED)

                    Exit Sub

                End If
                
                If .flags.ObjIndex > 0 Then
                    If .flags.ObjIndex <> .Invent.Object(Slot).ObjIndex Then
                        Call WriteConsoleMsg(UserIndex, "Ya tienes activado otro efecto. Haz clic sobre el objeto correspondiente", FontTypeNames.FONTTYPE_INFORED)

                        Exit Sub

                    End If
                    
                    Call WriteConsoleMsg(UserIndex, "Has regresado al mapa.", FontTypeNames.FONTTYPE_INFOGREEN)
                    WarpUserChar UserIndex, Obj.TelepMap, Obj.TelepX, Obj.TelepY, False
                Else
                    .flags.ObjIndex = .Invent.Object(Slot).ObjIndex
                    
                    If .flags.Premium > 0 Then
                        Obj.TelepTime = Obj.TelepTime + 10

                    End If
                    
                    .Counters.TimeTelep = Obj.TelepTime * 60
                    WarpUserChar UserIndex, Obj.TelepMap, Obj.TelepX, Obj.TelepY, False
                    WriteConsoleMsg UserIndex, "Has activado el efecto de la teletransportación.", FontTypeNames.FONTTYPE_INFO

                End If
                        
                'Quitamos del inv el item
                'Call QuitarUserInvItem(UserIndex, Slot, 1)
                'Call UpdateUserInv(False, UserIndex, Slot)

            Case eOBJType.otWeapon

                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                    
                If Not .Stats.MinSta > 5 And .Counters.Trabajando > 0 Then
                    .Counters.Trabajando = 0
                            
                    Call WriteUpdateUserTrabajo(UserIndex)

                End If
                        
                If Not .Stats.MinSta > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Estás muy cansad" & IIf(.Genero = eGenero.Hombre, "o", "a") & ".", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                
                If ObjData(ObjIndex).proyectil = 1 Then
                    If .Invent.Object(Slot).Equipped = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberías equipartela.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If

                    Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Proyectiles)  'Call WriteWorkRequestTarget(UserIndex, Proyectiles)

                Else
                    
                    Select Case ObjIndex
                    
                        Case CAÑA_PESCA, RED_PESCA, CAÑA_COFRES
                            
                            ' Lo tiene equipado?
                            If Not .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            Else
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Pesca)

                            End If
                            
                        Case HACHA_LEÑADOR
                            
                            ' Lo tiene equipado?
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Talar)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                        Case PIQUETE_MINERO
                        
                            ' Lo tiene equipado?
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Mineria)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                    End Select

                End If
            
            Case eOBJType.otLibroGuild
                    
                If .GuildIndex = 0 Then
                    Call WriteConsoleMsg(UserIndex, "¡¿A que clan quieres dar Experiencia si no posees ninguno?!", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                    
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Deshonrarás a tu clan si utilizas el Libro de Liderazgo", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                    
                If GuildsInfo(.GuildIndex).Lvl = MAX_GUILD_LEVEL Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Tu clan ya ha alcanzado el máximo nivel!!", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                    
                If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
                    Call WriteConsoleMsg(UserIndex, "¡Debes estar en Zona Segura para utilizar el Libro!", FontTypeNames.FONTTYPE_INFORED)
                    Exit Sub

                End If
    
                Call mGuilds.Guilds_AddExp(UserIndex, Obj.GuildExp)
                    
                ' Remove Object
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call UpdateUserInv(False, UserIndex, Slot)

            Case eOBJType.otTravel

                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! No puedes viajar en este estado.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                
                If .flags.TargetNPC = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Debes hacer clic sobre '" & GetVar(Npcs_FilePath, "NPC" & ObjData(ObjIndex).RequiredNpc, "NAME") & "'", FontTypeNames.FONTTYPE_INFORED)

                    Exit Sub

                End If
                
                If Npclist(.flags.TargetNPC).numero <> ObjData(ObjIndex).RequiredNpc Then
                    Call WriteConsoleMsg(UserIndex, "Debes hacer clic sobre '" & GetVar(Npcs_FilePath, "NPC" & ObjData(ObjIndex).RequiredNpc, "NAME") & "'", FontTypeNames.FONTTYPE_INFORED)

                    Exit Sub

                End If
                
                If Not MapInfo(Obj.TelepMap).CanTravel Then
                    Call WriteConsoleMsg(UserIndex, "No puedes viajar al destino ¡Será mejor que corras!", FontTypeNames.FONTTYPE_INFORED)

                    Exit Sub

                End If

                .Counters.Shield = 3
                Call FindLegalPos(UserIndex, Obj.TelepMap, Obj.TelepX, Obj.TelepY)
                Call WarpUserChar(UserIndex, Obj.TelepMap, Obj.TelepX, Obj.TelepY, True)
                
                Call WriteConsoleMsg(UserIndex, "Has llegado a tu destino.", FontTypeNames.FONTTYPE_INFOGREEN)
                
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                Call UpdateUserInv(False, UserIndex, Slot)
                
                Call RefreshCharStatus(UserIndex)

            Case eOBJType.otTransformVIP

                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar el skin estando vivo.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                
                If .flags.Mimetizado = 1 Or .flags.Transform Then
                    Call WriteConsoleMsg(UserIndex, "Ya tienes un efecto de transformación.", FontTypeNames.FONTTYPE_INFORED)

                    Exit Sub

                End If
                
                Call TransformVIP_User(UserIndex, Obj.Ropaje)
                
            Case eOBJType.otBebidas

                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If

                .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed

                If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
                .flags.Sed = 0
                Call WriteUpdateHungerAndThirst(UserIndex)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_BEBER, .Pos.X, .Pos.Y, .Char.charindex))

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
                 
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                                Call WriteConsoleMsg(UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)

                                Exit Sub

                            Else
                                Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)

                                Exit Sub

                            End If

                        Else

                            If TargObj.clave = Obj.clave Then
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                                Call WriteConsoleMsg(UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex

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

                If Not HayAgua(.Pos.Map, .flags.TargetX, .flags.TargetY) Then
                    Call WriteConsoleMsg(UserIndex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If

                MiObj.Amount = 1
                MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexAbierta
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

                If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
                .flags.Sed = 0
                Call WriteUpdateHungerAndThirst(UserIndex)
                MiObj.Amount = 1
                MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexCerrada
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
                
                If .Stats.MaxMan > 0 Then
                    If .flags.Hambre = 0 And .flags.Sed = 0 Then
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
                    
            Case eOBJType.otTeleportInvoker
                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, TeleportInvoker)
                         
            Case eOBJType.otInstrumentos

                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                
                If Obj.Real Then '¿Es el Cuerno Real?
                    If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                        If MapInfo(.Pos.Map).Pk = False Then
                            ' Call WriteConsoleMsg(UserIndex, "No hay peligro aquí. Es zona segura.", FontTypeNames.FONTTYPE_INFO)

                            '   Exit Sub

                        End If
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(Obj.Snd1, .Pos.X, .Pos.Y, .Char.charindex))
                        Else
                            Call AlertarFaccionarios(UserIndex)
                            Call SendData(SendTarget.ToFaction, UserIndex, PrepareMessagePlayEffect(Obj.Snd1, .Pos.X, .Pos.Y, .Char.charindex))

                        End If
                        
                        Exit Sub

                    Else
                        Call WriteConsoleMsg(UserIndex, "Sólo miembros del ejército real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If

                ElseIf Obj.Caos Then '¿Es el Cuerno Legión?

                    If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                        If MapInfo(.Pos.Map).Pk = False Then
                            Call WriteConsoleMsg(UserIndex, "No hay peligro aquí. Es zona segura.", FontTypeNames.FONTTYPE_INFO)

                            Exit Sub

                        End If
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(Obj.Snd1, .Pos.X, .Pos.Y, .Char.charindex))
                        Else
                            Call AlertarFaccionarios(UserIndex)
                            Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayEffect(Obj.Snd1, .Pos.X, .Pos.Y, .Char.charindex))

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
                    Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(Obj.Snd1, .Pos.X, .Pos.Y, .Char.charindex))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(Obj.Snd1, .Pos.X, .Pos.Y, .Char.charindex))

                End If
               
            Case eOBJType.otBarcos

                If .flags.Montando = 1 Then Exit Sub
                
                If ((LegalPos(.Pos.Map, .Pos.X - 1, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y - 1, True, False) Or LegalPos(.Pos.Map, .Pos.X + 1, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y + 1, True, False)) And .flags.Navegando = 0) Or .flags.Navegando = 1 Then
                    Call DoNavega(UserIndex, Obj, Slot)
                Else
                    Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)

                End If
                
            Case eOBJType.otMonturas

                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡No puedes montar tu mascota estando muerto!!", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                
                If ((LegalPos(.Pos.Map, .Pos.X, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y, True, False)) And .flags.Navegando = 0) Or .flags.Navegando = 1 Then
                        
                    Call WriteConsoleMsg(UserIndex, "¡No puedes montar en el agua!", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call DoEquita(UserIndex, Obj, Slot)

                End If

        End Select
    
    End With

    '<EhFooter>
    Exit Sub

UseInvItem_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.UseInvItem " & "at line " & Erl

    '</EhFooter>
End Sub

Sub TirarTodo(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        'If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        Call TirarTodosLosItems(UserIndex)

    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en TirarTodo. Error: " & Err.number & " - " & Err.description)

End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ItemSeCae_Err

    '</EhHeader>

    With ObjData(Index)
        ItemSeCae = (.Real <> 1 Or .NoSeCae = 0) And (.Caos <> 1 Or .NoSeCae = 0) And .OBJType <> eOBJType.otLlaves And .OBJType <> eOBJType.otBarcos And .NoSeCae = 0

    End With

    '<EhFooter>
    Exit Function

ItemSeCae_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.ItemSeCae " & "at line " & Erl
        
    '</EhFooter>
End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010 (ZaMa)
    '12/01/2010: ZaMa - Ahora los piratas no explotan items solo si estan entre 20 y 25
    '***************************************************
    On Error GoTo ErrHandler

    Dim i         As Byte

    Dim NuevaPos  As WorldPos

    Dim MiObj     As Obj

    Dim ItemIndex As Integer

    Dim DropAgua  As Boolean
    
    With UserList(UserIndex)

        For i = 1 To .CurrentInventorySlots
            ItemIndex = .Invent.Object(i).ObjIndex

            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo el Obj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.ObjIndex = ItemIndex

                    DropAgua = True

                    ' Es Ladron?
                    If .Clase = eClass.Thief Then

                        ' Si tiene galeon equipado
                        If .Invent.BarcoObjIndex = 187 Then

                            ' Limitación por nivel, después dropea normalmente
                            If .Stats.Elv >= 40 Then
                                ' No dropea en agua
                                DropAgua = False

                            End If

                        End If

                    End If
                    
                    Call Tilelibre(.Pos, NuevaPos, MiObj, DropAgua, True)
                    
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

                    End If

                End If

            End If

        Next i

    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error en TirarTodosLosItems. Error: " & Err.number & " - " & Err.description)

End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ItemNewbie_Err

    '</EhHeader>

    If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
    
    ItemNewbie = ObjData(ItemIndex).Newbie = 1
    '<EhFooter>
    Exit Function

ItemNewbie_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.ItemNewbie " & "at line " & Erl
        
    '</EhFooter>
End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo TirarTodosLosItemsNoNewbies_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: 23/11/2009
    '07/11/09: Pato - Fix bug #2819911
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************
    Dim i         As Byte

    Dim NuevaPos  As WorldPos

    Dim MiObj     As Obj

    Dim ItemIndex As Integer
    
    With UserList(UserIndex)

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        For i = 1 To UserList(UserIndex).CurrentInventorySlots
            ItemIndex = .Invent.Object(i).ObjIndex

            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo MiObj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.ObjIndex = ItemIndex
                    'Pablo (ToxicWaste) 24/01/2007
                    'Tira los Items no newbies en todos lados.
                    Tilelibre .Pos, NuevaPos, MiObj, True, True

                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

                    End If

                End If

            End If

        Next i

    End With

    '<EhFooter>
    Exit Sub

TirarTodosLosItemsNoNewbies_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.TirarTodosLosItemsNoNewbies " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub TirarTodosLosItemsEnMochila(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo TirarTodosLosItemsEnMochila_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/09 (Budi)
    '***************************************************
    Dim i         As Byte

    Dim NuevaPos  As WorldPos

    Dim MiObj     As Obj

    Dim ItemIndex As Integer
    
    With UserList(UserIndex)

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        For i = MAX_NORMAL_INVENTORY_SLOTS + 1 To .CurrentInventorySlots
            ItemIndex = .Invent.Object(i).ObjIndex

            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo MiObj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.ObjIndex = ItemIndex
                    Tilelibre .Pos, NuevaPos, MiObj, True, True

                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

                    End If

                End If

            End If

        Next i

    End With

    '<EhFooter>
    Exit Sub

TirarTodosLosItemsEnMochila_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.TirarTodosLosItemsEnMochila " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function getObjType(ByVal ObjIndex As Integer) As eOBJType

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo getObjType_Err

    '</EhHeader>

    If ObjIndex > 0 Then
        getObjType = ObjData(ObjIndex).OBJType

    End If
    
    '<EhFooter>
    Exit Function

getObjType_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.getObjType " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub moveItem(ByVal UserIndex As Integer, _
                    ByVal originalSlot As Integer, _
                    ByVal newSlot As Integer)

    '<EhHeader>
    On Error GoTo moveItem_Err

    '</EhHeader>

    Dim tmpObj      As UserOBJ

    Dim newObjIndex As Integer, originalObjIndex As Integer

    If (originalSlot <= 0) Or (newSlot <= 0) Then Exit Sub

    With UserList(UserIndex)

        If (originalSlot > .CurrentInventorySlots) Or (newSlot > .CurrentInventorySlots) Then Exit Sub
        If .flags.Comerciando Then Exit Sub
    
        tmpObj = .Invent.Object(originalSlot)
        .Invent.Object(originalSlot) = .Invent.Object(newSlot)
        .Invent.Object(newSlot) = tmpObj
    
        'Viva VB6 y sus putas deficiencias.
        If .Invent.AnilloEqpSlot = originalSlot Then
            .Invent.AnilloEqpSlot = newSlot
        ElseIf .Invent.AnilloEqpSlot = newSlot Then
            .Invent.AnilloEqpSlot = originalSlot

        End If
    
        If .Invent.AuraEqpSlot = originalSlot Then
            .Invent.AuraEqpSlot = newSlot
        ElseIf .Invent.AuraEqpSlot = newSlot Then
            .Invent.AuraEqpSlot = originalSlot

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
    
        If .Invent.MonturaSlot = originalSlot Then
            .Invent.MonturaSlot = newSlot
        ElseIf .Invent.MonturaSlot = newSlot Then
            .Invent.MonturaSlot = originalSlot

        End If
    
        If .Invent.ReliquiaSlot = originalSlot Then
            .Invent.ReliquiaSlot = newSlot
        ElseIf .Invent.ReliquiaSlot = newSlot Then
            .Invent.ReliquiaSlot = originalSlot

        End If
    
        If .Invent.MagicSlot = originalSlot Then
            .Invent.MagicSlot = newSlot
        ElseIf .Invent.MagicSlot = newSlot Then
            .Invent.MagicSlot = originalSlot

        End If
                
        If .Invent.PendientePartySlot = originalSlot Then
            .Invent.PendientePartySlot = newSlot
        ElseIf .Invent.PendientePartySlot = newSlot Then
            .Invent.PendientePartySlot = originalSlot

        End If

        Call UpdateUserInv(False, UserIndex, originalSlot)
        Call UpdateUserInv(False, UserIndex, newSlot)

    End With

    '<EhFooter>
    Exit Sub

moveItem_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.moveItem " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub MoveItem_Bank(ByVal UserIndex As Integer, _
                         ByVal originalSlot As Integer, _
                         ByVal newSlot As Integer, _
                         ByVal TypeBank As Byte)

    '<EhHeader>
    On Error GoTo MoveItem_Bank_Err

    '</EhHeader>

    Dim tmpObj      As UserOBJ

    Dim newObjIndex As Integer, originalObjIndex As Integer

    If (originalSlot <= 0) Or (newSlot <= 0) Then Exit Sub
    
    If TypeBank <> E_BANK.e_User And TypeBank <> E_BANK.e_Account Then Exit Sub
    
    With UserList(UserIndex)

        If (originalSlot > MAX_BANCOINVENTORY_SLOTS) Or (newSlot > MAX_BANCOINVENTORY_SLOTS) Then Exit Sub
        
        Select Case TypeBank

            Case E_BANK.e_User
                tmpObj = .BancoInvent.Object(originalSlot)
                .BancoInvent.Object(originalSlot) = .BancoInvent.Object(newSlot)
                .BancoInvent.Object(newSlot) = tmpObj
                
                Call UpdateBanUserInv(False, UserIndex, originalSlot)
                Call UpdateBanUserInv(False, UserIndex, newSlot)
                
            Case E_BANK.e_Account
                tmpObj = .Account.BancoInvent.Object(originalSlot)
                .Account.BancoInvent.Object(originalSlot) = .Account.BancoInvent.Object(newSlot)
                .Account.BancoInvent.Object(newSlot) = tmpObj
                
                Call UpdateBanUserInv_Account(False, UserIndex, originalSlot)
                Call UpdateBanUserInv_Account(False, UserIndex, newSlot)

        End Select
        
        Call UpdateVentanaBanco(UserIndex)
    
    End With
        
    '<EhFooter>
    Exit Sub

MoveItem_Bank_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.InvUsuario.MoveItem_Bank " & "at line " & Erl
        
    '</EhFooter>
End Sub
                   
