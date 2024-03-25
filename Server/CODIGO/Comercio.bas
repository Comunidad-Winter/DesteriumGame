Attribute VB_Name = "modSistemaComercio"
'*****************************************************
'Sistema de Comercio para Argentum Online
'Programado por Nacho (Integer)
'integer-x@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Enum eModoComercio

    Compra = 1
    Venta = 2

End Enum

Public Const REDUCTOR_PRECIOVENTA As Byte = 3

' Chequeamos que exista un mismo objeto para poder venderlo aquí.
Private Function Comercio_CheckItem(ByVal NpcIndex As Integer, _
                                    ByVal ObjIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo Comercio_CheckItem_Err

    '</EhHeader>

    Dim A As Long
    
    For A = 1 To MAX_INVENTORY_SLOTS

        With Npclist(NpcIndex).Invent

            If .Object(A).ObjIndex = ObjIndex Then
                Comercio_CheckItem = True

                Exit Function

            End If

        End With

    Next A

    '<EhFooter>
    Exit Function

Comercio_CheckItem_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSistemaComercio.Comercio_CheckItem " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Makes a trade. (Buy or Sell)
'
' @param Modo The trade type (sell or buy)
' @param UserIndex Specifies the index of the user
' @param NpcIndex specifies the index of the npc
' @param Slot Specifies which slot are you trying to sell / buy
' @param Cantidad Specifies how many items in that slot are you trying to sell / buy
Public Sub Comercio(ByVal Modo As eModoComercio, _
                    ByVal UserIndex As Integer, _
                    ByVal NpcIndex As Integer, _
                    ByVal Slot As Integer, _
                    ByVal cantidad As Integer, _
                    ByVal SelectedPrice As Byte)

    '<EhHeader>
    On Error GoTo Comercio_Err

    '</EhHeader>

    '*************************************************
    'Author: Nacho (Integer)
    'Last modified: 07/06/2010
    '27/07/08 (MarKoxX) | New changes in the way of trading (now when you buy it rounds to ceil and when you sell it rounds to floor)
    '  - 06/13/08 (NicoNZ)
    '07/06/2010: ZaMa - Los objetos se loguean si superan la cantidad de 1k (antes era solo si eran 1k).
    '*************************************************
    Dim PrecioDiamanteRojo As Long

    Dim PrecioDiamanteAzul As Long
            
    Dim PrecioPoints       As Long
        
    Dim Objeto             As Obj
    
    If cantidad < 1 Or Slot < 1 Then Exit Sub
    If SelectedPrice > 1 Then Exit Sub
          
    If Modo = eModoComercio.Compra Then
        If Slot > MAX_INVENTORY_SLOTS Then

            Exit Sub

        ElseIf cantidad > MAX_INVENTORY_OBJS Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha sido baneado por el sistema anti-cheats.", FontTypeNames.FONTTYPE_FIGHT))
            Call Ban(UserList(UserIndex).Name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados ítems:" & cantidad)
            UserList(UserIndex).flags.Ban = 1
            
            Call Protocol.Kick(UserIndex, "Has sido baneado por el Sistema AntiCheat.")
 
            Exit Sub

        ElseIf Not Npclist(NpcIndex).Invent.Object(Slot).Amount > 0 Then

            Exit Sub

        End If
            
        'Objeto.Amount = cantidad
        Objeto.ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex
              
        If ObjData(Objeto.ObjIndex).Upgrade.RequiredCant = 0 And ObjData(Objeto.ObjIndex).Points = 0 Then
            If SelectedPrice = 0 And ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Valor = 0 Then Exit Sub
            If SelectedPrice = 1 And ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).ValorEldhir = 0 Then Exit Sub

        End If
              
        If cantidad > Npclist(NpcIndex).Invent.Object(Slot).Amount Then cantidad = Npclist(UserList(UserIndex).flags.TargetNPC).Invent.Object(Slot).Amount
            
        'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
        'Es decir, 1.1 = 2, por lo cual se hace de la siguiente forma Precio = Clng(PrecioFinal + 0.5) Siempre va a darte el proximo numero. O el "Techo" (MarKoxX)
              
        If SelectedPrice = 0 Then
            PrecioDiamanteRojo = CLng((ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Valor / Descuento(UserIndex) * cantidad) + 0.5)
                
            If UserList(UserIndex).Stats.Gld < PrecioDiamanteRojo Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficientes Monedas de Oro.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If

        ElseIf SelectedPrice = 1 Then
            PrecioDiamanteAzul = CLng((ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).ValorEldhir / Descuento(UserIndex) * cantidad) + 0.5)

            If UserList(UserIndex).Stats.Eldhir < PrecioDiamanteAzul Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficientes DSP.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
        End If
            
        PrecioPoints = CLng((ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Points / Descuento(UserIndex) * cantidad) + 0.5)

        If UserList(UserIndex).Stats.Points < PrecioPoints Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes Puntos Desterium.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
            
        Dim SlotEvent As Byte

        SlotEvent = UserList(UserIndex).flags.SlotEvent
        
        If SlotEvent > 0 Then
            If Events(SlotEvent).LimitRed > 0 Then
                If Objeto.ObjIndex = POCION_ROJA Then
                    Call WriteConsoleMsg(UserIndex, "No puedes comprar pociones rojas en éste tipo de eventos.", FontTypeNames.FONTTYPE_INFORED)

                    Exit Sub

                End If

            End If
            
            If Events(SlotEvent).ChangeClass > 0 Or Events(SlotEvent).ChangeRaze > 0 Or Events(SlotEvent).ChangeLevel > 0 Then
                If Events(SlotEvent).TimeCancel > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Aún no está habilitada la compra de objetos. Espera a que se completen los cupos.", FontTypeNames.FONTTYPE_INFORED)

                    Exit Sub

                End If

            End If

        End If
                        
        Dim A As Long
            
        Objeto.Amount = cantidad
            
        ' @ Comprueba si tiene los objetos necesarios
        If ObjData(Objeto.ObjIndex).Upgrade.RequiredCant > 0 Then
            '  cantidad = 1        ' salvo las flechas es 1
                    
            If ObjData(Objeto.ObjIndex).OBJType = otFlechas Then
                Objeto.Amount = 500
                cantidad = 500
            Else
                Objeto.Amount = 1
                cantidad = 1

            End If
                    
            For A = 1 To ObjData(Objeto.ObjIndex).Upgrade.RequiredCant

                If Not TieneObjetos(ObjData(Objeto.ObjIndex).Upgrade.Required(A).ObjIndex, ObjData(Objeto.ObjIndex).Upgrade.Required(A).Amount, UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "No tienes " & ObjData(Objeto.ObjIndex).Upgrade.Required(A).ObjIndex & " (x " & ObjData(Objeto.ObjIndex).Upgrade.Required(A).Amount & ")", FontTypeNames.FONTTYPE_INFORED)
                    Exit Sub

                End If

            Next A
                
            For A = 1 To ObjData(Objeto.ObjIndex).Upgrade.RequiredCant
                Call QuitarObjetos(ObjData(Objeto.ObjIndex).Upgrade.Required(A).ObjIndex, ObjData(Objeto.ObjIndex).Upgrade.Required(A).Amount, UserIndex)
            Next A

        End If
        
        If MeterItemEnInventario(UserIndex, Objeto) = False Then Exit Sub
            
        ' # Compra un objeto por duración. Lo asignamos en un array()
        If ObjData(Objeto.ObjIndex).DurationDay > 0 Then

            Dim TempDate As String

            TempDate = DateAdd("d", ObjData(Objeto.ObjIndex).DurationDay, Now)
            Call Bonus_AddUser_Online(UserIndex, eBonusType.eObj, Objeto.ObjIndex, Objeto.Amount, 0, TempDate, False)

        End If
            
        If SelectedPrice = 0 Then
            UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld - PrecioDiamanteRojo
        ElseIf SelectedPrice = 1 Then
            UserList(UserIndex).Stats.Eldhir = UserList(UserIndex).Stats.Eldhir - PrecioDiamanteAzul

        End If
            
        UserList(UserIndex).Stats.Points = UserList(UserIndex).Stats.Points - PrecioPoints
            
        Dim CI As Integer
            
        CI = Npclist(NpcIndex).CommerceIndex

        If CI > 0 Then
            Comerciantes(CI).RewardDSP = Comerciantes(CI).RewardDSP + PrecioDiamanteAzul
            Comerciantes(CI).RewardGLD = Comerciantes(CI).RewardGLD + PrecioDiamanteRojo

        End If
        
        Call QuitarNpcInvItem(UserList(UserIndex).flags.TargetNPC, CByte(Slot), cantidad)
        
        If ObjData(Objeto.ObjIndex).Log = 1 Then
            Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eBuyObj, "compró del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
        ElseIf Objeto.Amount >= 100 Then 'Es mucha cantidad?

            'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
                Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eBuyObj, "compró del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)

            End If

        End If
        
        'Agregado para que no se vuelvan a vender las llaves si se recargan los .dat.
        If ObjData(Objeto.ObjIndex).OBJType = otLlaves Then
            Call WriteVar(Npcs_FilePath, "NPC" & Npclist(NpcIndex).numero, "obj" & Slot, Objeto.ObjIndex & "-0")
            Call logVentaCasa(UserList(UserIndex).Name & " compró " & ObjData(Objeto.ObjIndex).Name)

        End If
        
    ElseIf Modo = eModoComercio.Venta Then

        If cantidad > UserList(UserIndex).Invent.Object(Slot).Amount Then cantidad = UserList(UserIndex).Invent.Object(Slot).Amount
        
        Objeto.Amount = cantidad
        Objeto.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        
        If Objeto.ObjIndex = 0 Then

            Exit Sub
        ElseIf UserList(UserIndex).Invent.Object(Slot).Amount < 0 Or cantidad = 0 Then

            Exit Sub

        ElseIf Slot < LBound(UserList(UserIndex).Invent.Object()) Or Slot > UBound(UserList(UserIndex).Invent.Object()) Then

            Exit Sub
                
        ElseIf Npclist(NpcIndex).NPCtype = eCommerceChar And Npclist(NpcIndex).CommerceChar <> UCase$(UserList(UserIndex).Name) Then
            Call WriteConsoleMsg(UserIndex, "Sólo el dueño del personaje puede agregar objetos al mercado. ¿Deseas alquilarme luego? /ALQUILAR", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
                
        ElseIf ObjData(Objeto.ObjIndex).OBJType = otMonturas Then
            Call WriteConsoleMsg(UserIndex, "No puedes vender tu montura.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        ElseIf ObjData(Objeto.ObjIndex).NoNada = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes realizar ninguna acción con este objeto. ¡Podría ser de uso personal!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub
        ElseIf ObjData(Objeto.ObjIndex).Real = 1 Then

            If Npclist(NpcIndex).Name <> "SR" Then
                Call WriteConsoleMsg(UserIndex, "Las armaduras del ejército real sólo pueden ser vendidas a los sastres reales.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If

        ElseIf ObjData(Objeto.ObjIndex).Caos = 1 Then

            If Npclist(NpcIndex).Name <> "SC" Then
                Call WriteConsoleMsg(UserIndex, "Las armaduras de la legión oscura sólo pueden ser vendidas a los sastres del demonio.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If

        ElseIf (Npclist(NpcIndex).TipoItems <> ObjData(Objeto.ObjIndex).OBJType) Or Objeto.ObjIndex = iORO Then

            If Npclist(NpcIndex).TipoItems = 999 Then
                If ObjData(Objeto.ObjIndex).ValorEldhir <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "¡Ja ja ja! Vende tus baratijas en aquel mercado.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If

            End If

            If Npclist(NpcIndex).TipoItems <> 1000 Then
                
                ' Criaturas que venden items de todos los tipos.
                If Npclist(NpcIndex).TipoItems = 998 Then
                    If Not Comercio_CheckItem(NpcIndex, Objeto.ObjIndex) Then
                        Call WriteConsoleMsg(UserIndex, "Lo siento, debes vender tus objetos en el mercado global, no aquí.", FontTypeNames.FONTTYPE_INFO)
        
                        Exit Sub

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "Lo siento, debes vender tus objetos en el mercado global, no aquí.", FontTypeNames.FONTTYPE_INFO)
    
                    Exit Sub

                End If

            End If

        ElseIf UserList(UserIndex).flags.Privilegios And PlayerType.SemiDios Then
            Call WriteConsoleMsg(UserIndex, "No puedes vender ítems.", FontTypeNames.FONTTYPE_WARNING)

            Exit Sub

        End If
        
        If ObjData(Objeto.ObjIndex).OBJType = otGemaTelep Then
            Call WriteConsoleMsg(UserIndex, "No puedes vender este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
                
        If ObjData(Objeto.ObjIndex).OBJType = otTransformVIP Then
            Call WriteConsoleMsg(UserIndex, "No puedes vender este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
            
        If SelectedPrice = 1 Then
            Call WriteConsoleMsg(UserIndex, "Lo siento Joven, deberás vender tu objeto por Monedas de Oro o bien a los usuarios del Servidor.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If

        If Npclist(NpcIndex).NPCtype <> eNPCType.eCommerceChar Then
            ' Comprueba si no tenía que vender el Objeto para pasar a la próxima misión.
            Call Quests_AddSale(UserIndex, Objeto.ObjIndex, Objeto.Amount)
            Call QuitarUserInvItem(UserIndex, Slot, cantidad)
        
            PrecioDiamanteRojo = Fix(SalePrice(Objeto.ObjIndex) * cantidad)
            PrecioDiamanteAzul = Fix(SalePriceDiamanteAzul(Objeto.ObjIndex) * cantidad)
        
            If SelectedPrice = 0 Then
                UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld + PrecioDiamanteRojo
            ElseIf SelectedPrice = 1 Then
                UserList(UserIndex).Stats.Eldhir = UserList(UserIndex).Stats.Eldhir + PrecioDiamanteAzul

            End If

        Else

            If ObjData(Objeto.ObjIndex).ValorEldhir = 0 And ObjData(Objeto.ObjIndex).Valor = 0 Then
                Call WriteConsoleMsg(UserIndex, "Por el momento estos objetos no puedes venderlos aquí. Pero espera pronta noticias", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub

            End If

            Call QuitarUserInvItem(UserIndex, Slot, cantidad)

        End If
        
        If UserList(UserIndex).Stats.Gld > MAXORO Then UserList(UserIndex).Stats.Gld = MAXORO
        
        If UserList(UserIndex).Stats.Eldhir > MAXORO Then UserList(UserIndex).Stats.Eldhir = MAXORO
            
        If Not (ObjData(Objeto.ObjIndex).LvlMax > 0 And ObjData(Objeto.ObjIndex).LvlMax < UserList(UserIndex).Stats.Elv) Then

            Dim NpcSlot As Integer

            NpcSlot = SlotEnNPCInv(NpcIndex, Objeto.ObjIndex, Objeto.Amount)
            
            If NpcSlot <= MAX_INVENTORY_SLOTS Then 'Slot valido
                'Mete el obj en el slot
                Npclist(NpcIndex).Invent.Object(NpcSlot).ObjIndex = Objeto.ObjIndex
                Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = Npclist(NpcIndex).Invent.Object(NpcSlot).Amount + Objeto.Amount
    
                If Npclist(NpcIndex).Invent.Object(NpcSlot).Amount > MAX_INVENTORY_OBJS Then
                    Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = MAX_INVENTORY_OBJS

                End If
                
                Call EnviarNpcInv(NpcSlot, UserIndex, UserList(UserIndex).flags.TargetNPC)

            End If

        End If
            
        If ObjData(Objeto.ObjIndex).Log = 1 Then
            Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eSaleObj, "vendió del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name & " al NPC: " & Npclist(NpcIndex).numero)
        ElseIf Objeto.Amount >= 100 Then 'Es mucha cantidad?

            'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
                Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eSaleObj, "vendió del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name & " al NPC: " & Npclist(NpcIndex).numero)

            End If

        End If
        
    End If
    
    Call UpdateUserInv(False, UserIndex, Slot)
    Call WriteUpdateGold(UserIndex)
    Call WriteUpdateDsp(UserIndex)
    Call WriteTradeOK(UserIndex)
    
    Call SubirSkill(UserIndex, eSkill.Comerciar, True)
    
    '<EhFooter>
    Exit Sub

Comercio_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSistemaComercio.Comercio " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Public Sub IniciarComercioNPC(ByVal UserIndex As Integer)

    '*************************************************
    'Author: Nacho (Integer)
    'Last modified: 2/8/06
    '*************************************************
    '<EhHeader>
    On Error GoTo IniciarComercioNPC_Err

    '</EhHeader>
    Call EnviarNpcInv(0, UserIndex, UserList(UserIndex).flags.TargetNPC)
    UserList(UserIndex).flags.Comerciando = True
    Call WriteCommerceInit(UserIndex, Npclist(UserList(UserIndex).flags.TargetNPC).Name, Npclist(UserList(UserIndex).flags.TargetNPC).Quest, Npclist(UserList(UserIndex).flags.TargetNPC).Quests)
    '<EhFooter>
    Exit Sub

IniciarComercioNPC_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSistemaComercio.IniciarComercioNPC " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Function SlotEnNPCInv(ByVal NpcIndex As Integer, _
                              ByVal Objeto As Integer, _
                              ByVal cantidad As Integer) As Integer

    '*************************************************
    'Author: Nacho (Integer)
    'Last modified: 2/8/06
    '*************************************************
    '<EhHeader>
    On Error GoTo SlotEnNPCInv_Err

    '</EhHeader>
    SlotEnNPCInv = 1

    Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).ObjIndex = Objeto And Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).Amount + cantidad <= MAX_INVENTORY_OBJS
        
        SlotEnNPCInv = SlotEnNPCInv + 1

        If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
        
    Loop
    
    If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then
    
        SlotEnNPCInv = 1
        
        Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).ObjIndex = 0
        
            SlotEnNPCInv = SlotEnNPCInv + 1

            If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
            
        Loop
        
        If SlotEnNPCInv <= MAX_INVENTORY_SLOTS Then Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
    
    End If
    
    '<EhFooter>
    Exit Function

SlotEnNPCInv_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSistemaComercio.SlotEnNPCInv " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Send the inventory of the Npc to the user
'
' @param userIndex The index of the User
' @param npcIndex The index of the NPC

Private Sub EnviarNpcInv(ByVal UpdateSlot As Byte, _
                         ByVal UserIndex As Integer, _
                         ByVal NpcIndex As Integer)

    '<EhHeader>
    On Error GoTo EnviarNpcInv_Err

    '</EhHeader>

    '*************************************************
    'Author: Nacho (Integer)
    'Last Modified: 07/08/2022
    ' Actualiza solo los Slots necesarios
    '*************************************************
    Dim Slot     As Byte

    Dim val      As Single
    
    Dim val2     As Single
    
    Dim thisObj  As Obj

    Dim DummyObj As Obj
    
    If NpcIndex = 0 Then Exit Sub
    
    If UpdateSlot > 0 Then
        If Npclist(NpcIndex).Invent.Object(UpdateSlot).ObjIndex > 0 Then
                
            thisObj.ObjIndex = Npclist(NpcIndex).Invent.Object(UpdateSlot).ObjIndex
            thisObj.Amount = Npclist(NpcIndex).Invent.Object(UpdateSlot).Amount
                
            val = (ObjData(thisObj.ObjIndex).Valor) / Descuento(UserIndex)
            val2 = (ObjData(thisObj.ObjIndex).ValorEldhir) / Descuento(UserIndex)
                
            Call WriteChangeNPCInventorySlot(UserIndex, UpdateSlot, thisObj, val, val2)
        Else
    
            Call WriteChangeNPCInventorySlot(UserIndex, UpdateSlot, DummyObj, 0, 0)

        End If

    Else

        For Slot = 1 To MAX_NORMAL_INVENTORY_SLOTS
    
            If Npclist(NpcIndex).Invent.Object(Slot).ObjIndex > 0 Then

                thisObj.ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex
                thisObj.Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount
                
                val = (ObjData(thisObj.ObjIndex).Valor) / Descuento(UserIndex)
                val2 = (ObjData(thisObj.ObjIndex).ValorEldhir) / Descuento(UserIndex)
                
                Call WriteChangeNPCInventorySlot(UserIndex, Slot, thisObj, val, val2)
            Else
    
                Call WriteChangeNPCInventorySlot(UserIndex, Slot, DummyObj, 0, 0)

            End If
        
        Next Slot

    End If
    
    '<EhFooter>
    Exit Sub

EnviarNpcInv_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSistemaComercio.EnviarNpcInv " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Devuelve el valor de venta del objeto
'
' @param ObjIndex  El número de objeto al cual le calculamos el precio de venta

Public Function SalePrice(ByVal ObjIndex As Integer) As Single

    '<EhHeader>
    On Error GoTo SalePrice_Err

    '</EhHeader>

    '*************************************************
    'Author: Nicolás (NicoNZ)
    '
    '*************************************************
    If ObjIndex < 1 Or ObjIndex > UBound(ObjData) Then Exit Function
    If ItemNewbie(ObjIndex) Then Exit Function
    
    SalePrice = ObjData(ObjIndex).Valor / REDUCTOR_PRECIOVENTA
    '<EhFooter>
    Exit Function

SalePrice_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSistemaComercio.SalePrice " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Devuelve el valor de venta del objeto
'
' @param ObjIndex  El número de objeto al cual le calculamos el precio de venta

Public Function SalePriceDiamanteAzul(ByVal ObjIndex As Integer) As Single

    '<EhHeader>
    On Error GoTo SalePriceDiamanteAzul_Err

    '</EhHeader>

    '*************************************************
    'Author: Nicolás (NicoNZ)
    '
    '*************************************************
    If ObjIndex < 1 Or ObjIndex > UBound(ObjData) Then Exit Function
    If ItemNewbie(ObjIndex) Then Exit Function
    
    SalePriceDiamanteAzul = ObjData(ObjIndex).ValorEldhir / REDUCTOR_PRECIOVENTA
    '<EhFooter>
    Exit Function

SalePriceDiamanteAzul_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSistemaComercio.SalePriceDiamanteAzul " & "at line " & Erl
        
    '</EhFooter>
End Function

Private Function Descuento(ByVal UserIndex As Integer) As Single

    '*************************************************
    'Author: Nacho (Integer)
    'Last modified: 2/8/06
    '*************************************************
    '<EhHeader>
    On Error GoTo Descuento_Err

    '</EhHeader>
    Descuento = 1 + UserList(UserIndex).Stats.UserSkills(eSkill.Comerciar) / 100
    '<EhFooter>
    Exit Function

Descuento_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modSistemaComercio.Descuento " & "at line " & Erl
        
    '</EhFooter>
End Function
