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

Public Sub Comercio(ByVal Modo As eModoComercio, ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Slot As Integer, ByVal Cantidad As Integer)
      '*************************************************
      'Author: Nacho (Integer)
      'Last modified: 27/07/08 (MarKoxX) | New changes in the way of trading (now when you buy it rounds to ceil and when you sell it rounds to floor)
      '  - 06/13/08 (NicoNZ)
      '*************************************************
          Dim Precio As Long
          Dim ValorCopas As Long
          Dim ValorEldhir As Long
          Dim Objeto As Obj
         
10        If Cantidad < 1 Or Slot < 1 Then Exit Sub
         
20        If Modo = eModoComercio.Compra Then
30            If Slot > MAX_INVENTORY_SLOTS Then
40                Exit Sub
50            ElseIf Cantidad > MAX_INVENTORY_OBJS Then
60                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha sido baneado por el sistema anti-cheats.", FontTypeNames.FONTTYPE_FIGHT))
70                Call Ban(UserList(UserIndex).Name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados ítems:" & Cantidad)
80                UserList(UserIndex).flags.Ban = 1
90                Call WriteErrorMsg(UserIndex, "Has sido baneado por el Sistema AntiCheat.")
100               Call FlushBuffer(UserIndex)
110               Call CloseSocket(UserIndex)
120               Exit Sub
130           ElseIf Not Npclist(NpcIndex).Invent.Object(Slot).Amount > 0 Then
140               Exit Sub
150           End If
             
160           If Cantidad > Npclist(NpcIndex).Invent.Object(Slot).Amount Then Cantidad = Npclist(UserList(UserIndex).flags.TargetNPC).Invent.Object(Slot).Amount
             
170           Objeto.Amount = Cantidad
180           Objeto.ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex
             
              'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
              'Es decir, 1.1 = 2, por lo cual se hace de la siguiente forma Precio = Clng(PrecioFinal + 0.5) Siempre va a darte el proximo numero. O el "Techo" (MarKoxX)
             
190           Precio = CLng((ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).valor / Descuento(UserIndex) * Cantidad) + 0.5)
200           ValorCopas = ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).copaS * Cantidad
210          ValorEldhir = ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Eldhir * Cantidad
              
220           If UserList(UserIndex).Stats.Gld < Precio Then
230               Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero.", FontTypeNames.FONTTYPE_INFO)
240               Exit Sub
250           End If
             
260           If TieneObjetos(880, ValorCopas, UserIndex) = False Then
270            Call WriteConsoleMsg(UserIndex, "No tienes suficientes Ds Points para negociar conmigo.", FontTypeNames.FONTTYPE_INFO)
280            Exit Sub
290           End If
              
300                   If TieneObjetos(943, ValorEldhir, UserIndex) = False Then
310            Call WriteConsoleMsg(UserIndex, "No tienes suficiente Eldhires para negociar conmigo.", FontTypeNames.FONTTYPE_INFO)
320            Exit Sub
330           End If
             
340           If MeterItemEnInventario(UserIndex, Objeto) = False Then
                'Call WriteConsoleMsg(UserIndex, "No puedes cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
350                Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
360                Call WriteTradeOK(UserIndex)
370               Exit Sub
380           End If
             
390           UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld - Precio
400           If ValorCopas > 0 Then Call QuitarObjetos(880, ValorCopas, UserIndex): Call UpdateUserInv(True, UserIndex, 0)
410           If ValorEldhir > 0 Then Call QuitarObjetos(943, ValorEldhir, UserIndex): Call UpdateUserInv(True, UserIndex, 0)
420           Call QuitarNpcInvItem(UserList(UserIndex).flags.TargetNPC, CByte(Slot), Cantidad)
             
              'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
              'Es un Objeto que tenemos que loguear?
430           If ObjData(Objeto.ObjIndex).LOG = 1 Then
440               Call LogDesarrollo(UserList(UserIndex).Name & " compró del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
450           ElseIf Objeto.Amount = 1000 Then 'Es mucha cantidad?
                  'Si no es de los prohibidos de loguear, lo logueamos.
460               If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
470                   Call LogDesarrollo(UserList(UserIndex).Name & " compró del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
480               End If
490           End If
             
              'Agregado para que no se vuelvan a vender las llaves si se recargan los .dat.
500           If ObjData(Objeto.ObjIndex).OBJType = otLlaves Then
510               Call WriteVar(DatPath & "NPCs.dat", "NPC" & Npclist(NpcIndex).Numero, "obj" & Slot, Objeto.ObjIndex & "-0")
520               Call logVentaCasa(UserList(UserIndex).Name & " compró " & ObjData(Objeto.ObjIndex).Name)
530           End If
             
540       ElseIf Modo = eModoComercio.Venta Then
            If Slot > MAX_INVENTORY_SLOTS Then
                Exit Sub
            ElseIf Cantidad > MAX_INVENTORY_OBJS Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha sido baneado por el sistema anti-cheats.", FontTypeNames.FONTTYPE_FIGHT))
                Call Ban(UserList(UserIndex).Name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados ítems:" & Cantidad)
                UserList(UserIndex).flags.Ban = 1
                Call WriteErrorMsg(UserIndex, "Has sido baneado por el Sistema AntiCheat.")
               Call FlushBuffer(UserIndex)
              Call CloseSocket(UserIndex)
               Exit Sub
           ElseIf Not UserList(UserIndex).Invent.Object(Slot).Amount > 0 Then
              Exit Sub
          End If

550           If Cantidad > UserList(UserIndex).Invent.Object(Slot).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Slot).Amount
             
560           Objeto.Amount = Cantidad
570           Objeto.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
             
580           If Objeto.ObjIndex = 0 Then
590               Exit Sub
600           ElseIf (Npclist(NpcIndex).TipoItems <> ObjData(Objeto.ObjIndex).OBJType And Npclist(NpcIndex).TipoItems <> eOBJType.otCualquiera) Or Objeto.ObjIndex = iORO Then
610               Call WriteConsoleMsg(UserIndex, "Lo siento, no estoy interesado en este tipo de objetos.", FontTypeNames.FONTTYPE_INFO)
620               Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
630               Call WriteTradeOK(UserIndex)
640               Exit Sub
650           ElseIf ObjData(Objeto.ObjIndex).Real = 1 Then
660               If Npclist(NpcIndex).Name <> "SR" Then
670                   Call WriteConsoleMsg(UserIndex, "Las armaduras del ejército real sólo pueden ser vendidas a los sastres reales.", FontTypeNames.FONTTYPE_INFO)
680                   Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
690                   Call WriteTradeOK(UserIndex)
700                   Exit Sub
710               End If
720           ElseIf ObjData(Objeto.ObjIndex).Caos = 1 Then
730               If Npclist(NpcIndex).Name <> "SC" Then
740                   Call WriteConsoleMsg(UserIndex, "Las armaduras de la legión oscura sólo pueden ser vendidas a los sastres del demonio.", FontTypeNames.FONTTYPE_INFO)
750                   Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
760                   Call WriteTradeOK(UserIndex)
770                   Exit Sub
780               End If
790           ElseIf UserList(UserIndex).Invent.Object(Slot).Amount < 0 Or Cantidad = 0 Then
800               Exit Sub
810           ElseIf Slot < LBound(UserList(UserIndex).Invent.Object()) Or Slot > UBound(UserList(UserIndex).Invent.Object()) Then
820               Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
830               Exit Sub
840           ElseIf UserList(UserIndex).flags.Privilegios And PlayerType.Consejero Then
850               Call WriteConsoleMsg(UserIndex, "No puedes vender ítems.", FontTypeNames.FONTTYPE_WARNING)
860               Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
870               Call WriteTradeOK(UserIndex)
880               Exit Sub
890           End If
             
900           Call QuitarUserInvItem(UserIndex, Slot, Cantidad)
             
              'Precio = Round(ObjData(Objeto.ObjIndex).valor / REDUCTOR_PRECIOVENTA * Cantidad, 0)
910           Precio = Fix(SalePrice(Objeto.ObjIndex) * Cantidad)
920           UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld + Precio
             
930           If UserList(UserIndex).Stats.Gld > MaxOro Then _
                  UserList(UserIndex).Stats.Gld = MaxOro
             
              Dim NpcSlot As Integer
940           NpcSlot = SlotEnNPCInv(NpcIndex, Objeto.ObjIndex, Objeto.Amount)
             
950           If NpcSlot <= MAX_INVENTORY_SLOTS Then 'Slot valido
                  'Mete el obj en el slot
960               Npclist(NpcIndex).Invent.Object(NpcSlot).ObjIndex = Objeto.ObjIndex
970               Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = Npclist(NpcIndex).Invent.Object(NpcSlot).Amount + Objeto.Amount
980               If Npclist(NpcIndex).Invent.Object(NpcSlot).Amount > MAX_INVENTORY_OBJS Then
990                   Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = MAX_INVENTORY_OBJS
1000              End If
1010          End If
             
              'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
              'Es un Objeto que tenemos que loguear?
1020          If ObjData(Objeto.ObjIndex).LOG = 1 Then
1030              Call LogDesarrollo(UserList(UserIndex).Name & " vendió al NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
1040          ElseIf Objeto.Amount = 1000 Then 'Es mucha cantidad?
                  'Si no es de los prohibidos de loguear, lo logueamos.
1050              If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
1060                  Call LogDesarrollo(UserList(UserIndex).Name & " vendió al NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
1070              End If
1080          End If
             
1090      End If
           
1100      Call UpdateUserInv(True, UserIndex, 0)
1110      Call WriteUpdateUserStats(UserIndex)
1120      Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
1130      Call WriteTradeOK(UserIndex)
             
1140      Call SubirSkill(UserIndex, eSkill.Comerciar, True)
End Sub
    Public Sub IniciarComercioNPC(ByVal UserIndex As Integer)
      '*************************************************
      'Author: Nacho (Integer)
      'Last modified: 2/8/06
      '*************************************************
10        Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
20        UserList(UserIndex).flags.Comerciando = True
30        Call WriteCommerceInit(UserIndex)
End Sub

Private Function SlotEnNPCInv(ByVal NpcIndex As Integer, ByVal Objeto As Integer, ByVal Cantidad As Integer) As Integer
      '*************************************************
      'Author: Nacho (Integer)
      'Last modified: 2/8/06
      '*************************************************
10        SlotEnNPCInv = 1
20        Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).ObjIndex = Objeto _
            And Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).Amount + Cantidad <= MAX_INVENTORY_OBJS
              
30            SlotEnNPCInv = SlotEnNPCInv + 1
40            If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
              
50        Loop
          
60        If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then
          
70            SlotEnNPCInv = 1
              
80            Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).ObjIndex = 0
              
90                SlotEnNPCInv = SlotEnNPCInv + 1
100               If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
                  
110           Loop
              
120           If SlotEnNPCInv <= MAX_INVENTORY_SLOTS Then Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
          
130       End If
          
End Function

Private Function Descuento(ByVal UserIndex As Integer) As Single
      '*************************************************
      'Author: Nacho (Integer)
      'Last modified: 2/8/06
      '*************************************************
10        Descuento = 1 + UserList(UserIndex).Stats.UserSkills(eSkill.Comerciar) / 100
End Function

''
' Send the inventory of the Npc to the user
'
' @param userIndex The index of the User
' @param npcIndex The index of the NPC

Private Sub EnviarNpcInv(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
      '*************************************************
      'Author: Nacho (Integer)
      'Last Modified: 06/14/08
      'Last Modified By: Nicolás Ezequiel Bouhid (NicoNZ)
      '*************************************************
          Dim Slot As Byte
          Dim val As Single
          
          If NpcIndex = 0 Then
            'LogError "El personaje " & UserList(UserIndex).Name & " uso el NpcIndex EnviarNpcInv 0"
10          Exit Sub
          End If
          
For Slot = 1 To MAX_NORMAL_INVENTORY_SLOTS
20            If Npclist(NpcIndex).Invent.Object(Slot).ObjIndex > 0 Then
                  Dim thisObj As Obj
                  
30                thisObj.ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex
40                thisObj.Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount
                  
50                val = (ObjData(thisObj.ObjIndex).valor) / Descuento(UserIndex)
                  
60                Call WriteChangeNPCInventorySlot(UserIndex, Slot, thisObj, val)
70            Else
                  Dim DummyObj As Obj
80                Call WriteChangeNPCInventorySlot(UserIndex, Slot, DummyObj, 0)
90            End If
100       Next Slot
End Sub

''
' Devuelve el valor de venta del objeto
'
' @param ObjIndex  El número de objeto al cual le calculamos el precio de venta

Public Function SalePrice(ByVal ObjIndex As Integer) As Single
      '*************************************************
      'Author: Nicolás (NicoNZ)
      '
      '*************************************************
10        If ObjIndex < 1 Or ObjIndex > UBound(ObjData) Then Exit Function
20        If ItemNewbie(ObjIndex) Then Exit Function
          
30        SalePrice = ObjData(ObjIndex).valor / REDUCTOR_PRECIOVENTA
End Function
