Attribute VB_Name = "modBanco"

Option Explicit

Sub IniciarDeposito(ByVal Userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

      'Hacemos un Update del inventario del usuario
20    Call UpdateBanUserInv(True, Userindex, 0)
      'Actualizamos el dinero
30    Call WriteUpdateUserStats(Userindex)
      'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
40    Call WriteBankInit(Userindex)
50    UserList(Userindex).flags.Comerciando = True

Errhandler:

End Sub

Sub SendBanObj(Userindex As Integer, Slot As Byte, Object As UserOBJ)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    UserList(Userindex).BancoInvent.Object(Slot) = Object

20    Call WriteChangeBankSlot(Userindex, Slot)

End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, ByVal Userindex As Integer, ByVal Slot As Byte)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim NullObj As UserOBJ
      Dim LoopC As Byte

10    With UserList(Userindex)
          'Actualiza un solo slot
20        If Not UpdateAll Then
              'Actualiza el inventario
30            If .BancoInvent.Object(Slot).objindex > 0 Then
40                Call SendBanObj(Userindex, Slot, .BancoInvent.Object(Slot))
50            Else
60                Call SendBanObj(Userindex, Slot, NullObj)
70            End If
80        Else
          'Actualiza todos los slots
90            For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
                  'Actualiza el inventario
100               If .BancoInvent.Object(LoopC).objindex > 0 Then
110                   Call SendBanObj(Userindex, LoopC, .BancoInvent.Object(LoopC))
120               Else
130                   Call SendBanObj(Userindex, LoopC, NullObj)
140               End If
150           Next LoopC
160       End If
170   End With

End Sub

Sub UserRetiraItem(ByVal Userindex As Integer, ByVal i As Integer, ByVal Cantidad As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler


20    If Cantidad < 1 Then Exit Sub

30    Call WriteUpdateUserStats(Userindex)

40           If UserList(Userindex).BancoInvent.Object(i).Amount > 0 Then
50                If Cantidad > UserList(Userindex).BancoInvent.Object(i).Amount Then Cantidad = UserList(Userindex).BancoInvent.Object(i).Amount
                  'Agregamos el obj que compro al inventario
60                Call UserReciveObj(Userindex, CInt(i), Cantidad)
                  'Actualizamos el inventario del usuario
70                Call UpdateUserInv(True, Userindex, 0)
                  'Actualizamos el banco
80                Call UpdateBanUserInv(True, Userindex, 0)
90           End If
             
              'Actualizamos la ventana de comercio
100           Call UpdateVentanaBanco(Userindex)

Errhandler:

End Sub

Sub UserReciveObj(ByVal Userindex As Integer, ByVal objindex As Integer, ByVal Cantidad As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim Slot As Integer
      Dim obji As Integer

10    With UserList(Userindex)
20        If .BancoInvent.Object(objindex).Amount <= 0 Then Exit Sub
          
30        obji = .BancoInvent.Object(objindex).objindex
          
          
          '¿Ya tiene un objeto de este tipo?
40        Slot = 1
50        Do Until .Invent.Object(Slot).objindex = obji And _
             .Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
              
60            Slot = Slot + 1
70            If Slot > .CurrentInventorySlots Then
80                Exit Do
90            End If
100       Loop
          
          'Sino se fija por un slot vacio
110       If Slot > .CurrentInventorySlots Then
120           Slot = 1
130           Do Until .Invent.Object(Slot).objindex = 0
140               Slot = Slot + 1

150               If Slot > .CurrentInventorySlots Then
160                   Call WriteConsoleMsg(Userindex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
170                   Exit Sub
180               End If
190           Loop
200           .Invent.NroItems = .Invent.NroItems + 1
210       End If
          
          
          
          'Mete el obj en el slot
220       If .Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
              'Menor que MAX_INV_OBJS
230           .Invent.Object(Slot).objindex = obji
240           .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + Cantidad
              
250           Call QuitarBancoInvItem(Userindex, CByte(objindex), Cantidad)
260       Else
270           Call WriteConsoleMsg(Userindex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
280       End If
290   End With

End Sub

Sub QuitarBancoInvItem(ByVal Userindex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim objindex As Integer

10    With UserList(Userindex)
20        objindex = .BancoInvent.Object(Slot).objindex

          'Quita un Obj

30        .BancoInvent.Object(Slot).Amount = .BancoInvent.Object(Slot).Amount - Cantidad
          
40        If .BancoInvent.Object(Slot).Amount <= 0 Then
50            .BancoInvent.NroItems = .BancoInvent.NroItems - 1
60            .BancoInvent.Object(Slot).objindex = 0
70            .BancoInvent.Object(Slot).Amount = 0
80        End If
90    End With
          
End Sub

Sub UpdateVentanaBanco(ByVal Userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        Call WriteBankOK(Userindex)
End Sub

Sub UserDepositaItem(ByVal Userindex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler
20        If UserList(Userindex).Invent.Object(Item).Amount > 0 And Cantidad > 0 Then
30            If Cantidad > UserList(Userindex).Invent.Object(Item).Amount Then Cantidad = UserList(Userindex).Invent.Object(Item).Amount
              
              'Agregamos el obj que deposita al banco
40            Call UserDejaObj(Userindex, CInt(Item), Cantidad)
              
              'Actualizamos el inventario del usuario
50            Call UpdateUserInv(True, Userindex, 0)
              
              'Actualizamos el inventario del banco
60            Call UpdateBanUserInv(True, Userindex, 0)
70        End If
          
          'Actualizamos la ventana del banco
80        Call UpdateVentanaBanco(Userindex)
Errhandler:
End Sub

Sub UserDejaObj(ByVal Userindex As Integer, ByVal objindex As Integer, ByVal Cantidad As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim Slot As Integer
          Dim obji As Integer
          
10        If Cantidad < 1 Then Exit Sub
          
20        With UserList(Userindex)
30            obji = .Invent.Object(objindex).objindex
              
              '¿Ya tiene un objeto de este tipo?
40            Slot = 1
50            Do Until .BancoInvent.Object(Slot).objindex = obji And _
                  .BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
60                Slot = Slot + 1
                  
70                If Slot > MAX_BANCOINVENTORY_SLOTS Then
80                    Exit Do
90                End If
100           Loop
              
              'Sino se fija por un slot vacio antes del slot devuelto
110           If Slot > MAX_BANCOINVENTORY_SLOTS Then
120               Slot = 1
130               Do Until .BancoInvent.Object(Slot).objindex = 0
140                   Slot = Slot + 1
                      
150                   If Slot > MAX_BANCOINVENTORY_SLOTS Then
160                       Call WriteConsoleMsg(Userindex, "No tienes mas espacio en el banco!!", FontTypeNames.FONTTYPE_INFO)
170                       Exit Sub
180                   End If
190               Loop
                  
200               .BancoInvent.NroItems = .BancoInvent.NroItems + 1
210           End If
              
220           If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido
                  'Mete el obj en el slot
230               If .BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
                      
                      'Menor que MAX_INV_OBJS
240                   .BancoInvent.Object(Slot).objindex = obji
250                   .BancoInvent.Object(Slot).Amount = .BancoInvent.Object(Slot).Amount + Cantidad
                      
260                   Call QuitarUserInvItem(Userindex, CByte(objindex), Cantidad)
270               Else
280                   Call WriteConsoleMsg(Userindex, "El banco no puede cargar tantos objetos.", FontTypeNames.FONTTYPE_INFO)
290               End If
300           End If
310       End With
End Sub

Sub SendUserBovedaTxt(ByVal SendIndex As Integer, ByVal Userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error Resume Next
      Dim j As Integer

20    Call WriteConsoleMsg(SendIndex, UserList(Userindex).Name, FontTypeNames.FONTTYPE_INFO)
30    Call WriteConsoleMsg(SendIndex, "Tiene " & UserList(Userindex).BancoInvent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)

40    For j = 1 To MAX_BANCOINVENTORY_SLOTS
50        If UserList(Userindex).BancoInvent.Object(j).objindex > 0 Then
60            Call WriteConsoleMsg(SendIndex, "Objeto " & j & " " & ObjData(UserList(Userindex).BancoInvent.Object(j).objindex).Name & " Cantidad:" & UserList(Userindex).BancoInvent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)
70        End If
80    Next

End Sub

Sub SendUserBovedaTxtFromChar(ByVal SendIndex As Integer, ByVal charName As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error Resume Next
      Dim j As Integer
      Dim CharFile As String, tmp As String
      Dim ObjInd As Long, ObjCant As Long

20    CharFile = CharPath & charName & ".chr"

30    If FileExist(CharFile, vbNormal) Then
40        Call WriteConsoleMsg(SendIndex, charName, FontTypeNames.FONTTYPE_INFO)
50        Call WriteConsoleMsg(SendIndex, "Tiene " & GetVar(CharFile, "BancoInventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
60        For j = 1 To MAX_BANCOINVENTORY_SLOTS
70            tmp = GetVar(CharFile, "BancoInventory", "Obj" & j)
80            ObjInd = ReadField(1, tmp, Asc("-"))
90            ObjCant = ReadField(2, tmp, Asc("-"))
100           If ObjInd > 0 Then
110               Call WriteConsoleMsg(SendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)
120           End If
130       Next
140   Else
150       Call WriteConsoleMsg(SendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
160   End If

End Sub

