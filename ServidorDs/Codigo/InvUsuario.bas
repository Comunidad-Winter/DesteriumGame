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
       
10    On Error GoTo Errhandler
       
          Dim i As Integer
          Dim ObjIndex As Integer
         
20        For i = 1 To UserList(UserIndex).CurrentInventorySlots
30            ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
40            If ObjIndex > 0 Then
50                If (ObjData(ObjIndex).ObjType <> eOBJType.otLlaves And _
                  ObjData(ObjIndex).ObjType <> eOBJType.otMonturas And _
                  ObjData(ObjIndex).ObjType <> eOBJType.otMonturasDraco And _
                      ObjData(ObjIndex).ObjType <> eOBJType.otBarcos And _
                      Not ItemNewbie(ObjIndex)) Then
60                      TieneObjetosRobables = True
70                      Exit Function
80                End If
90            End If
100       Next i
         
110       Exit Function
       
Errhandler:
120       Call LogError("Error en TieneObjetosRobables. Error: " & Err.Number & " - " & Err.Description)
End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional ByRef sMotivo As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: 14/01/2010 (ZaMa)
      '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
      '***************************************************

10    On Error GoTo manejador

          Dim flag As Boolean
          
          
          'Admins can use ANYTHING!
20        If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
30            If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then
                  Dim i As Integer
40                For i = 1 To NUMCLASES
50                    If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).clase Then
60                        ClasePuedeUsarItem = False
70                        sMotivo = "Tu clase no puede usar este objeto."
80                        Exit Function
90                    End If
100               Next i
110           End If
120       End If
          
130       ClasePuedeUsarItem = True

140   Exit Function

manejador:
150       LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim j As Integer

10    With UserList(UserIndex)
20        For j = 1 To UserList(UserIndex).CurrentInventorySlots
30            If .Invent.Object(j).ObjIndex > 0 Then
                   
40                 If ObjData(.Invent.Object(j).ObjIndex).Newbie = 1 Then _
                          Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
50                        Call UpdateUserInv(False, UserIndex, j)
              
60            End If
70        Next j
          
          '[Barrin 17-12-03] Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
          'es transportado a su hogar de origen ;)
80        If UCase$(MapInfo(.Pos.map).Restringir) = "NEWBIE" Then
              
              Dim DeDonde As WorldPos
              
90            Select Case .Hogar
                  Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
100                   DeDonde = Lindos
110               Case eCiudad.cUllathorpe
120                   DeDonde = Ullathorpe
130               Case eCiudad.cBanderbill
140                   DeDonde = Banderbill
150               Case Else
160                   DeDonde = Nix
170           End Select
              
180           Call WarpUserChar(UserIndex, DeDonde.map, DeDonde.X, DeDonde.Y, True)
          
190       End If
          '[/Barrin]
200   End With

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim j As Integer

10    With UserList(UserIndex)
20        For j = 1 To .CurrentInventorySlots
30            .Invent.Object(j).ObjIndex = 0
40            .Invent.Object(j).Amount = 0
50            .Invent.Object(j).Equipped = 0
60        Next j
          
70        .Invent.NroItems = 0
          
80        .Invent.ArmourEqpObjIndex = 0
90        .Invent.ArmourEqpSlot = 0
          
100       .Invent.WeaponEqpObjIndex = 0
110       .Invent.WeaponEqpSlot = 0
          
120       .Invent.CascoEqpObjIndex = 0
130       .Invent.CascoEqpSlot = 0
          
140       .Invent.EscudoEqpObjIndex = 0
150       .Invent.EscudoEqpSlot = 0
          
160       .Invent.AnilloEqpObjIndex = 0
170       .Invent.AnilloEqpSlot = 0
          
180       .Invent.MunicionEqpObjIndex = 0
190       .Invent.MunicionEqpSlot = 0
          
200       .Invent.BarcoObjIndex = 0
210       .Invent.BarcoSlot = 0
          
220       .Invent.MonturaObjIndex = 0
230       .Invent.MonturaSlot = 0
          
240       .Invent.MochilaEqpObjIndex = 0
250       .Invent.MochilaEqpSlot = 0
          
260       .Invent.AnilloNpcSlot = 0
270       .Invent.AnilloNpcObjIndex = 0
280   End With

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 23/01/2007
      '23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
      '***************************************************
10    On Error GoTo Errhandler


20    If Cantidad > 100000 Then Exit Sub
30    If Cantidad < 0 Then Exit Sub

40    With UserList(UserIndex)
50         If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then
60            Exit Sub
70        End If

          'SI EL Pjta TIENE ORO LO TIRAMOS
80        If (Cantidad > 0) And (Cantidad <= .Stats.Gld) Then
                  Dim i As Byte
                  Dim MiObj As Obj
                  'info debug
                  Dim loops As Integer
                  
                  'Seguridad Alkon (guardo el oro tirado si supera los 50k)
90                If Cantidad > 50000 Then
                      Dim j As Integer
                      Dim k As Integer
                      Dim m As Integer
                      Dim Cercanos As String
100                   m = .Pos.map
110                   For j = .Pos.X - 10 To .Pos.X + 10
120                       For k = .Pos.Y - 10 To .Pos.Y + 10
130                           If InMapBounds(m, j, k) Then
140                               If MapData(m, j, k).UserIndex > 0 Then
150                                   Cercanos = Cercanos & UserList(MapData(m, j, k).UserIndex).Name & ","
160                               End If
170                           End If
180                       Next k
190                   Next j
200                   Call LogDesarrollo(.Name & " tira oro. Cercanos: " & Cercanos)
210               End If
                  '/Seguridad
                  Dim Extra As Long
                  Dim TeniaOro As Long
220               TeniaOro = .Stats.Gld
230               If Cantidad > 500000 Then 'Para evitar explotar demasiado
240                   Extra = Cantidad - 500000
250                   Cantidad = 500000
260               End If
                  
270               Do While (Cantidad > 0)
                      
280                   If Cantidad > MAX_INVENTORY_OBJS And .Stats.Gld > MAX_INVENTORY_OBJS Then
290                       MiObj.Amount = MAX_INVENTORY_OBJS
300                       Cantidad = Cantidad - MiObj.Amount
310                   Else
320                       MiObj.Amount = Cantidad
330                       Cantidad = Cantidad - MiObj.Amount
340                   End If
          
350                   MiObj.ObjIndex = iORO
                      
360                   If EsGm(UserIndex) Then Call LogGM(.Name, "Tiró cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
                      Dim AuxPos As WorldPos
                      
370                   If .clase = eClass.Pirat And .Invent.BarcoObjIndex = 476 Then
380                       AuxPos = TirarItemAlPiso(.Pos, MiObj, False)
390                       If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
400                           .Stats.Gld = .Stats.Gld - MiObj.Amount
410                       End If
420                   Else
430                       AuxPos = TirarItemAlPiso(.Pos, MiObj, True)
440                       If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
450                           .Stats.Gld = .Stats.Gld - MiObj.Amount
460                       End If
470                   End If
                      
                      'info debug
480                   loops = loops + 1
490                   If loops > 100 Then
500                       LogError ("Error en tiraroro")
510                       Exit Sub
520                   End If
                      
530               Loop
540               If TeniaOro = .Stats.Gld Then Extra = 0
550               If Extra > 0 Then
560                   .Stats.Gld = .Stats.Gld - Extra
570               End If
              
580       End If
590   End With

600   Exit Sub

Errhandler:
610       Call LogError("Error en TirarOro. Error " & Err.Number & " : " & Err.Description)
End Sub
Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

20        If Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots Then Exit Sub
          
30        With UserList(UserIndex).Invent.Object(Slot)
40            If .Amount <= Cantidad And .Equipped = 1 Then
50                Call Desequipar(UserIndex, Slot)
60            End If
              
              'Quita un objeto
70            .Amount = .Amount - Cantidad
              '¿Quedan mas?
80            If .Amount <= 0 Then
90                UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
100               .ObjIndex = 0
110               .Amount = 0
120           End If
130       End With

140   Exit Sub

Errhandler:
150       Call LogError("Error en QuitarUserInvItem. Error " & Err.Number & " : " & Err.Description)
          
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

      Dim NullObj As UserOBJ
      Dim LoopC As Long

20    With UserList(UserIndex)
          'Actualiza un solo slot
30        If Not UpdateAll Then
          
              'Actualiza el inventario
40            If .Invent.Object(Slot).ObjIndex > 0 Then
50                Call ChangeUserInv(UserIndex, Slot, .Invent.Object(Slot))
60            Else
70                Call ChangeUserInv(UserIndex, Slot, NullObj)
80            End If
          
90        Else
          
          'Actualiza todos los slots
100           For LoopC = 1 To .CurrentInventorySlots
                  'Actualiza el inventario
110               If .Invent.Object(LoopC).ObjIndex > 0 Then
120                   Call ChangeUserInv(UserIndex, LoopC, .Invent.Object(LoopC))
130               Else
140                   Call ChangeUserInv(UserIndex, LoopC, NullObj)
150               End If
160           Next LoopC
170       End If
          
180       Exit Sub
190   End With

Errhandler:
200       Call LogError("Error en UpdateUserInv. Error " & Err.Number & " : " & Err.Description)

End Sub
Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 11/5/2010
      '11/5/2010 - ZaMa: Arreglo bug que permitia apilar mas de 10k de items.
      '***************************************************

      Dim DropObj As Obj
      Dim MapObj As Obj

10    With UserList(UserIndex)

20        If Num > 0 Then
              
30            DropObj.ObjIndex = .Invent.Object(Slot).ObjIndex
              
40            If (ItemNewbie(DropObj.ObjIndex) And (.flags.Privilegios And PlayerType.User)) And .flags.Muerto <> 1 Then
50                Call WriteConsoleMsg(UserIndex, "No puedes tirar objetos newbie.", FontTypeNames.FONTTYPE_INFO)
60                Exit Sub
70            End If
              
80            If (ItemFaccionario(DropObj.ObjIndex) And (.flags.Privilegios And PlayerType.User)) Then
90                Call WriteConsoleMsg(UserIndex, "¡¡No puedes tirar tu armadura faccionaria!!", FontTypeNames.FONTTYPE_INFO)
100               Exit Sub
110           End If
              
120           If (ItemVIP(DropObj.ObjIndex) And (.flags.Privilegios And PlayerType.User)) Then
130               Call WriteConsoleMsg(UserIndex, "Por seguridad no puedes arrojar tus objetos Oro, Plata o Bronce.", FontTypeNames.FONTTYPE_INFO)
140               Exit Sub
150           End If
              
160           If (ItemVIPB(DropObj.ObjIndex) And (.flags.Privilegios And PlayerType.User)) Then
170               Call WriteConsoleMsg(UserIndex, "Por seguridad no puedes arrojar tus objetos Oro, Plata o Bronce.", FontTypeNames.FONTTYPE_INFO)
180               Exit Sub
190           End If
              
200           If (ItemVIPP(DropObj.ObjIndex) And (.flags.Privilegios And PlayerType.User)) Then
210               Call WriteConsoleMsg(UserIndex, "Por seguridad no puedes arrojar tus objetos Oro, Plata o Bronce.", FontTypeNames.FONTTYPE_INFO)
220               Exit Sub
230           End If
              
240           If Not EsGm(UserIndex) Then
250               If ObjData(DropObj.ObjIndex).NpcTipo <> 0 Then
260                   WriteConsoleMsg UserIndex, "¡No puedes tirar la transformación del anillo!", FontTypeNames.FONTTYPE_INFO
270                   Exit Sub
280               End If
290           End If
              
300           DropObj.Amount = MinimoInt(Num, .Invent.Object(Slot).Amount)

              'Check objeto en el suelo
310           MapObj.ObjIndex = MapData(.Pos.map, X, Y).ObjInfo.ObjIndex
320           MapObj.Amount = MapData(.Pos.map, X, Y).ObjInfo.Amount
              
330           If MapObj.ObjIndex = 0 Or MapObj.ObjIndex = DropObj.ObjIndex Then
              
340               If MapObj.Amount = MAX_INVENTORY_OBJS Then
350                   Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
360                   Exit Sub
370               End If
                  
380                           If ObjData(DropObj.ObjIndex).Caos = 1 Or ObjData(DropObj.ObjIndex).Real = 1 Then
390               WriteConsoleMsg UserIndex, "¡¡ATENCIÓN!! ¡ACABAS DE TIRAR TU ARMADURA FACCIONARIA!", FontTypeNames.FONTTYPE_GUILD
400               End If
                  
410                If ObjData(DropObj.ObjIndex).Premium = 1 And (.flags.Privilegios = PlayerType.User) Then
420           WriteConsoleMsg UserIndex, "No puedes tirar items PREMIUM!", FontTypeNames.FONTTYPE_INFO
430           Exit Sub
440           End If
                  
450               If DropObj.Amount + MapObj.Amount > MAX_INVENTORY_OBJS Then
460                   DropObj.Amount = MAX_INVENTORY_OBJS - MapObj.Amount
470               End If
                  
         
480               Call MakeObj(DropObj, map, X, Y)

490               Call QuitarUserInvItem(UserIndex, Slot, DropObj.Amount)
500               Call UpdateUserInv(False, UserIndex, Slot)
                  
510               If ObjData(DropObj.ObjIndex).ObjType = eOBJType.otBarcos Then
520                   Call WriteConsoleMsg(UserIndex, "¡¡ATENCIÓN!! ¡ACABAS DE TIRAR TU BARCA!", FontTypeNames.FONTTYPE_GUILD)
530               End If
                  
540               If ObjData(DropObj.ObjIndex).ObjType = eOBJType.otMonturas Then
550               WriteConsoleMsg UserIndex, "¡¡ATENCIÓN!! ¡ACABAS DE TIRAR TU MONTURA!", FontTypeNames.FONTTYPE_GUILD
560               End If
                  
570                            If ObjData(DropObj.ObjIndex).ObjType = eOBJType.otMonturasDraco Then
580                   Call WriteConsoleMsg(UserIndex, "¡¡ATENCIÓN!! ¡ACABAS DE TIRAR TU MONTURA!", FontTypeNames.FONTTYPE_TALK)
590               End If
                  
600               If ObjData(DropObj.ObjIndex).ObjType = eOBJType.otLunar Then
610   Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha tirado una Gema Lunar. Se encuentra en el mapa " & UserList(UserIndex).Pos.map & ", " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_GUILD))
620   End If

630               If ObjData(DropObj.ObjIndex).ObjType = eOBJType.otvioleta Then
640   Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha tirado una Gema Violeta. Se encuentra en el mapa " & UserList(UserIndex).Pos.map & ", " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_GUILD))
650   End If
           
660                    If ObjData(DropObj.ObjIndex).ObjType = eOBJType.otAzul Then
670   Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha tirado una Gema Azul. Se encuentra en el mapa " & UserList(UserIndex).Pos.map & ", " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_GUILD))
680   End If

690                  If ObjData(DropObj.ObjIndex).ObjType = eOBJType.otroja Then
700   Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha tirado una Gema Roja. Se encuentra en el mapa " & UserList(UserIndex).Pos.map & ", " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_GUILD))
710   End If

720                  If ObjData(DropObj.ObjIndex).ObjType = eOBJType.otverde Then
730   Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha tirado una Gema Verde. Se encuentra en el mapa " & UserList(UserIndex).Pos.map & ", " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_GUILD))
740   End If

750                  If ObjData(DropObj.ObjIndex).ObjType = eOBJType.otLila Then
760   Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha tirado una Gema Lila. Se encuentra en el mapa " & UserList(UserIndex).Pos.map & ", " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_GUILD))
770   End If
                  
780                  If ObjData(DropObj.ObjIndex).ObjType = eOBJType.otNaranja Then
790   Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha tirado una Gema Naranja. Se encuentra en el mapa " & UserList(UserIndex).Pos.map & ", " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_GUILD))
800   End If

810                  If ObjData(DropObj.ObjIndex).ObjType = eOBJType.otCeleste Then
820   Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha tirado una Gema Celeste. Se encuentra en el mapa " & UserList(UserIndex).Pos.map & ", " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_GUILD))
830   End If

                  
840               If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.Name, "Tiró cantidad:" & Num & " Objeto:" & ObjData(DropObj.ObjIndex).Name)
                  
                  'Log de Objetos que se tiran al piso. Pablo (ToxicWaste) 07/09/07
                  'Es un Objeto que tenemos que loguear?
850               If ObjData(DropObj.ObjIndex).LOG = 1 Then
860                   Call LogDesarrollo(.Name & " tiró al piso " & DropObj.Amount & " " & ObjData(DropObj.ObjIndex).Name & " Mapa: " & map & " X: " & X & " Y: " & Y)
870               ElseIf DropObj.Amount > 5000 Then 'Es mucha cantidad? > Subí a 5000 el minimo porque si no se llenaba el log de cosas al pedo. (NicoNZ)
                      'Si no es de los prohibidos de loguear, lo logueamos.
880                   If ObjData(DropObj.ObjIndex).NoLog <> 1 Then
890                       Call LogDesarrollo(.Name & " tiró al piso " & DropObj.Amount & " " & ObjData(DropObj.ObjIndex).Name & " Mapa: " & map & " X: " & X & " Y: " & Y)
900                   End If
910               End If
920           Else
930               Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
940           End If
950       End If
                             
960   End With

End Sub

Sub EraseObj(ByVal Num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

10    MapData(map, X, Y).ObjInfo.Amount = MapData(map, X, Y).ObjInfo.Amount - Num

20    If MapData(map, X, Y).ObjInfo.Amount <= 0 Then
30        MapData(map, X, Y).ObjInfo.ObjIndex = 0
40        MapData(map, X, Y).ObjInfo.Amount = 0
          
50        Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectDelete(X, Y))
60    End If

End Sub
Sub MakeObj(ByRef Obj As Obj, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

10    If Obj.ObjIndex > 0 And Obj.ObjIndex <= UBound(ObjData) Then

20        If MapData(map, X, Y).ObjInfo.ObjIndex = Obj.ObjIndex Then
30            MapData(map, X, Y).ObjInfo.Amount = MapData(map, X, Y).ObjInfo.Amount + Obj.Amount
40        Else
50            MapData(map, X, Y).ObjInfo = Obj
              
60            Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(Obj.ObjIndex).GrhIndex, X, Y))
70        End If
80    End If

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As Obj) As Boolean
10    On Error GoTo Errhandler

      'Call LogTarea("MeterItemEnInventario")
       
      Dim X As Integer
      Dim Y As Integer
      Dim Slot As Byte

      '¿el user ya tiene un objeto del mismo tipo?
20    Slot = 1
30    Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And _
               UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
40       Slot = Slot + 1
50       If Slot > MAX_INVENTORY_SLOTS Then
60             Exit Do
70       End If
80    Loop
          
      'Sino busca un slot vacio
90    If Slot > MAX_INVENTORY_SLOTS Then
100      Slot = 1
110      Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
120          Slot = Slot + 1
130          If Slot > MAX_INVENTORY_SLOTS Then
140              Call WriteConsoleMsg(UserIndex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
150              MeterItemEnInventario = False
160              Exit Function
170          End If
180      Loop
190      UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
200   End If
          
      'Mete el objeto
210   If UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
         'Menor que MAX_INV_OBJS
220      UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
230      UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount
240   Else
250      UserList(UserIndex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
260   End If
          
270   MeterItemEnInventario = True
             
280   Call UpdateUserInv(False, UserIndex, Slot)


290   Exit Function
Errhandler:

End Function

Sub GetObj(ByVal UserIndex As Integer)
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 18/12/2009
      '30/08/2011: Shak - Ahora el oro va al inventario como los objetos.
      '***************************************************
       
On Error GoTo Errhandler

          Dim Obj As ObjData
          Dim MiObj As Obj
          Dim ObjPos As String
         
10        With UserList(UserIndex)
              '¿Hay algun obj?
20            If MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex > 0 Then
                  '¿Esta permitido agarrar este obj?
30                If ObjData(MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then
                      Dim X As Integer
                      Dim Y As Integer
                      Dim Slot As Byte
                     
40                    X = .Pos.X
50                    Y = .Pos.Y
                     
60                    Obj = ObjData(MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex)
70                    MiObj.Amount = MapData(.Pos.map, X, Y).ObjInfo.Amount
80                    MiObj.ObjIndex = MapData(.Pos.map, X, Y).ObjInfo.ObjIndex
                      
90    If ObjData(MiObj.ObjIndex).ObjType = eOBJType.otAzul Then
100                   Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha agarrado una Gema Azul. Se encuentra en el mapa " & UserList(UserIndex).Pos.map, FontTypeNames.FONTTYPE_GUILD))
110                     End If
                       
120   If ObjData(MiObj.ObjIndex).ObjType = eOBJType.otNaranja Then
130                   Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha agarrado una Gema Naranja. Se encuentra en el mapa " & UserList(UserIndex).Pos.map, FontTypeNames.FONTTYPE_GUILD))
140                     End If
150   If ObjData(MiObj.ObjIndex).ObjType = eOBJType.otCeleste Then
160                   Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha agarrado una Gema Celeste. Se encuentra en el mapa " & UserList(UserIndex).Pos.map, FontTypeNames.FONTTYPE_GUILD))
170                     End If
                       
180   If ObjData(MiObj.ObjIndex).ObjType = eOBJType.otLila Then
190                   Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha agarrado una Gema Lila. Se encuentra en el mapa " & UserList(UserIndex).Pos.map, FontTypeNames.FONTTYPE_GUILD))
200                     End If
                       
210   If ObjData(MiObj.ObjIndex).ObjType = eOBJType.otroja Then
220                   Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha agarrado una Gema Roja. Se encuentra en el mapa " & UserList(UserIndex).Pos.map, FontTypeNames.FONTTYPE_GUILD))
230                     End If
                       
240   If ObjData(MiObj.ObjIndex).ObjType = eOBJType.otverde Then
250                   Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha agarrado una Gema Verde. Se encuentra en el mapa " & UserList(UserIndex).Pos.map, FontTypeNames.FONTTYPE_GUILD))
260                     End If
                       
270   If ObjData(MiObj.ObjIndex).ObjType = eOBJType.otvioleta Then
280                   Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha agarrado una Gema Violeta. Se encuentra en el mapa " & UserList(UserIndex).Pos.map, FontTypeNames.FONTTYPE_GUILD))
290                     End If
                        
300   If ObjData(MiObj.ObjIndex).ObjType = eOBJType.otLunar Then
310                   Call SendData(ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha agarrado una Gema Lunar. Se encuentra en el mapa " & UserList(UserIndex).Pos.map, FontTypeNames.FONTTYPE_GUILD))
320                     End If
              
                      ' El oro se va al inventario
330                   If MeterItemEnInventario(UserIndex, MiObj) Then
                              'Quitamos el objeto
340                           Call EraseObj(MapData(.Pos.map, X, Y).ObjInfo.Amount, .Pos.map, .Pos.X, .Pos.Y)
350                           If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
             
                              'Log de Objetos que se agarran del piso. Pablo (ToxicWaste) 07/09/07
                              'Es un Objeto que tenemos que loguear?
360                           If ObjData(MiObj.ObjIndex).LOG = 1 Then
370                               ObjPos = " Mapa: " & .Pos.map & " X: " & .Pos.X & " Y: " & .Pos.Y
380                               Call LogDesarrollo(.Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)
390                           ElseIf MiObj.Amount > MAX_INVENTORY_OBJS - 1000 Then 'Es mucha cantidad?
                                  'Si no es de los prohibidos de loguear, lo logueamos.
400                               If ObjData(MiObj.ObjIndex).NoLog <> 1 Then
410                                   ObjPos = " Mapa: " & .Pos.map & " X: " & .Pos.X & " Y: " & .Pos.Y
420                                   Call LogDesarrollo(.Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)
430                               End If
440                           End If
450                       End If
                     
460               End If
470           Else
480               Call WriteConsoleMsg(UserIndex, "No hay nada aquí.", FontTypeNames.FONTTYPE_INFO)
490          End If
500       End With
       
       Exit Sub
       
Errhandler:
    LogError "Error en Sub GetObj(ByVal UserIndex As Integer) at line " & Erl
End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          'Desequipa el item slot del inventario
          Dim Obj As ObjData
          
20        With UserList(UserIndex)
30            With .Invent
40                If (Slot < LBound(.Object)) Or (Slot > UBound(.Object)) Then
50                    Exit Sub
60                ElseIf .Object(Slot).ObjIndex = 0 Then
70                    Exit Sub
80                End If
                  
90                Obj = ObjData(.Object(Slot).ObjIndex)
100           End With
              
110           Select Case Obj.ObjType
                  Case eOBJType.otWeapon
120                   With .Invent
130                       .Object(Slot).Equipped = 0
140                       .WeaponEqpObjIndex = 0
150                       .WeaponEqpSlot = 0
160                   End With
                      
170                   If Not .flags.Mimetizado = 1 Then
180                       With .Char
190                           .WeaponAnim = NingunArma
200                           Call ChangeUserChar(UserIndex, .body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
210                       End With
220                   End If
                  
230               Case eOBJType.otFlechas
240                   With .Invent
250                       .Object(Slot).Equipped = 0
260                       .MunicionEqpObjIndex = 0
270                       .MunicionEqpSlot = 0
280                   End With
                  
                  
290               Case eOBJType.otManchas
300                   With .Invent
310                       .Object(Slot).Equipped = 0
320                       .MunicionEqpObjIndex = 0
330                       .MunicionEqpSlot = 0
340                   End With
       
                  
350               Case eOBJType.otAnillo
360                   With .Invent
370                       .Object(Slot).Equipped = 0
380                       .AnilloEqpObjIndex = 0
390                       .AnilloEqpSlot = 0
400                   End With
                  
410               Case eOBJType.otarmadura
420                   With .Invent
430                       .Object(Slot).Equipped = 0
440                       .ArmourEqpObjIndex = 0
450                       .ArmourEqpSlot = 0
460                   End With
                      
470                   Call DarCuerpoDesnudo(UserIndex, .flags.Mimetizado = 1)

480                   With .Char
490                       Call ChangeUserChar(UserIndex, .body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
500                   End With
                       
510               Case eOBJType.otcasco
520                   With .Invent
530                       .Object(Slot).Equipped = 0
540                       .CascoEqpObjIndex = 0
550                       .CascoEqpSlot = 0
560                   End With
                      
570                   If Not .flags.Mimetizado = 1 Then
580                       With .Char
590                           .CascoAnim = NingunCasco
600                           Call ChangeUserChar(UserIndex, .body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
610                       End With
620                   End If
                  
630               Case eOBJType.otescudo
640                   With .Invent
650                       .Object(Slot).Equipped = 0
660                       .EscudoEqpObjIndex = 0
670                       .EscudoEqpSlot = 0
680                   End With
                      
690                   If Not .flags.Mimetizado = 1 Then
700                       With .Char
710                           .ShieldAnim = NingunEscudo
720                           Call ChangeUserChar(UserIndex, .body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
730                       End With
740                   End If
                  
750               Case eOBJType.otMochilas
760                   With .Invent
770                       .Object(Slot).Equipped = 0
780                       .MochilaEqpObjIndex = 0
790                       .MochilaEqpSlot = 0
800                   End With
                      
810                   Call InvUsuario.TirarTodosLosItemsEnMochila(UserIndex)
820                   .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
                      
830               Case eOBJType.otAnilloNpc
840                   General.TransformUserNpc UserIndex, ObjData(.Invent.AnilloNpcObjIndex).NpcTipo, False
                      
850                   With .Invent
860                       .Object(Slot).Equipped = 0
870                       .AnilloNpcObjIndex = 0
880                       .AnilloNpcSlot = 0
890                   End With
900           End Select
910       End With
          
920       Call WriteUpdateUserStats(UserIndex)
930       Call UpdateUserInv(False, UserIndex, Slot)
          
940       Exit Sub

Errhandler:
950       Call LogError("Error en Desquipar. Error " & Err.Number & " : " & Err.Description & " at line " & Erl)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional ByRef sMotivo As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: 14/01/2010 (ZaMa)
      '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
      '***************************************************

10    On Error GoTo Errhandler
          
20        If ObjData(ObjIndex).Mujer = 1 Then
30            SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Hombre
40        ElseIf ObjData(ObjIndex).Hombre = 1 Then
50            SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Mujer
60        Else
70            SexoPuedeUsarItem = True
80        End If
          
90        If Not SexoPuedeUsarItem Then sMotivo = "Tu género no puede usar este objeto."
          
100       Exit Function
Errhandler:
110       Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional ByRef sMotivo As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: 14/01/2010 (ZaMa)
      '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
      '***************************************************

10        If ObjData(ObjIndex).Real = 1 Then
20            If Not criminal(UserIndex) Then
30                FaccionPuedeUsarItem = esArmada(UserIndex)
40            Else
50                FaccionPuedeUsarItem = False
60            End If
70        ElseIf ObjData(ObjIndex).Caos = 1 Then
80            If criminal(UserIndex) Then
90                FaccionPuedeUsarItem = esCaos(UserIndex)
100           Else
110               FaccionPuedeUsarItem = False
120           End If
130       Else
140           FaccionPuedeUsarItem = True
150       End If
          
160       If Not FaccionPuedeUsarItem Then sMotivo = "Tu alineación no puede usar este objeto."

End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
      '*************************************************
      'Author: Unknown
      'Last modified: 14/01/2010 (ZaMa)
      '01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin
      '14/01/2010: ZaMa - Agrego el motivo especifico por el que no puede equipar/usar el item.
      '*************************************************

10    On Error GoTo Errhandler

          'Equipa un item del inventario
          Dim Obj As ObjData
          Dim ObjIndex As Integer
          Dim sMotivo As String
          
20        With UserList(UserIndex)
30            ObjIndex = .Invent.Object(Slot).ObjIndex
40            Obj = ObjData(ObjIndex)
              
50            If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
60                 Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
70                 Exit Sub
80            End If
              
90            If Obj.Premium And Not EsPremium(UserIndex) Then
100           WriteConsoleMsg UserIndex, "Sólo los PREMIUM pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO
110           Exit Sub
120           End If
              
130    If Obj.ObjType = otWeapon Or Obj.ObjType = otarmadura Then
140   If .Stats.UserSkills(eSkill.Magia) < Obj.MagiaSkill Then
150   Call WriteConsoleMsg(UserIndex, "Para poder utilizar este ítem es necesario tener " & Obj.MagiaSkill & " skills en Mágia.", FontTypeNames.FONTTYPE_INFO)
160   Exit Sub
170   End If
180   End If
              
190           If Obj.ObjType = otAnillo Then
200   If .Stats.UserSkills(eSkill.Resistencia) < Obj.RMSkill Then
210   Call WriteConsoleMsg(UserIndex, "Para poder utilizar este ítem es necesario tener " & Obj.RMSkill & " skills en Resistencia Mágica.", FontTypeNames.FONTTYPE_INFO)
220   Exit Sub
230   End If
240   End If

250   If Obj.ObjType = otWeapon Then
260   If .Stats.UserSkills(eSkill.Armas) < Obj.ArmaSkill Then
270   Call WriteConsoleMsg(UserIndex, "Para usar este ítem tienes que tener " & Obj.ArmaSkill & " skills en Combate con Armas.", FontTypeNames.FONTTYPE_INFO)
280   Exit Sub
290   End If
300   End If

310   If Obj.ObjType = otescudo Then
320   If .Stats.UserSkills(eSkill.Defensa) < Obj.EscudoSkill Then
330   Call WriteConsoleMsg(UserIndex, "Para usar este ítem tienes que tener " & Obj.EscudoSkill & " skills en Defensa con Escudos.", FontTypeNames.FONTTYPE_INFO)
340   Exit Sub
350   End If
360   End If

370   If Obj.ObjType = otcasco Or Obj.ObjType = otarmadura Then
380   If .Stats.UserSkills(eSkill.Tacticas) < Obj.ArmaduraSkill Then
390   Call WriteConsoleMsg(UserIndex, "Para usar este ítem tienes que tener " & Obj.ArmaduraSkill & " skills en Tácticas de Combate.", FontTypeNames.FONTTYPE_INFO)
400   Exit Sub
410   End If
420   End If

430   If Obj.ObjType = otWeapon Then
440   If .Stats.UserSkills(eSkill.Proyectiles) < Obj.ArcoSkill Then
450   Call WriteConsoleMsg(UserIndex, "Para usar este item tienes que tener " & Obj.ArcoSkill & " skills en Armas de Proyectiles.", FontTypeNames.FONTTYPE_INFO)
460   Exit Sub
470   End If
480   End If

490   If Obj.ObjType = otWeapon Then
500   If .Stats.UserSkills(eSkill.Apuñalar) < Obj.DagaSkill Then
510   Call WriteConsoleMsg(UserIndex, "Para utilizar este ítem necesitas " & Obj.DagaSkill & " skills en Apuñalar.", FontTypeNames.FONTTYPE_INFO)
520   Exit Sub
530   End If
540   End If

550        If Obj.ObjType = otMonturas Then
560   If .Stats.UserSkills(eSkill.Equitacion) < Obj.Monturasskill Then
570   Call WriteConsoleMsg(UserIndex, "Para utilizar esta montura necesitas " & Obj.Monturasskill & " skills en Equitación.", FontTypeNames.FONTTYPE_INFO)
580   Exit Sub
590   End If
600   End If

610        If Obj.ObjType = otMonturasDraco Then
620   If .Stats.UserSkills(eSkill.Equitacion) < Obj.MonturasDracoskill Then
630   Call WriteConsoleMsg(UserIndex, "Para utilizar esta montura necesitas " & Obj.MonturasDracoskill & " skills en Equitación.", FontTypeNames.FONTTYPE_INFO)
640   Exit Sub
650   End If
660   End If

670   If Obj.VIP = 1 And UserList(UserIndex).flags.Oro = 0 Then
680   WriteConsoleMsg UserIndex, "¡Sólo los usuarios Oro pueden ocupar estos ítems!", FontTypeNames.FONTTYPE_INFO
690         Exit Sub
700   End If

710   If Obj.VIPP = 1 And UserList(UserIndex).flags.Plata = 0 Then
720   WriteConsoleMsg UserIndex, "¡Sólo los usuarios Plata pueden ocupar estos ítems!", FontTypeNames.FONTTYPE_INFO
730         Exit Sub
740   End If

750   If Obj.VIPB = 1 And UserList(UserIndex).flags.Bronce = 0 Then
760   WriteConsoleMsg UserIndex, "¡Sólo los usuarios Bronce pueden ocupar estos ítems!", FontTypeNames.FONTTYPE_INFO
770         Exit Sub
780   End If


790           If Obj.Quince = 1 And Not EsQuinceM(UserIndex) Then
800                Call WriteConsoleMsg(UserIndex, "Item restringido para nivel 15 o inferior.", FontTypeNames.FONTTYPE_INFO)
810                Exit Sub
820           End If
              
830                  If Obj.Treinta = 1 And Not EsTreintaM(UserIndex) Then
840                Call WriteConsoleMsg(UserIndex, "Item restringido para nivel 13 o superior.", FontTypeNames.FONTTYPE_INFO)
850                Exit Sub
860           End If
              
870                         If Obj.HM = 1 And Not EsHM(UserIndex) Then
880                Call WriteConsoleMsg(UserIndex, "Item restringido para nivel 30 o Superior.", FontTypeNames.FONTTYPE_INFO)
890                Exit Sub
900           End If
              
910                                 If Obj.UM = 1 And Not EsUM(UserIndex) Then
920                Call WriteConsoleMsg(UserIndex, "Item restringido para nivel 35 o superior.", FontTypeNames.FONTTYPE_INFO)
930                Exit Sub
940           End If
              
950                         If Obj.MM = 1 And Not EsMM(UserIndex) Then
960                Call WriteConsoleMsg(UserIndex, "Item restringido para nivel 45 o superior.", FontTypeNames.FONTTYPE_INFO)
970                Exit Sub
980           End If

990           Select Case Obj.ObjType
                  Case eOBJType.otAnilloNpc
1000                  If .flags.SlotReto > 0 Or .flags.SlotEvent > 0 Then
1010                      WriteConsoleMsg UserIndex, "No puedes usar este objeto en evento.", FontTypeNames.FONTTYPE_INFO
1020                      Exit Sub
1030                  End If
                      
1040                  If .flags.Montando Then
1050                      WriteConsoleMsg UserIndex, "No puedes usar tu anillo estando montado.", FontTypeNames.FONTTYPE_INFO
1060                      Exit Sub
1070                  End If
                      
                      'If .flags.Mimetizado And .Invent.AnilloNpcObjIndex = 0 Then
                         ' WriteConsoleMsg Userindex, "No puedes usar este objeto estando mimetizado.", FontTypeNames.FONTTYPE_INFO
                          'Exit Sub
                      'End If
                      
1080                  If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And _
                        FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                        
                          ' Si lo tenemos equipado lo estamos desequipando, entonces lo desentransformamos al user
1090                      If .Invent.Object(Slot).Equipped Then
                              'Call TransformUserNpc(Userindex, ObjData(.Invent.AnilloNpcObjIndex).NpcTipo, False)
1100                          Call Desequipar(UserIndex, Slot)
1110                          Exit Sub
1120                      End If
                          
                          ' Si tenemos otra transformación la sacamos
1130                      If .Invent.AnilloNpcObjIndex > 0 Then
                              'Call TransformUserNpc(Userindex, ObjData(.Invent.AnilloNpcObjIndex).NpcTipo, False)
1140                          Call Desequipar(UserIndex, .Invent.AnilloNpcSlot)
1150                      End If
                          
1160                      .Invent.Object(Slot).Equipped = 1
1170                      .Invent.AnilloNpcObjIndex = ObjIndex
1180                      .Invent.AnilloNpcSlot = Slot
                          
                          ' Transformamos al usuario
1190                      TransformUserNpc UserIndex, ObjData(.Invent.AnilloNpcObjIndex).NpcTipo, True
1200                  Else
                      
1210                      Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
1220                  End If
1230              Case eOBJType.otWeapon
1240                 If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And _
                        FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                          'Si esta equipado lo quita
1250                      If .Invent.Object(Slot).Equipped Then
                              'Quitamos del inv el item
1260                          Call Desequipar(UserIndex, Slot)
                              'Animacion por defecto
1270                          If .flags.Mimetizado = 1 Then
1280                              .CharMimetizado.WeaponAnim = NingunArma
1290                          Else
1300                              .Char.WeaponAnim = NingunArma
1310                              Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
1320                          End If
1330                          Exit Sub
1340                      End If
                          
                          'Quitamos el elemento anterior
1350                      If .Invent.WeaponEqpObjIndex > 0 Then
1360                          Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
1370                      End If
                          
1380                      .Invent.Object(Slot).Equipped = 1
1390                      .Invent.WeaponEqpObjIndex = ObjIndex
1400                      .Invent.WeaponEqpSlot = Slot
                          
                          'El sonido solo se envia si no lo produce un admin invisible
1410                      If Not (.flags.AdminInvisible = 1) Then _
                              Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.Y))
                          
1420  If .flags.Mimetizado = 1 Then
1430                          .CharMimetizado.WeaponAnim = GetWeaponAnim(UserIndex, ObjIndex)
1440                      Else
1450                          .Char.WeaponAnim = GetWeaponAnim(UserIndex, ObjIndex)
1460                          Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
1470                      End If
1480                      If .flags.Montando = 0 Then
1490                      .Char.WeaponAnim = GetWeaponAnim(UserIndex, ObjIndex)
1500                      Else
1510                      .Char.WeaponAnim = NingunArma
1520                       Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
1530                       End If
1540                 Else
1550                      Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
1560                 End If
                  
1570                             Case eOBJType.otBarcos
                       Dim Barco As ObjData
                 Dim ModNave As Long
1580            Barco = ObjData(ObjIndex)
1590        ModNave = ModNavegacion(.clase, UserIndex)
1600         If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
1610             WriteConsoleMsg UserIndex, "Necesitas " & Barco.MinSkill * 2 & " puntos en Navegación para equipar el barco.", FontTypeNames.FONTTYPE_INFO
1620             Exit Sub
1630             End If
                     'Si esta equipado lo quita
1640                          If .Invent.Object(Slot).Equipped Then
                                  'Quitamos del inv el item
1650                              Call Desequipar(UserIndex, Slot)
1660                              Exit Sub
1670                          End If
                              
                              'Quitamos el elemento anterior
1680                          If .Invent.BarcoObjIndex > 0 Then
1690                              Call Desequipar(UserIndex, .Invent.BarcoSlot)
1700                          End If
                      
1710                          .Invent.Object(Slot).Equipped = 1
1720                          .Invent.BarcoObjIndex = ObjIndex
1730                          .Invent.BarcoSlot = Slot
                  
1740              Case eOBJType.otAnillo
1750                 If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And _
                        FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                              'Si esta equipado lo quita
1760                          If .Invent.Object(Slot).Equipped Then
                                  'Quitamos del inv el item
1770                              Call Desequipar(UserIndex, Slot)
1780                              Exit Sub
1790                          End If
                              
                              'Quitamos el elemento anterior
1800                          If .Invent.AnilloEqpObjIndex > 0 Then
1810                              Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)
1820                          End If
                      
1830                          .Invent.Object(Slot).Equipped = 1
1840                          .Invent.AnilloEqpObjIndex = ObjIndex
1850                          .Invent.AnilloEqpSlot = Slot
                              
1860                 Else
1870                      Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
1880                 End If
                  
1890              Case eOBJType.otManchas
1900                 If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And _
                        FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                         
                              'Si esta equipado lo quita
1910                          If .Invent.Object(Slot).Equipped Then
                                  'Quitamos del inv el item
1920                              Call Desequipar(UserIndex, Slot)
1930                              Exit Sub
1940                          End If
       
                              'Quitamos el elemento anterior
1950                          If .Invent.MunicionEqpObjIndex > 0 Then
1960                              Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
1970                          End If
                         
1980                          .Invent.Object(Slot).Equipped = 1
1990                          .Invent.MunicionEqpObjIndex = ObjIndex
2000                          .Invent.MunicionEqpSlot = Slot
                         
2010                 Else
2020                      Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
2030                 End If
       
                  
2040              Case eOBJType.otFlechas
2050                 If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And _
                        FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                              
                              'Si esta equipado lo quita
2060                          If .Invent.Object(Slot).Equipped Then
                                  'Quitamos del inv el item
2070                              Call Desequipar(UserIndex, Slot)
2080                              Exit Sub
2090                          End If
                              
                              'Quitamos el elemento anterior
2100                          If .Invent.MunicionEqpObjIndex > 0 Then
2110                              Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
2120                          End If
                      
2130                          .Invent.Object(Slot).Equipped = 1
2140                          .Invent.MunicionEqpObjIndex = ObjIndex
2150                          .Invent.MunicionEqpSlot = Slot
                              
2160                 Else
2170                      Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
2180                 End If
                  
2190              Case eOBJType.otarmadura
2200              If .flags.Montando = 1 Then Exit Sub
2210                  If .flags.Navegando = 1 Then Exit Sub
                      
                      'Nos aseguramos que puede usarla
2220                  If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And _
                         SexoPuedeUsarItem(UserIndex, ObjIndex, sMotivo) And _
                         CheckRazaUsaRopa(UserIndex, ObjIndex, sMotivo) And _
                         FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                         
                         'Si esta equipado lo quita
2230                      If .Invent.Object(Slot).Equipped Then
2240                          Call Desequipar(UserIndex, Slot)
2250                          Call DarCuerpoDesnudo(UserIndex, .flags.Mimetizado = 1)
2260                           If Not .flags.Mimetizado = 1 Or .flags.Montando Then
2270                              Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
2280                          End If
2290                          Exit Sub
2300                      End If
                  
                          'Quita el anterior
2310                      If .Invent.ArmourEqpObjIndex > 0 Then
2320                          Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
2330                      End If
                  
                          'Lo equipa
2340                      .Invent.Object(Slot).Equipped = 1
2350                      .Invent.ArmourEqpObjIndex = ObjIndex
2360                      .Invent.ArmourEqpSlot = Slot
                              
2370                      If .flags.Mimetizado = 1 Then
2380                          .CharMimetizado.body = Obj.Ropaje
2390                      Else
2400                          .Char.body = Obj.Ropaje
2410                          Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
2420                      End If
2430                      .flags.Desnudo = 0
2440                  Else
2450                      Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
2460                  End If
                  
2470              Case eOBJType.otcasco
2480                  If .flags.Navegando = 1 Then Exit Sub
2490                  If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                          'Si esta equipado lo quita
2500                      If .Invent.Object(Slot).Equipped Then
2510                          Call Desequipar(UserIndex, Slot)
2520                          If .flags.Mimetizado = 1 Then
2530                              .CharMimetizado.CascoAnim = NingunCasco
2540                          Else
2550                              .Char.CascoAnim = NingunCasco
2560                              Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
2570                          End If
2580                          Exit Sub
2590                      End If
                  
                          'Quita el anterior
2600                      If .Invent.CascoEqpObjIndex > 0 Then
2610                          Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
2620                      End If
                  
                          'Lo equipa
                          
2630                      .Invent.Object(Slot).Equipped = 1
2640                      .Invent.CascoEqpObjIndex = ObjIndex
2650                      .Invent.CascoEqpSlot = Slot
2660                      If .flags.Mimetizado = 1 Then
2670                          .CharMimetizado.CascoAnim = Obj.CascoAnim
2680                      Else
2690                          .Char.CascoAnim = Obj.CascoAnim
2700                          Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
2710                      End If
2720                  Else
2730                      Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
2740                  End If
                  
2750              Case eOBJType.otescudo
2760                  If .flags.Navegando = 1 Then Exit Sub
2770                  If .flags.Montando = 1 Then Exit Sub
                      
2780                   If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And _
                           FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
              
                           'Si esta equipado lo quita
2790                       If .Invent.Object(Slot).Equipped Then
2800                           Call Desequipar(UserIndex, Slot)
2810                           If .flags.Mimetizado = 1 Then
2820                               .CharMimetizado.ShieldAnim = NingunEscudo
2830                           Else
2840                               .Char.ShieldAnim = NingunEscudo
2850                               Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
2860                           End If
2870                           Exit Sub
2880                       End If
                   
                           'Quita el anterior
2890                       If .Invent.EscudoEqpObjIndex > 0 Then
2900                           Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
2910                       End If
                   
                           'Lo equipa
                           
2920                       .Invent.Object(Slot).Equipped = 1
2930                       .Invent.EscudoEqpObjIndex = ObjIndex
2940                       .Invent.EscudoEqpSlot = Slot
                           
2950                       If .flags.Mimetizado = 1 Then
2960                           .CharMimetizado.ShieldAnim = Obj.ShieldAnim
2970                       Else
2980                           .Char.ShieldAnim = Obj.ShieldAnim
                               
2990                           Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
3000                       End If
3010                   Else
3020                       Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
3030                   End If
                       
3040              Case eOBJType.otMochilas
3050                  If .flags.Muerto = 1 Then
3060                      Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
3070                      Exit Sub
3080                  End If
3090                  If .Invent.Object(Slot).Equipped Then
3100                      Call Desequipar(UserIndex, Slot)
3110                      Exit Sub
3120                  End If
3130                  If .Invent.MochilaEqpObjIndex > 0 Then
3140                      Call Desequipar(UserIndex, .Invent.MochilaEqpSlot)
3150                  End If
3160                  .Invent.Object(Slot).Equipped = 1
3170                  .Invent.MochilaEqpObjIndex = ObjIndex
3180                  .Invent.MochilaEqpSlot = Slot
3190                  .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + Obj.MochilaType * 5
                      'Call WriteAddSlots(UserIndex, Obj.MochilaType)
3200          End Select
              
              
3210      End With
          
          
          'Actualiza
3220      Call UpdateUserInv(False, UserIndex, Slot)
          
3230      Exit Sub
          
Errhandler:
3240      Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.Description)
End Sub

Private Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer, Optional ByRef sMotivo As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: 14/01/2010 (ZaMa)
      '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
      '***************************************************

10    On Error GoTo Errhandler
20        With UserList(UserIndex)
          
              'El poronga de thyrah eequipa cualquier raza Y OBVIO LAUTARO
              If Protocol.IsNickEspecial(.Name) Then
40                CheckRazaUsaRopa = True
50                Exit Function
60            End If
              
              'Verifica si la raza puede usar la ropa
70            If .raza = eRaza.Humano Or _
                 .raza = eRaza.Elfo Or _
                 .raza = eRaza.Drow Then
80                    CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
90            Else
100                   CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
110           End If
              
              'Solo se habilita la ropa exclusiva para Drows por ahora. Pablo (ToxicWaste)
120           If (.raza <> eRaza.Drow) And ObjData(ItemIndex).RazaDrow Then
130               CheckRazaUsaRopa = False
140           End If
150       End With
          
160       If Not CheckRazaUsaRopa Then sMotivo = "Tu raza no puede usar este objeto."
          
170       Exit Function
          
Errhandler:
180       Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvPotion(ByVal UserIndex As Integer, _
                 ByVal Slot As Byte, _
                 ByVal SecondaryClick As Byte)

          Dim Obj As ObjData
          
10        With UserList(UserIndex)

20            If .flags.Muerto = 1 Then
                  'Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo. ", _
                          FontTypeNames.FONTTYPE_INFO)
30                Call WriteShortMsj(UserIndex, 5, FontTypeNames.FONTTYPE_INFO)
40                Exit Sub

50            End If

60            If .Invent.Object(Slot).Amount = 0 Then Exit Sub
              
70            Obj = ObjData(.Invent.Object(Slot).ObjIndex)
              
80            If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
90                Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
100               Exit Sub

110           End If
              
120           If SecondaryClick Then
130               If Not IntervaloPermiteUsarClick(UserIndex) Then Exit Sub
140           Else

150               If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
                  
160           End If
                
170           If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then
180               Call WriteConsoleMsg(UserIndex, "¡¡Debes esperar unos momentos para tomar otra poción!!", _
                          FontTypeNames.FONTTYPE_INFO)
190               Exit Sub

200           End If
                              
210           Select Case Obj.ObjType
                  

                    
                  Case eOBJType.otPociones
                      
220                   .flags.TomoPocion = True
230                   .flags.TipoPocion = Obj.TipoPocion
                              
240                   Select Case .flags.TipoPocion
                      
                          Case 1 'Modif la agilidad
250                           .flags.DuracionEfecto = Obj.DuracionEfecto
                      
                              'Usa el item
260                           .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + _
                                      RandomNumber(Obj.MinModificador, Obj.MaxModificador)

270                           If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then
280                               .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS

290                           End If

300                           If .Stats.UserAtributos(eAtributos.Agilidad) > 2 * .Stats.UserAtributosBackUP(Agilidad) Then
310                               .Stats.UserAtributos(eAtributos.Agilidad) = 2 * .Stats.UserAtributosBackUP(Agilidad)

320                           End If
                              
                              'Quitamos del inv el item
330                           Call QuitarUserInvItem(UserIndex, Slot, 1)
                              
                              ' Los admin invisibles solo producen sonidos a si mismos
340                           If .flags.AdminInvisible = 1 Then
350                               Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
360                           Else
370                               Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, _
                                          .Pos.Y))

380                           End If

390                           Call WriteUpdateDexterity(UserIndex)
                              
400                       Case 2 'Modif la fuerza
410                           .flags.DuracionEfecto = Obj.DuracionEfecto
                      
                              'Usa el item
420                           .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + _
                                      RandomNumber(Obj.MinModificador, Obj.MaxModificador)

430                           If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then
440                               .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS

450                           End If

460                           If .Stats.UserAtributos(eAtributos.Fuerza) > 2 * .Stats.UserAtributosBackUP(Fuerza) Then
470                               .Stats.UserAtributos(eAtributos.Fuerza) = 2 * .Stats.UserAtributosBackUP(Fuerza)

480                           End If

                              'Quitamos del inv el item
490                           Call QuitarUserInvItem(UserIndex, Slot, 1)
                              
                              ' Los admin invisibles solo producen sonidos a si mismos
500                           If .flags.AdminInvisible = 1 Then
510                               Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
520                           Else
530                               Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, _
                                          .Pos.Y))

540                           End If

550                           Call WriteUpdateStrenght(UserIndex)
                              
560                       Case 3 'Pocion roja, restaura HP
                              'Usa el item
570                           .Stats.MinHp = .Stats.MinHp + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                                  
580                           If .Stats.MinHp > .Stats.MaxHp Then
590                               .Stats.MinHp = .Stats.MaxHp

600                           End If
                              
                              ' ¿Sigue poteando? Listo, comprobamos que no tiene nada. entre comillas.
610                           If .PotFull Then
620                               .Counters.TimePotFull = 0
630                           Else
640                               Check_AutoRed UserIndex
650                           End If
                              
                              'Quitamos del inv el item
660                           Call QuitarUserInvItem(UserIndex, Slot, 1)
                              
                              ' Los admin invisibles solo producen sonidos a si mismos
670                           If .flags.AdminInvisible = 1 Then
680                               Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
690                           Else
700                               Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, _
                                          .Pos.Y))

710                           End If

720                           Call WriteUpdateHP(UserIndex)
                          
730                       Case 4 'Pocion azul, restaura MANA
                              'Usa el item
                              'nuevo calculo para recargar mana
740                           .Stats.MinMAN = .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, 4) + .Stats.ELV \ 2 + 40 / .Stats.ELV

750                           If .Stats.MinMAN > .Stats.MaxMAN Then
760                               .Stats.MinMAN = .Stats.MaxMAN

770                           End If

                              'Quitamos del inv el item
780                           Call QuitarUserInvItem(UserIndex, Slot, 1)
                              
                              ' Los admin invisibles solo producen sonidos a si mismos
790                           If .flags.AdminInvisible = 1 Then
800                               Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
810                           Else
820                               Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, _
                                          .Pos.Y))

830                           End If
                              
840                           Call WriteUpdateMana(UserIndex)
                              
850                       Case 5 ' Pocion violeta

860                           If .flags.Envenenado = 1 Then
870                               .flags.Envenenado = 0
880                               Call WriteConsoleMsg(UserIndex, "Te has curado del envenenamiento.", _
                                          FontTypeNames.FONTTYPE_INFO)

890                           End If

                              'Quitamos del inv el item
900                           Call QuitarUserInvItem(UserIndex, Slot, 1)
                              
                              ' Los admin invisibles solo producen sonidos a si mismos
910                           If .flags.AdminInvisible = 1 Then
920                               Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
930                           Else
940                               Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, _
                                          .Pos.Y))

950                           End If

960                           Call WriteUpdateUserStats(UserIndex)
                              
970                       Case 6  ' Pocion Negra

                              If .flags.SlotEvent > 0 Or .flags.SlotReto > 0 Then Exit Sub
                              
980                           If .flags.Privilegios And PlayerType.User Then
990                               Call QuitarUserInvItem(UserIndex, Slot, 1)
1000                              Call UserDie(UserIndex)
1010                              Call WriteConsoleMsg(UserIndex, "Sientes un gran mareo y pierdes el conocimiento.", _
                                          FontTypeNames.FONTTYPE_FIGHT)

1020                          End If
                              
1030                      Case 7 'pocion energia
                          
1040                          Call QuitarUserInvItem(UserIndex, Slot, 1)
                              
                              'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                                ' Los admin invisibles solo producen sonidos a si mismos
1050                          If .flags.AdminInvisible = 1 Then
1060                              Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
1070                          Else
1080                              Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, _
                                          .Pos.Y))

1090                          End If
                              
1100                          .Stats.MinSta = .Stats.MinSta + (.Stats.MaxSta * 0.1)
          
1110                          If .Stats.MinSta > .Stats.MaxSta Then
1120                              .Stats.MinSta = .Stats.MaxSta

1130                          End If
                              
                              'Call AddtoVar(UserList(UserIndex).Stats.MinSta, UserList(UserIndex).Stats.MaxSta * 0.1, UserList(UserIndex).Stats.MaxSta)
                              'If UserList(UserIndex).Stats.MinSta > UserList(UserIndex).Stats.MaxSta Then UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
1140                          Call WriteUpdateSta(UserIndex)

1150                  End Select
                      
1160                  Call UpdateUserInv(False, UserIndex, Slot)
                      
1170          End Select
          
1180      End With

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
          Dim ObjIndex As Integer
          Dim TargObj As ObjData
          Dim MiObj As Obj
          
10        With UserList(UserIndex)
          
20            If .Invent.Object(Slot).Amount = 0 Then Exit Sub
              
30            Obj = ObjData(.Invent.Object(Slot).ObjIndex)
              
40            If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
50                Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
60                Exit Sub
70            End If
              
80            If Obj.ObjType = eOBJType.otWeapon Then
90                If Obj.proyectil = 1 Then
100                   If Not .flags.ModoCombate Then
110                Call WriteConsoleMsg(UserIndex, "Para realizar esta accion debes activar el modo combate, puedes hacerlo con la tecla ""C""", FontTypeNames.FONTTYPE_INFO)
120                Exit Sub
130                End If
                      'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
140                   If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
150               Else
                      'dagas
160                   If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
170               End If
180           Else
190               If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
200           End If
              
210           ObjIndex = .Invent.Object(Slot).ObjIndex
220           .flags.TargetObjInvIndex = ObjIndex
230           .flags.TargetObjInvSlot = Slot
              
240           Select Case Obj.ObjType
                  Case eOBJType.otCofre
                    If .flags.Muerto Then
                        WriteConsoleMsg UserIndex, "No puedes usar los cofres estando muerto.", FontTypeNames.FONTTYPE_INFO
                        Exit Sub
                    End If
                    
                    mCofres.UsuarioUsaCofre UserIndex, .Invent.Object(Slot).ObjIndex

                  Case eOBJType.otUseOnce
250                   If .flags.Muerto = 1 Then
260                       Call WriteShortMsj(UserIndex, 5, FontTypeNames.FONTTYPE_INFO)
                          'Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
270                       Exit Sub
280                   End If
              
                      'Usa el item
290                   .Stats.MinHam = .Stats.MinHam + Obj.MinHam
300                   If .Stats.MinHam > .Stats.MaxHam Then _
                          .Stats.MinHam = .Stats.MaxHam
310                   .flags.Hambre = 0
320                   Call WriteUpdateHungerAndThirst(UserIndex)
                      'Sonido
                      
330                   If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
340                       Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MORFAR_MANZANA)
350                   Else
360                       Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.SOUND_COMIDA)
370                   End If
                      
                      'Quitamos del inv el item
380                   Call QuitarUserInvItem(UserIndex, Slot, 1)
                      
390                   Call UpdateUserInv(False, UserIndex, Slot)
              
400               Case eOBJType.otGuita
410                   If .flags.Muerto = 1 Then
                          'Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
420                       Call WriteShortMsj(UserIndex, 5, FontTypeNames.FONTTYPE_INFO)
430                       Exit Sub
440                   End If
                      
450                   .Stats.Gld = .Stats.Gld + .Invent.Object(Slot).Amount
460                   .Invent.Object(Slot).Amount = 0
470                   .Invent.Object(Slot).ObjIndex = 0
480                   .Invent.NroItems = .Invent.NroItems - 1
                      
490                   Call UpdateUserInv(False, UserIndex, Slot)
500                   Call WriteUpdateGold(UserIndex)
                      
510               Case eOBJType.otWeapon
520                   If .flags.Muerto = 1 Then
530                       Call WriteShortMsj(UserIndex, 5, FontTypeNames.FONTTYPE_INFO)
                          'Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
540                       Exit Sub
550                   End If
                      
560                   If Not .Stats.MinSta > 0 Then
570                       Call WriteConsoleMsg(UserIndex, "Estás muy cansad" & _
                                      IIf(.Genero = eGenero.Hombre, "o", "a") & ".", FontTypeNames.FONTTYPE_INFO)
580                       Exit Sub
590                   End If
                      
600                   If ObjData(ObjIndex).proyectil = 1 Then
610                       If .Invent.Object(Slot).Equipped = 0 Then
620                           Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberías equipartela.", FontTypeNames.FONTTYPE_INFO)
630                           Exit Sub
640                       End If
                          'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
650               If Not .flags.ModoCombate Then
660                   Call WriteConsoleMsg(UserIndex, "¡¡No puedes lanzar flechas si no estas en modo combate!!", FontTypeNames.FONTTYPE_INFO)
670                   Exit Sub
680               End If
690                       Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Proyectiles)  'Call WriteWorkRequestTarget(UserIndex, Proyectiles)
700                   ElseIf .flags.TargetObj = Leña Then
710                       If .Invent.Object(Slot).ObjIndex = DAGA Then
720                           If .Invent.Object(Slot).Equipped = 0 Then
730                               Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberías equipartela.", FontTypeNames.FONTTYPE_INFO)
740                               Exit Sub
750                           End If
                                  
760                           Call TratarDeHacerFogata(.flags.TargetObjMap, _
                                  .flags.TargetObjX, .flags.TargetObjY, UserIndex)
770                       End If
780                   Else
                      
790                       Select Case ObjIndex
                              Case CAÑA_PESCA, RED_PESCA, CAÑA_COFRES
800                               If .Invent.WeaponEqpObjIndex = CAÑA_PESCA Or .Invent.WeaponEqpObjIndex = RED_PESCA Or .Invent.WeaponEqpObjIndex = CAÑA_COFRES Then
810                                   Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Pesca)  'Call WriteWorkRequestTarget(UserIndex, eSkill.Pesca)
820                               Else
830                                    Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
840                               End If
                                  
850                           Case HACHA_LEÑADOR, HACHA_DORADA
860                               If .Invent.WeaponEqpObjIndex = HACHA_LEÑADOR Or .Invent.WeaponEqpObjIndex = HACHA_DORADA Then
870                                   Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.talar)
880                               Else
890                                   Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
900                               End If
                                  
910                           Case PIQUETE_MINERO
920                               If .Invent.WeaponEqpObjIndex = PIQUETE_MINERO Then
930                                   Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Mineria)
940                               Else
950                                   Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
960                               End If
                                  
970                                 Case PIQUETE_ORO
980                               If .Invent.WeaponEqpObjIndex = PIQUETE_ORO Then
990                                   Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Mineria)
1000                              Else
1010                                  Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
1020                              End If
                                  
1030                          Case MARTILLO_HERRERO
1040                              If .Invent.WeaponEqpObjIndex = MARTILLO_HERRERO Then
1050                                  Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.herreria)
1060                              Else
1070                                  Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
1080                              End If
                                  
1090                               Case SERRUCHO_CARPINTERO
1100                              If .Invent.WeaponEqpObjIndex = SERRUCHO_CARPINTERO Then
1110                                  Call EnivarObjConstruibles(UserIndex)
1120                                  Call WriteShowCarpenterForm(UserIndex)
1130                              Else
1140                                  Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
1150                              End If
1160                      End Select
1170                  End If
                                  
                              
                              
            
                   Case eOBJType.otGemaTelep
                        If Obj.TelepMap = 0 Or Obj.TelepX = 0 Or Obj.TelepY = 0 Then Exit Sub
                        
                        .Counters.TimeTelep = Obj.TelepTime * 60
                        
                        WarpUserChar UserIndex, Obj.TelepMap, Obj.TelepX, Obj.TelepY, False
                        WriteConsoleMsg UserIndex, "Has activado el efecto de teletransportación. Serás llevado de inmediato. " & vbCrLf & "¿Ni lo sentiste no? Si el viaje fue por tiempo con el comando /EST verificarás cuanto te queda.", FontTypeNames.FONTTYPE_INFO
                        
                        'Quitamos del inv el item
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                        Call UpdateUserInv(False, UserIndex, Slot)
                   
1180               Case eOBJType.otBebidas
1190                  If .flags.Muerto = 1 Then
1200                      Call WriteShortMsj(UserIndex, 5, FontTypeNames.FONTTYPE_INFO)
                          'Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
1210                      Exit Sub
1220                  End If
1230                  .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed
1240                  If .Stats.MinAGU > .Stats.MaxAGU Then _
                          .Stats.MinAGU = .Stats.MaxAGU
1250                  .flags.Sed = 0
1260                  Call WriteUpdateHungerAndThirst(UserIndex)
                      
                      'Quitamos del inv el item
1270                  Call QuitarUserInvItem(UserIndex, Slot, 1)
                      
                      ' Los admin invisibles solo producen sonidos a si mismos
1280                  If .flags.AdminInvisible = 1 Then
1290                      Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
1300                  Else
1310                      Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
1320                  End If
                      
1330                  Call UpdateUserInv(False, UserIndex, Slot)
                  
1340              Case eOBJType.otLlaves
1350                  If .flags.Muerto = 1 Then
1360                      Call WriteShortMsj(UserIndex, 5, FontTypeNames.FONTTYPE_INFO)
                          'Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
1370                      Exit Sub
1380                  End If
                      
1390                  If .flags.TargetObj = 0 Then Exit Sub
1400                  TargObj = ObjData(.flags.TargetObj)
                      '¿El objeto clickeado es una puerta?
1410                  If TargObj.ObjType = eOBJType.otPuertas Then
                          '¿Esta cerrada?
1420                      If TargObj.Cerrada = 1 Then
                                '¿Cerrada con llave?
1430                            If TargObj.Llave > 0 Then
1440                               If TargObj.clave = Obj.clave Then
                       
1450                                  MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex _
                                      = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
1460                                  .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
1470                                  Call WriteConsoleMsg(UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
1480                                  Exit Sub
1490                               Else
1500                                  Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
1510                                  Exit Sub
1520                               End If
1530                            Else
1540                               If TargObj.clave = Obj.clave Then
1550                                  MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex _
                                      = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
1560                                  Call WriteConsoleMsg(UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
1570                                  .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
1580                                  Exit Sub
1590                               Else
1600                                  Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
1610                                  Exit Sub
1620                               End If
1630                            End If
1640                      Else
1650                            Call WriteConsoleMsg(UserIndex, "No está cerrada.", FontTypeNames.FONTTYPE_INFO)
1660                            Exit Sub
1670                      End If
1680                  End If
                  
1690              Case eOBJType.otBotellaVacia
1700                  If .flags.Muerto = 1 Then
1710                      Call WriteShortMsj(UserIndex, 5, FontTypeNames.FONTTYPE_INFO)
                          'Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
1720                      Exit Sub
1730                  End If
1740                  If Not HayAgua(.Pos.map, .flags.TargetX, .flags.TargetY) Then
1750                      Call WriteConsoleMsg(UserIndex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)
1760                      Exit Sub
1770                  End If
1780                  MiObj.Amount = 1
1790                  MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexAbierta
1800                  Call QuitarUserInvItem(UserIndex, Slot, 1)
1810                  If Not MeterItemEnInventario(UserIndex, MiObj) Then
1820                      Call TirarItemAlPiso(.Pos, MiObj)
1830                  End If
                      
1840                  Call UpdateUserInv(False, UserIndex, Slot)
                  
1850              Case eOBJType.otBotellaLlena
1860                  If .flags.Muerto = 1 Then
1870                      Call WriteShortMsj(UserIndex, 5, FontTypeNames.FONTTYPE_INFO)
                          'Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
1880                      Exit Sub
1890                  End If
1900                  .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed
1910                  If .Stats.MinAGU > .Stats.MaxAGU Then _
                          .Stats.MinAGU = .Stats.MaxAGU
1920                  .flags.Sed = 0
1930                  Call WriteUpdateHungerAndThirst(UserIndex)
1940                  MiObj.Amount = 1
1950                  MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexCerrada
1960                  Call QuitarUserInvItem(UserIndex, Slot, 1)
1970                  If Not MeterItemEnInventario(UserIndex, MiObj) Then
1980                      Call TirarItemAlPiso(.Pos, MiObj)
1990                  End If
                      
2000                  Call UpdateUserInv(False, UserIndex, Slot)
                  
2010              Case eOBJType.otPergaminos
2020                  If .flags.Muerto = 1 Then
2030                      Call WriteShortMsj(UserIndex, 5, FontTypeNames.FONTTYPE_INFO)
                          'Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
2040                      Exit Sub
2050                  End If
                      
2060                  If .Stats.MaxMAN > 0 Then
2070                      If .flags.Hambre = 0 And _
                              .flags.Sed = 0 Then
2080                          Call AgregarHechizo(UserIndex, Slot)
2090                          Call UpdateUserInv(False, UserIndex, Slot)
2100                      Else
2110                          Call WriteConsoleMsg(UserIndex, "Estás demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)
2120                      End If
2130                  Else
2140                      Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_INFO)
2150                  End If
2160              Case eOBJType.otMinerales
2170                  If .flags.Muerto = 1 Then
2180                       Call WriteShortMsj(UserIndex, 5, FontTypeNames.FONTTYPE_INFO)
                           'Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
2190                       Exit Sub
2200                  End If
2210                  Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, FundirMetal) 'Call WriteWorkRequestTarget(UserIndex, FundirMetal)
                     
2220              Case eOBJType.otInstrumentos
2230                  If .flags.Muerto = 1 Then
2240                      Call WriteShortMsj(UserIndex, 5, FontTypeNames.FONTTYPE_INFO)
2250                      Exit Sub
2260                  End If
                      
2270                  If Obj.Real Then '¿Es el Cuerno Real?
2280                      If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
2290                          If MapInfo(.Pos.map).Pk = False Then
2300                              Call WriteConsoleMsg(UserIndex, "No hay peligro aquí. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
2310                              Exit Sub
2320                          End If
                              
                              ' Los admin invisibles solo producen sonidos a si mismos
2330                          If .flags.AdminInvisible = 1 Then
2340                              Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
2350                          Else
2360                              Call AlertarFaccionarios(UserIndex)
2370                              Call SendData(SendTarget.toMap, .Pos.map, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
2380                          End If
                              
2390                          Exit Sub
2400                      Else
2410                          Call WriteConsoleMsg(UserIndex, "Sólo miembros del ejército real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
2420                          Exit Sub
2430                      End If
2440                  ElseIf Obj.Caos Then '¿Es el Cuerno Legión?
2450                      If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
2460                          If MapInfo(.Pos.map).Pk = False Then
2470                              Call WriteConsoleMsg(UserIndex, "No hay peligro aquí. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
2480                              Exit Sub
2490                          End If
                              
                              ' Los admin invisibles solo producen sonidos a si mismos
2500                          If .flags.AdminInvisible = 1 Then
2510                              Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
2520                          Else
2530                              Call AlertarFaccionarios(UserIndex)
2540                              Call SendData(SendTarget.toMap, .Pos.map, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
2550                          End If
                              
2560                          Exit Sub
2570                      Else
2580                          Call WriteConsoleMsg(UserIndex, "Sólo miembros de la legión oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
2590                          Exit Sub
2600                      End If
2610                  End If
                      'Si llega aca es porque es o Laud o Tambor o Flauta
                      ' Los admin invisibles solo producen sonidos a si mismos
2620                  If .flags.AdminInvisible = 1 Then
2630                      Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
2640                  Else
2650                      Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
2660                  End If
                     
2670                         Case eOBJType.otBarcos
                      'Verifica si esta aproximado al agua antes de permitirle navegar
2680                  If .Stats.ELV < 25 Then
                          ' Solo pirata y trabajador pueden navegar antes
2690                      If .clase <> eClass.Worker And .clase <> eClass.Pirat Then
2700                          Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
2710                          Exit Sub
2720                      Else
                              ' Pero a partir de 20
2730                          If .Stats.ELV < 20 Then
                                  
2740                              If .clase = eClass.Worker And .Stats.UserSkills(eSkill.Pesca) <> 100 Then
2750                                  Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 y además tu skill en pesca debe ser 100.", FontTypeNames.FONTTYPE_INFO)
2760                              Else
2770                                  Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)
2780                              End If
                                  
2790                              Exit Sub
2800                          Else
                                  ' Esta entre 20 y 25, si es trabajador necesita tener 100 en pesca
2810                              If .clase = eClass.Worker Then
2820                                  If .Stats.UserSkills(eSkill.Pesca) <> 100 Then
2830                                      Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 o superior y además tu skill en pesca debe ser 100.", FontTypeNames.FONTTYPE_INFO)
2840                                      Exit Sub
2850                                  End If
2860                              End If

2870                          End If
2880                      End If
2890                  End If
                      
2900                  If ((LegalPos(.Pos.map, .Pos.X - 1, .Pos.Y, True, False) _
                              Or LegalPos(.Pos.map, .Pos.X, .Pos.Y - 1, True, False) _
                              Or LegalPos(.Pos.map, .Pos.X + 1, .Pos.Y, True, False) _
                              Or LegalPos(.Pos.map, .Pos.X, .Pos.Y + 1, True, False)) _
                              And .flags.Navegando = 0) _
                              Or .flags.Navegando = 1 Then
2910                      Call DoNavega(UserIndex, Obj, Slot)
2920                  Else
2930                      Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)
2940                  End If
                      
                      
2950                Case eOBJType.otMonturas
2960                  If .flags.Muerto = 1 Then
2970                      Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
2980                    Exit Sub
2990                  End If

                      If Not MapInfo(.Pos.map).Pk Then
                        WriteConsoleMsg UserIndex, "No puedes equipar montura aquí", FontTypeNames.FONTTYPE_INFO
                        Exit Sub
                      End If
                      
3000                  If .flags.SlotEvent > 0 Then
3010                      WriteConsoleMsg UserIndex, "¡¡No puedes traer tu montura aquí!!", FontTypeNames.FONTTYPE_INFO
3020                      Exit Sub
3030                  End If
                      
3040                  If .flags.Mimetizado Then
3050                      WriteConsoleMsg UserIndex, "¡¡No puedes utilizar tu montura estando transformado.!!", FontTypeNames.FONTTYPE_INFO
3060                      Exit Sub
3070                  End If
                      
3080                  If ((LegalPos(.Pos.map, .Pos.X, .Pos.Y, True, False) _
                              Or LegalPos(.Pos.map, .Pos.X, .Pos.Y, True, False) _
                              Or LegalPos(.Pos.map, .Pos.X, .Pos.Y, True, False) _
                              Or LegalPos(.Pos.map, .Pos.X, .Pos.Y, True, False)) _
                              And .flags.Navegando = 0) _
                              Or .flags.Navegando = 1 Then
3090                      Call WriteConsoleMsg(UserIndex, "¡No puedes montar en el agua!", FontTypeNames.FONTTYPE_INFO)
3100                  Else
3110                   Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARMONTURA, .Pos.X, .Pos.Y))
3120                  Call DoEquita(UserIndex, Obj, Slot)
3130                  End If
                      
3140              Case eOBJType.otMonturasDraco
                      
3150                  If .flags.Muerto = 1 Then
3160                      Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
3170                  Exit Sub
3180                  End If
                      
                                            If Not MapInfo(.Pos.map).Pk Then
                        WriteConsoleMsg UserIndex, "No puedes equipar montura aquí", FontTypeNames.FONTTYPE_INFO
                        Exit Sub
                      End If
                      
3190                  If .flags.SlotEvent > 0 Then
3200                      WriteConsoleMsg UserIndex, "¡¡No puedes traer tu montura aquí!!", FontTypeNames.FONTTYPE_INFO
3210                      Exit Sub
3220                  End If
                      
3230                  If .flags.Mimetizado Then
3240                      WriteConsoleMsg UserIndex, "¡¡No puedes utilizar tu montura estando transformado.!!", FontTypeNames.FONTTYPE_INFO
3250                      Exit Sub
3260                  End If
                      
3270                  If ((LegalPos(.Pos.map, .Pos.X, .Pos.Y, True, False) _
                              Or LegalPos(.Pos.map, .Pos.X, .Pos.Y, True, False) _
                              Or LegalPos(.Pos.map, .Pos.X, .Pos.Y, True, False) _
                              Or LegalPos(.Pos.map, .Pos.X, .Pos.Y, True, False)) _
                              And .flags.Navegando = 0) _
                              Or .flags.Navegando = 1 Then
3280                      Call WriteConsoleMsg(UserIndex, "¡No puedes montar en el agua!", FontTypeNames.FONTTYPE_INFO)
3290                  Else
3300                   Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARMONTURADRACO, .Pos.X, .Pos.Y))
3310                  Call DoEquita(UserIndex, Obj, Slot)
3320                  End If
                      
3330          End Select
          
3340      End With

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        Call WriteBlacksmithWeapons(UserIndex)
End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        Call WriteCarpenterObjects(UserIndex)
End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        Call WriteBlacksmithArmors(UserIndex)
End Sub

Sub TirarTodo(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error Resume Next

20        With UserList(UserIndex)
30            If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
              
40            Call TirarTodosLosItems(UserIndex)
              
50        End With

End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With ObjData(Index)
20            ItemSeCae = (.Real <> 1 Or .NoSeCae = 0) And _
                          (.Caos <> 1 Or .NoSeCae = 0) And _
                          .ObjType <> eOBJType.otLlaves And _
                          .ObjType <> eOBJType.otBarcos And _
                          .ObjType <> eOBJType.otMonturas And _
                          .ObjType <> eOBJType.otMonturasDraco And _
                          .NoSeCae = 0
30        End With

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
          
10        With UserList(UserIndex)
20            For i = 1 To .CurrentInventorySlots
30                ItemIndex = .Invent.Object(i).ObjIndex
40                If ItemIndex > 0 Then
50                     If ItemSeCae(ItemIndex) Then
60                        NuevaPos.X = 0
70                        NuevaPos.Y = 0
                          
                          'Creo el Obj
80                        MiObj.Amount = .Invent.Object(i).Amount
90                        MiObj.ObjIndex = ItemIndex

100                       DropAgua = True
                          ' Es pirata?
110                       If .clase = eClass.Pirat Then
                              ' Si tiene galeon equipado
120                           If .Invent.BarcoObjIndex = 476 Then
                                  ' Limitación por nivel, después dropea normalmente
130                               If .Stats.ELV >= 20 And .Stats.ELV <= 25 Then
                                      ' No dropea en agua
140                                   DropAgua = False
150                               End If
160                           End If
170                       End If
                          
180                       Call Tilelibre(.Pos, NuevaPos, MiObj, DropAgua, True)
                          
190                       If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
200                           Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
210                       End If
220                    End If
230               End If
240           Next i
250       End With
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
          
20        ItemNewbie = ObjData(ItemIndex).Newbie = 1
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
          
10        With UserList(UserIndex)
20            If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
              
30            For i = 1 To UserList(UserIndex).CurrentInventorySlots
40                ItemIndex = .Invent.Object(i).ObjIndex
50                If ItemIndex > 0 Then
60                    If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
70                        NuevaPos.X = 0
80                        NuevaPos.Y = 0
                          
                          'Creo MiObj
90                        MiObj.Amount = .Invent.Object(i).Amount
100                       MiObj.ObjIndex = ItemIndex
                          'Pablo (ToxicWaste) 24/01/2007
                          'Tira los Items no newbies en todos lados.
110                       Tilelibre .Pos, NuevaPos, MiObj, True, True
120                       If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
130                           Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
140                       End If
150                   End If
160               End If
170           Next i
180       End With

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
          
10        With UserList(UserIndex)
20            If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
              
30            For i = MAX_NORMAL_INVENTORY_SLOTS + 1 To .CurrentInventorySlots
40                ItemIndex = .Invent.Object(i).ObjIndex
50                If ItemIndex > 0 Then
60                    If ItemSeCae(ItemIndex) Then
70                        NuevaPos.X = 0
80                        NuevaPos.Y = 0
                          
                          'Creo MiObj
90                        MiObj.Amount = .Invent.Object(i).Amount
100                       MiObj.ObjIndex = ItemIndex
110                       Tilelibre .Pos, NuevaPos, MiObj, True, True
120                       If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
130                           Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
140                       End If
150                   End If
160               End If
170           Next i
180       End With

End Sub

Public Function getObjType(ByVal ObjIndex As Integer) As eOBJType
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If ObjIndex > 0 Then
20            getObjType = ObjData(ObjIndex).ObjType
30        End If
          
End Function
Function ItemFaccionario(ByVal ItemIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
          
20        ItemFaccionario = ObjData(ItemIndex).Caos Or ObjData(ItemIndex).Real = 1
End Function
Function ItemVIP(ByVal ItemIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function

20         ItemVIP = ObjData(ItemIndex).VIP = 1
End Function
Function ItemVIPB(ByVal ItemIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function

20         ItemVIPB = ObjData(ItemIndex).VIPB = 1
End Function
Function ItemVIPP(ByVal ItemIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function

20         ItemVIPP = ObjData(ItemIndex).VIPP = 1
End Function
Function ItemQuince(ByVal ItemIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
          
20        ItemQuince = ObjData(ItemIndex).Quince = 1
End Function
Function ItemTreinta(ByVal ItemIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
          
20        ItemTreinta = ObjData(ItemIndex).Treinta = 1
End Function
Function ItemHM(ByVal ItemIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
          
20        ItemHM = ObjData(ItemIndex).HM = 1
End Function
Function ItemUM(ByVal ItemIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
          
20        ItemUM = ObjData(ItemIndex).UM = 1
End Function
Public Sub moveItem(ByVal UserIndex As Integer, ByVal originalSlot As Integer, ByVal NewSlot As Integer)
       
      Dim tmpObj As UserOBJ
      Dim newObjIndex As Integer, originalObjIndex As Integer
   On Error GoTo moveItem_Error

10    If (originalSlot <= 0) Or (NewSlot <= 0) Then Exit Sub
       
20    With UserList(UserIndex)
30        If (originalSlot > .CurrentInventorySlots) Or (NewSlot > .CurrentInventorySlots) Then Exit Sub
         
40        tmpObj = .Invent.Object(originalSlot)
50        .Invent.Object(originalSlot) = .Invent.Object(NewSlot)
60        .Invent.Object(NewSlot) = tmpObj
         
          'Viva VB6 y sus putas deficiencias.
70        If .Invent.AnilloEqpSlot = originalSlot Then
80            .Invent.AnilloEqpSlot = NewSlot
90        ElseIf .Invent.AnilloEqpSlot = NewSlot Then
100           .Invent.AnilloEqpSlot = originalSlot
110       End If
         
120       If .Invent.ArmourEqpSlot = originalSlot Then
130           .Invent.ArmourEqpSlot = NewSlot
140       ElseIf .Invent.ArmourEqpSlot = NewSlot Then
150           .Invent.ArmourEqpSlot = originalSlot
160       End If
         
170       If .Invent.BarcoSlot = originalSlot Then
180           .Invent.BarcoSlot = NewSlot
190       ElseIf .Invent.BarcoSlot = NewSlot Then
200           .Invent.BarcoSlot = originalSlot
210       End If
         
220       If .Invent.CascoEqpSlot = originalSlot Then
230            .Invent.CascoEqpSlot = NewSlot
240       ElseIf .Invent.CascoEqpSlot = NewSlot Then
250            .Invent.CascoEqpSlot = originalSlot
260       End If
         
270       If .Invent.EscudoEqpSlot = originalSlot Then
280           .Invent.EscudoEqpSlot = NewSlot
290       ElseIf .Invent.EscudoEqpSlot = NewSlot Then
300           .Invent.EscudoEqpSlot = originalSlot
310       End If
         
320       If .Invent.MochilaEqpSlot = originalSlot Then
330           .Invent.MochilaEqpSlot = NewSlot
340       ElseIf .Invent.MochilaEqpSlot = NewSlot Then
350           .Invent.MochilaEqpSlot = originalSlot
360       End If
         
370       If .Invent.MunicionEqpSlot = originalSlot Then
380           .Invent.MunicionEqpSlot = NewSlot
390       ElseIf .Invent.MunicionEqpSlot = NewSlot Then
400           .Invent.MunicionEqpSlot = originalSlot
410       End If
         
420       If .Invent.WeaponEqpSlot = originalSlot Then
430           .Invent.WeaponEqpSlot = NewSlot
440       ElseIf .Invent.WeaponEqpSlot = NewSlot Then
450           .Invent.WeaponEqpSlot = originalSlot
460       End If
       
470       Call UpdateUserInv(False, UserIndex, originalSlot)
480       Call UpdateUserInv(False, UserIndex, NewSlot)
490   End With

   On Error GoTo 0
   Exit Sub

moveItem_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure moveItem of Módulo InvUsuario in line " & Erl
End Sub
 
