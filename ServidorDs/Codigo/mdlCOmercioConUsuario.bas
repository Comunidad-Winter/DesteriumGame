Attribute VB_Name = "mdlCOmercioConUsuario"
'**************************************************************
' mdlComercioConUsuarios.bas - Allows players to commerce between themselves.
'
' Designed and implemented by Alejandro Santos (AlejoLP)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

'[Alejo]
Option Explicit

Private Const MAX_ORO_LOGUEABLE As Long = 50000
Private Const MAX_OBJ_LOGUEABLE As Long = 1000

Public Const MAX_OFFER_SLOTS As Integer = 20
Public Const GOLD_OFFER_SLOT As Integer = MAX_OFFER_SLOTS + 1

Public Type tCOmercioUsuario
    DestUsu As Integer 'El otro Usuario
    DestNick As String
    Objeto(1 To MAX_OFFER_SLOTS) As Integer 'Indice de los objetos que se desea dar
    GoldAmount As Long
    
    cant(1 To MAX_OFFER_SLOTS) As Long 'Cuantos objetos desea dar
    Acepto As Boolean
    Confirmo As Boolean
End Type

Private Type tOfferItem
    objindex As Integer
    Amount As Long
End Type

'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Sub IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer)
      '***************************************************
      'Autor: Unkown
      'Last Modification: 25/11/2009
      '
      '***************************************************
10        On Error GoTo Errhandler
          
          'Si ambos pusieron /comerciar entonces
20        If UserList(Origen).ComUsu.DestUsu = Destino And _
             UserList(Destino).ComUsu.DestUsu = Origen Then
             
30           If UserList(Origen).flags.Comerciando Or UserList(Destino).flags.Comerciando Then
40                Call WriteConsoleMsg(Origen, "No puedes comerciar en este momento", FontTypeNames.FONTTYPE_TALK)
50                Call WriteConsoleMsg(Destino, "No puedes comerciar en este momento", FontTypeNames.FONTTYPE_TALK)
60                Exit Sub
70            End If
              
              'Actualiza el inventario del usuario
80            Call UpdateUserInv(True, Origen, 0)
              'Decirle al origen que abra la ventanita.
90            Call WriteUserCommerceInit(Origen)
100           UserList(Origen).flags.Comerciando = True
          
              'Actualiza el inventario del usuario
110           Call UpdateUserInv(True, Destino, 0)
              'Decirle al origen que abra la ventanita.
120           Call WriteUserCommerceInit(Destino)
130           UserList(Destino).flags.Comerciando = True
          
              'Call EnviarObjetoTransaccion(Origen)
140       Else
              'Es el primero que comercia ?
150           Call WriteConsoleMsg(Destino, UserList(Origen).Name & " desea comerciar. Si deseas aceptar, escribe /COMERCIAR.", FontTypeNames.FONTTYPE_TALK)
160           UserList(Destino).flags.TargetUser = Origen
              
170       End If
          
180       Call FlushBuffer(Destino)
          
190       Exit Sub
Errhandler:
200           Call LogError("Error en IniciarComercioConUsuario: " & Err.Description)
End Sub

Public Sub EnviarOferta(ByVal Userindex As Integer, ByVal OfferSlot As Byte)
      '***************************************************
      'Autor: Unkown
      'Last Modification: 25/11/2009
      'Sends the offer change to the other trading user
      '25/11/2009: ZaMa - Implementado nuevo sistema de comercio con ofertas variables.
      '***************************************************
          Dim objindex As Integer
          Dim ObjAmount As Long
          Dim OtherUserIndex As Integer
          
          
10        OtherUserIndex = UserList(Userindex).ComUsu.DestUsu
          
20        With UserList(OtherUserIndex)
30            If OfferSlot = GOLD_OFFER_SLOT Then
40                objindex = iORO
50                ObjAmount = .ComUsu.GoldAmount
60            Else
70                objindex = .ComUsu.Objeto(OfferSlot)
80                ObjAmount = .ComUsu.cant(OfferSlot)
90            End If
100       End With
         
110       Call WriteChangeUserTradeSlot(Userindex, OfferSlot, objindex, ObjAmount)
120       Call FlushBuffer(Userindex)

End Sub

Public Sub FinComerciarUsu(ByVal Userindex As Integer)
      '***************************************************
      'Autor: Unkown
      'Last Modification: 25/11/2009
      '25/11/2009: ZaMa - Limpio los arrays (por el nuevo sistema)
      '***************************************************
          Dim i As Long
          
10        With UserList(Userindex)
20            If .ComUsu.DestUsu > 0 Then
30                Call WriteUserCommerceEnd(Userindex)
40            End If
              
50            .ComUsu.Acepto = False
60            .ComUsu.Confirmo = False
70            .ComUsu.DestUsu = 0
              
80            For i = 1 To MAX_OFFER_SLOTS
90                .ComUsu.cant(i) = 0
100               .ComUsu.Objeto(i) = 0
110           Next i
              
120           .ComUsu.GoldAmount = 0
130           .ComUsu.DestNick = vbNullString
140           .flags.Comerciando = False
150       End With
End Sub

Public Sub AceptarComercioUsu(ByVal Userindex As Integer)
      '***************************************************
      'Autor: Unkown
      'Last Modification: 06/05/2010
      '25/11/2009: ZaMa - Ahora se traspasan hasta 5 items + oro al comerciar
      '06/05/2010: ZaMa - Ahora valida si los usuarios tienen los items que ofertan.
      '***************************************************
          Dim TradingObj As Obj
          Dim OtroUserIndex As Integer
          Dim OfferSlot As Integer

10        UserList(Userindex).ComUsu.Acepto = True
          
20        OtroUserIndex = UserList(Userindex).ComUsu.DestUsu
          
          ' Acepto el otro?
30        If UserList(OtroUserIndex).ComUsu.Acepto = False Then
40            Exit Sub
50        End If
          
          ' User valido?
60        If OtroUserIndex <= 0 Or OtroUserIndex > MaxUsers Then
70            Call FinComerciarUsu(Userindex)
80            Exit Sub
90        End If
          
          
          ' Aceptaron ambos, chequeo que tengan los items que ofertaron
100       If Not HasOfferedItems(Userindex) Then
              
110           Call WriteConsoleMsg(Userindex, "¡¡¡El comercio se canceló porque no posees los ítems que ofertaste!!!", FontTypeNames.FONTTYPE_FIGHT)
120           Call WriteConsoleMsg(OtroUserIndex, "¡¡¡El comercio se canceló porque " & UserList(Userindex).Name & " no posee los ítems que ofertó!!!", FontTypeNames.FONTTYPE_FIGHT)
              
130           Call FinComerciarUsu(Userindex)
140           Call FinComerciarUsu(OtroUserIndex)
150           Call Protocol.FlushBuffer(OtroUserIndex)
              
160           Exit Sub
              
170       ElseIf Not HasOfferedItems(OtroUserIndex) Then
              
180           Call WriteConsoleMsg(Userindex, "¡¡¡El comercio se canceló porque " & UserList(OtroUserIndex).Name & " no posee los ítems que ofertó!!!", FontTypeNames.FONTTYPE_FIGHT)
190           Call WriteConsoleMsg(OtroUserIndex, "¡¡¡El comercio se canceló porque no posees los ítems que ofertaste!!!", FontTypeNames.FONTTYPE_FIGHT)
              
200           Call FinComerciarUsu(Userindex)
210           Call FinComerciarUsu(OtroUserIndex)
220           Call Protocol.FlushBuffer(OtroUserIndex)
              
230           Exit Sub
              
240       End If
          
          ' Envio los items a quien corresponde
250       For OfferSlot = 1 To MAX_OFFER_SLOTS + 1
              
              ' Items del 1er usuario
260           With UserList(Userindex)
                  ' Le pasa el oro
270               If OfferSlot = GOLD_OFFER_SLOT Then
                      ' Quito la cantidad de oro ofrecida
280                   .Stats.Gld = .Stats.Gld - .ComUsu.GoldAmount
                      ' Log
290                   If .ComUsu.GoldAmount > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(.Name & " soltó oro en comercio seguro con " & UserList(OtroUserIndex).Name & ". Cantidad: " & .ComUsu.GoldAmount)
                      ' Update Usuario
300                   Call WriteUpdateUserStats(Userindex)
                      ' Se la doy al otro
310                   UserList(OtroUserIndex).Stats.Gld = UserList(OtroUserIndex).Stats.Gld + .ComUsu.GoldAmount
                      ' Update Otro Usuario
320                   Call WriteUpdateUserStats(OtroUserIndex)
                      
                  ' Le pasa lo ofertado de los slots con items
330               ElseIf .ComUsu.Objeto(OfferSlot) > 0 Then
340                   TradingObj.objindex = .ComUsu.Objeto(OfferSlot)
350                   TradingObj.Amount = .ComUsu.cant(OfferSlot)
                                      
                      'Quita el objeto y se lo da al otro
360                   If Not MeterItemEnInventario(OtroUserIndex, TradingObj) Then
370                       Call TirarItemAlPiso(UserList(OtroUserIndex).Pos, TradingObj)
380                   End If
                  
390                   Call QuitarObjetos(TradingObj.objindex, TradingObj.Amount, Userindex)
                      
                      'Es un Objeto que tenemos que loguear? Pablo (ToxicWaste) 07/09/07
400                   If ObjData(TradingObj.objindex).LOG = 1 Then
410                       Call LogDesarrollo(.Name & " le pasó en comercio seguro a " & UserList(OtroUserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.objindex).Name)
420                   End If
                  
                      'Es mucha cantidad?
430                   If TradingObj.Amount > MAX_OBJ_LOGUEABLE Then
                      'Si no es de los prohibidos de loguear, lo logueamos.
440                       If ObjData(TradingObj.objindex).NoLog <> 1 Then
450                           Call LogDesarrollo(UserList(OtroUserIndex).Name & " le pasó en comercio seguro a " & .Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.objindex).Name)
460                       End If
470                   End If
480               End If
490           End With
              
              ' Items del 2do usuario
500           With UserList(OtroUserIndex)
                  ' Le pasa el oro
510               If OfferSlot = GOLD_OFFER_SLOT Then
                      ' Quito la cantidad de oro ofrecida
520                   .Stats.Gld = .Stats.Gld - .ComUsu.GoldAmount
                      ' Log
530                   If .ComUsu.GoldAmount > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(.Name & " soltó oro en comercio seguro con " & UserList(Userindex).Name & ". Cantidad: " & .ComUsu.GoldAmount)
                      ' Update Usuario
540                   Call WriteUpdateUserStats(OtroUserIndex)
                      'y se la doy al otro
550                   UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld + .ComUsu.GoldAmount
560                   If .ComUsu.GoldAmount > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(UserList(Userindex).Name & " recibió oro en comercio seguro con " & .Name & ". Cantidad: " & .ComUsu.GoldAmount)
                      ' Update Otro Usuario
570                   Call WriteUpdateUserStats(Userindex)
                      
                  ' Le pasa la oferta de los slots con items
580               ElseIf .ComUsu.Objeto(OfferSlot) > 0 Then
590                   TradingObj.objindex = .ComUsu.Objeto(OfferSlot)
600                   TradingObj.Amount = .ComUsu.cant(OfferSlot)
                                      
                      'Quita el objeto y se lo da al otro
610                   If Not MeterItemEnInventario(Userindex, TradingObj) Then
620                       Call TirarItemAlPiso(UserList(Userindex).Pos, TradingObj)
630                   End If
                  
640                   Call QuitarObjetos(TradingObj.objindex, TradingObj.Amount, OtroUserIndex)
                      
                      'Es un Objeto que tenemos que loguear? Pablo (ToxicWaste) 07/09/07
650                   If ObjData(TradingObj.objindex).LOG = 1 Then
660                       Call LogDesarrollo(.Name & " le pasó en comercio seguro a " & UserList(Userindex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.objindex).Name)
670                   End If
                  
                      'Es mucha cantidad?
680                   If TradingObj.Amount > MAX_OBJ_LOGUEABLE Then
                      'Si no es de los prohibidos de loguear, lo logueamos.
690                       If ObjData(TradingObj.objindex).NoLog <> 1 Then
700                           Call LogDesarrollo(.Name & " le pasó en comercio seguro a " & UserList(Userindex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.objindex).Name)
710                       End If
720                   End If
730               End If
740           End With
              
750       Next OfferSlot

          ' End Trade
760       Call FinComerciarUsu(Userindex)
770       Call FinComerciarUsu(OtroUserIndex)
780       Call Protocol.FlushBuffer(OtroUserIndex)
          
End Sub

Public Sub AgregarOferta(ByVal Userindex As Integer, ByVal OfferSlot As Byte, ByVal objindex As Integer, ByVal Amount As Long, ByVal IsGold As Boolean)
      '***************************************************
      'Autor: ZaMa
      'Last Modification: 24/11/2009
      'Adds gold or items to the user's offer
      '***************************************************

10        If PuedeSeguirComerciando(Userindex) Then
20            With UserList(Userindex).ComUsu
                  ' Si ya confirmo su oferta, no puede cambiarla!
30                If Not .Confirmo Then
40                    If IsGold Then
                      ' Agregamos (o quitamos) mas oro a la oferta
50                        .GoldAmount = .GoldAmount + Amount
                          
                          ' Imposible que pase, pero por las dudas..
60                        If .GoldAmount < 0 Then .GoldAmount = 0
70                    Else
                      ' Agreamos (o quitamos) el item y su cantidad en el slot correspondiente
                          ' Si es 0 estoy modificando la cantidad, no agregando
80                        If objindex > 0 Then .Objeto(OfferSlot) = objindex
90                        .cant(OfferSlot) = .cant(OfferSlot) + Amount
                          
                          'Quitó todos los items de ese tipo
100                       If .cant(OfferSlot) <= 0 Then
                              ' Removemos el objeto para evitar conflictos
110                           .Objeto(OfferSlot) = 0
120                           .cant(OfferSlot) = 0
130                       End If
140                   End If
150               End If
160           End With
170       End If

End Sub

Public Function PuedeSeguirComerciando(ByVal Userindex As Integer) As Boolean
      '***************************************************
      'Autor: ZaMa
      'Last Modification: 24/11/2009
      'Validates wether the conditions for the commerce to keep going are satisfied
      '***************************************************
      Dim OtroUserIndex As Integer
      Dim ComercioInvalido As Boolean

10    With UserList(Userindex)
          ' Usuario valido?
20        If .ComUsu.DestUsu <= 0 Or .ComUsu.DestUsu > MaxUsers Then
30            ComercioInvalido = True
40        End If
          
50        OtroUserIndex = .ComUsu.DestUsu
          
60        If Not ComercioInvalido Then
              ' Estan logueados?
70            If UserList(OtroUserIndex).flags.UserLogged = False Or .flags.UserLogged = False Then
80                ComercioInvalido = True
90            End If
100       End If
          
110       If Not ComercioInvalido Then
              ' Se estan comerciando el uno al otro?
120           If UserList(OtroUserIndex).ComUsu.DestUsu <> Userindex Then
130               ComercioInvalido = True
140           End If
150       End If
          
160       If Not ComercioInvalido Then
              ' El nombre del otro es el mismo que al que le comercio?
170           If UserList(OtroUserIndex).Name <> .ComUsu.DestNick Then
180               ComercioInvalido = True
190           End If
200       End If
          
210       If Not ComercioInvalido Then
              ' Mi nombre  es el mismo que al que el le comercia?
220           If .Name <> UserList(OtroUserIndex).ComUsu.DestNick Then
230               ComercioInvalido = True
240           End If
250       End If
          
260       If Not ComercioInvalido Then
              ' Esta vivo?
270           If UserList(OtroUserIndex).flags.Muerto = 1 Then
280               ComercioInvalido = True
290           End If
300       End If
          
          ' Fin del comercio
310       If ComercioInvalido = True Then
320           Call FinComerciarUsu(Userindex)
              
330           If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
340               Call FinComerciarUsu(OtroUserIndex)
350               Call Protocol.FlushBuffer(OtroUserIndex)
360           End If
              
370           Exit Function
380       End If
390   End With

400   PuedeSeguirComerciando = True

End Function

Private Function HasOfferedItems(ByVal Userindex As Integer) As Boolean
      '***************************************************
      'Autor: ZaMa
      'Last Modification: 05/06/2010
      'Checks whether the user has the offered items in his inventory or not.
      '***************************************************

          Dim OfferedItems(MAX_OFFER_SLOTS - 1) As tOfferItem
          Dim Slot As Long
          Dim SlotAux As Long
          Dim SlotCount As Long
          
          Dim objindex As Integer
          
10        With UserList(Userindex).ComUsu
              
              ' Agrupo los items que son iguales
20            For Slot = 1 To MAX_OFFER_SLOTS
                          
30                objindex = .Objeto(Slot)
                  
40                If objindex > 0 Then
                  
50                    For SlotAux = 0 To SlotCount - 1
                          
60                        If objindex = OfferedItems(SlotAux).objindex Then
                              ' Son iguales, aumento la cantidad
70                            OfferedItems(SlotAux).Amount = OfferedItems(SlotAux).Amount + .cant(Slot)
80                            Exit For
90                        End If
                          
100                   Next SlotAux
                      
                      ' No encontro otro igual, lo agrego
110                   If SlotAux = SlotCount Then
120                       OfferedItems(SlotCount).objindex = objindex
130                       OfferedItems(SlotCount).Amount = .cant(Slot)
                          
140                       SlotCount = SlotCount + 1
150                   End If
                      
160               End If
                  
170           Next Slot
              
              ' Chequeo que tengan la cantidad en el inventario
180           For Slot = 0 To SlotCount - 1
190               If Not HasEnoughItems(Userindex, OfferedItems(Slot).objindex, OfferedItems(Slot).Amount) Then Exit Function
200           Next Slot
              
              ' Compruebo que tenga el oro que oferta
210           If UserList(Userindex).Stats.Gld < .GoldAmount Then Exit Function
              
220       End With
          
230       HasOfferedItems = True

End Function


