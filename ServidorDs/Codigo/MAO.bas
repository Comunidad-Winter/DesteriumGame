Attribute VB_Name = "MAO"
Option Explicit

Private Const MAX_MAO As Integer = 500
Public Const MAX_OFFER As Integer = 20

Public Type tUserMercado
    InList As Integer
    Dsp As Long
    Gld As Long
    Change As Byte
    dstName As String
    
    lstOffer(1 To MAX_OFFER) As String
    lstOfferSent(1 To MAX_OFFER) As String
    
End Type

Private lstMercado(1 To MAX_MAO) As String


Private Sub CreateMAO()
          Dim intFile As Integer
          Dim i As Integer
          
10        intFile = FreeFile

20        Open App.Path & "\DAT\MAO.DAT" For Output As #intFile
30        Print #intFile, "[INIT]"

40        For i = 1 To 500
50            Print #intFile, "Pj" & i & "=0"
60        Next i
70        Close #intFile
End Sub

Public Sub Add_Change(ByVal UserIndex As Integer, _
                        ByVal Email As String, _
                        ByVal Passwd As String, _
                        ByVal Pin As String)
                              
          Dim ErrorMsg As String
          Dim Slot As Integer
          
   On Error GoTo Add_Change_Error

10        If Check_Publication(UserIndex, Email, Passwd, Pin, ErrorMsg) Then
20            If Add_Mercado(UserIndex, UserList(UserIndex).Name, Slot) Then
30                With UserList(UserIndex).Mercado
40                    .InList = Slot
50                    .Gld = 0
60                    .Dsp = 0
70                    .Change = 1
80                End With
                  
90                SaveMercadoUser UserIndex
100               WriteConsoleMsg UserIndex, "Personaje " & UserList(UserIndex).Name & " agregado exitosamente. Puedes verlo en la lista de personajes POR CAMBIO", FontTypeNames.FONTTYPE_INFO
110           End If
120       Else
130           WriteConsoleMsg UserIndex, ErrorMsg, FontTypeNames.FONTTYPE_PARTY
140       End If

   On Error GoTo 0
   Exit Sub

Add_Change_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Add_Change of Módulo MAO in line " & Erl
End Sub

Public Sub Add_Gld_Dsp(ByVal UserIndex As Integer, _
                        ByVal Email As String, _
                        ByVal Passwd As String, _
                        ByVal Pin As String, _
                        ByVal dstName As String, _
                        ByVal Gld As Long, _
                        ByVal Dsp As Long)
          
          Dim ErrorMsg As String
          Dim Slot As Integer
          
   On Error GoTo Add_Gld_Dsp_Error

10        If (dstName = vbNullString) Then Exit Sub
          
20        If Check_Publication(UserIndex, Email, Passwd, Pin, ErrorMsg, dstName, Gld, Dsp) Then
30            If Add_Mercado(UserIndex, UserList(UserIndex).Name, Slot) Then
              
40                With UserList(UserIndex).Mercado
50                    .Change = 0
60                    .Dsp = Dsp
70                    .Gld = Gld
80                    .dstName = dstName
90                    .InList = Slot
100               End With
                  
110               SaveMercadoUser UserIndex
120               WriteConsoleMsg UserIndex, "Personaje " & UserList(UserIndex).Name & " agregado exitosamente. Puede verlo ahora mismo en la lista de personajes en VENTA", FontTypeNames.FONTTYPE_INFO
130           End If
140       Else
150           WriteConsoleMsg UserIndex, ErrorMsg, FontTypeNames.FONTTYPE_INFO
              
160       End If

   On Error GoTo 0
   Exit Sub

Add_Gld_Dsp_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Add_Gld_Dsp of Módulo MAO in line " & Erl
End Sub
Public Sub Buy_Pj(ByVal UserIndex As Integer, _
                    ByVal dstName As String)
          
          Dim ErrorMsg As String
          Dim Gld As Long
          Dim Dsp As Long
          Dim Account As String
          
   On Error GoTo Buy_Pj_Error

10        With UserList(UserIndex)
20            If Check_Buy_Pj(UserIndex, dstName, ErrorMsg, Gld, Dsp) Then
                  Account = GetVar(CharPath & dstName & ".chr", "INIT", "ACCOUNT")
                  
                    If Account = vbNullString Then
                        WriteConsoleMsg UserIndex, "El personaje no está habilitado para comprar. Contacte al dueño y pidele que lo vuelva a postear desde su cuenta", FontTypeNames.FONTTYPE_INFO
                        Exit Sub
                   End If
                   
30                WriteConsoleMsg UserIndex, "Has comprado el personaje " & dstName & ".", FontTypeNames.FONTTYPE_INFO
                  
40                Call QuitarObjetos(880, Dsp, UserIndex)
50                .Stats.Gld = .Stats.Gld - Gld
60                WriteUpdateGold UserIndex
                  
70                Call Add_Gld_Dsp_Vendedor(UserIndex, dstName)
80                Call CopyData(UserIndex, dstName)
90                Call Remove_Mercado(dstName)

                
                  UpdateAccountUserName dstName, Account
                  mCuenta.AddCharAccount .Account, dstName
                  
                  LogMao "El personaje " & .Name & " compró el personaje " & dstName & "Cuenta destino: " & .Account
100           Else
110               WriteConsoleMsg UserIndex, ErrorMsg, FontTypeNames.FONTTYPE_WARNING
120           End If
          
          
130       End With

   On Error GoTo 0
   Exit Sub

Buy_Pj_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Buy_Pj of Módulo MAO in line " & Erl
End Sub

Private Sub Add_Gld_Dsp_Vendedor(ByVal UserIndex As Integer, ByVal dstName As String)
          Dim LoopC As Integer
          Dim Gld As Long
          Dim Dsp As Long
          Dim Gld_Bank As Long
          Dim NamePaga As String
   On Error GoTo Add_Gld_Dsp_Vendedor_Error

10        Dim exito As Boolean: exito = False
          Dim PagaIndex As Integer
          
20        NamePaga = GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "dstName")
30        Gld = val(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "GLD"))
40        Dsp = val(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "DSP"))
              
50        If FileExist(CharPath & UCase$(NamePaga) & ".chr") Then
60            PagaIndex = NameIndex(NamePaga)
              
70            If Gld > 0 Then
80                If PagaIndex > 0 Then
90                    Gld_Bank = UserList(PagaIndex).Stats.Banco
100                   UserList(PagaIndex).Stats.Banco = UserList(PagaIndex).Stats.Banco + Gld
110                   exito = True
120               Else
130                   Gld_Bank = val(GetVar(CharPath & UCase$(NamePaga) & ".chr", "STATS", "BANCO"))
140                   exito = True
                      
150               End If
                  
160               WriteVar CharPath & UCase$(NamePaga) & ".chr", "STATS", "BANCO", Gld_Bank + Gld
170           End If
              
180           If Dsp > 0 Then
190               For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
200                   If PagaIndex > 0 Then
210                       With UserList(PagaIndex)
220                           If .BancoInvent.Object(LoopC).ObjIndex = 0 Then
230                               .BancoInvent.Object(LoopC).ObjIndex = 880
240                               .BancoInvent.Object(LoopC).Amount = Dsp
250                               exito = True
260                               Exit For
270                           End If
280                       End With
290                   Else
300                       If GetVar(CharPath & UCase$(NamePaga) & ".chr", "BANCOINVENTORY", "OBJ" & LoopC) = "0-0" Then
310                           WriteVar CharPath & UCase$(NamePaga) & ".chr", "BANCOINVENTORY", "OBJ" & LoopC, "880-" & Dsp
320                           exito = True
330                           Exit For
340                       End If
                      
350                   End If
360               Next LoopC
370           End If
              
380           If exito = False Then
390               MAO.LogMao "El personaje " & dstName & " comprado por " & UserList(UserIndex).Name & " no se concretó (depositante " & NamePaga & ") de ORO: " & Gld & " y DSP: " & Dsp
400           Else
410               MAO.LogMao "El personaje " & UserList(UserIndex).Name & " ha comprado al personaje " & dstName & " a cambio de " & Gld & " monedas de oro y " & Dsp & " monedas DSP. Depositante: " & NamePaga

End If
430       Else
440           MAO.LogMao "El personaje " & dstName & " no pudo ser comprado ya que el depositante se encuentra inexistente. La paga se realiza en " & NamePaga & ""
450       End If
          

   On Error GoTo 0
   Exit Sub

Add_Gld_Dsp_Vendedor_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Add_Gld_Dsp_Vendedor of Módulo MAO in line " & Erl
          
End Sub
Private Sub SaveMercadoUser(ByVal UserIndex As Integer)
   On Error GoTo SaveMercadoUser_Error

10        With UserList(UserIndex).Mercado
20            WriteVar CharPath & UCase$(UserList(UserIndex).Name) & ".chr", "MERCADO", "InList", .InList
30            WriteVar CharPath & UCase$(UserList(UserIndex).Name) & ".chr", "MERCADO", "Change", .Change
40            WriteVar CharPath & UCase$(UserList(UserIndex).Name) & ".chr", "MERCADO", "Dsp", .Dsp
50            WriteVar CharPath & UCase$(UserList(UserIndex).Name) & ".chr", "MERCADO", "dstName", .dstName
60            WriteVar CharPath & UCase$(UserList(UserIndex).Name) & ".chr", "MERCADO", "GLD", .Gld
70        End With

   On Error GoTo 0
   Exit Sub

SaveMercadoUser_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveMercadoUser of Módulo MAO in line " & Erl
End Sub
Private Sub Remove_Mercado(ByVal dstName As String)
          Dim InList As Integer
          Dim LoopC As Integer
          Dim NamePaga As String
          Dim UserIndex As Integer
          
   On Error GoTo Remove_Mercado_Error

10        InList = val(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "InList"))
20        NamePaga = GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "dstName")
          

          
         ' If StrComp(UCase$(dstName), lstMercado(InList)) = 0 Then
30            UserIndex = NameIndex(dstName)
40            If UserIndex > 0 Then ResetMercado UserIndex
              
50            If InList > 0 Then
60                lstMercado(InList) = "0"
                  
70                WriteVar App.Path & "\DAT\MAO.DAT", "INIT", "Pj" & InList, "0"
80            End If
              
90            WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "InList", "0"
100           WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "Change", "0"
110           WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "Dsp", "0"
120           WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "Gld", "0"
130           WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "dstName", "0"
             
140           For LoopC = 1 To MAX_OFFER
150               WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "Offer" & LoopC, "0"
160               WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "OfferSent" & LoopC, "0"
170           Next LoopC
          'End If

   On Error GoTo 0
   Exit Sub

Remove_Mercado_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Remove_Mercado of Módulo MAO in line " & Erl
End Sub
Private Sub CopyData(ByVal UserIndex As Integer, ByVal NewPj As String)
          
          Dim Passwd As String
          
   On Error GoTo CopyData_Error

10        With UserList(UserIndex)
20            Passwd = GetVar(CharPath & UCase$(.Name) & ".chr", "INIT", "PASSWORD")
          
30            WriteVar CharPath & NewPj & ".chr", "CONTACTO", "EMAIL", .Email
40            WriteVar CharPath & NewPj & ".chr", "INIT", "Pin", .Pin
50            WriteVar CharPath & NewPj & ".chr", "INIT", "PASSWORD", Passwd
60        End With

   On Error GoTo 0
   Exit Sub

CopyData_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure CopyData of Módulo MAO in line " & Erl
End Sub

Public Sub Remove_Pj(ByVal UserIndex As Integer)
   On Error GoTo Remove_Pj_Error

10        With UserList(UserIndex)
20            If Not .Mercado.InList > 0 Then
30                WriteConsoleMsg UserIndex, "Tu personaje no está en el mercado de DESTERIUM AO.", FontTypeNames.FONTTYPE_INFO
40                Exit Sub
50            End If
              
60            Remove_Mercado .Name
70            WriteConsoleMsg UserIndex, "Has quitado tu personaje del mercado. Se han borrado todos los datos relacionados al mercado, tanto ofertas como demandas.", FontTypeNames.FONTTYPE_INFO
80        End With

   On Error GoTo 0
   Exit Sub

Remove_Pj_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Remove_Pj of Módulo MAO in line " & Erl
End Sub
Private Function Check_Buy_Pj(ByVal UserIndex As Integer, _
                            ByVal dstName As String, _
                            ByRef ErrorMsg As String, _
                            ByRef Gld As Long, _
                            ByRef Dsp As Long) As Boolean
                                  
   On Error GoTo Check_Buy_Pj_Error

10        Check_Buy_Pj = False
          
20        With UserList(UserIndex)
30            If Not FileExist(CharPath & UCase$(dstName) & ".chr") Then
40                ErrorMsg = "El personaje " & dstName & " no existe."
50                'MAO.LogMao "El personaje " & .Name & " ha intentado comprar un PERSONAJE INEXISTENTE con NICK: " & dstName
60                Exit Function
70            End If
              
80            If StrComp(UCase$(.Name), UCase$(dstName)) = 0 Then
90                ErrorMsg = "No puedes realizar esta opción. Si lo vuelves a hacer se sospechará de tu CLIENTE."
100               'MAO.LogMao "El personaje " & .Name & " ha intentado comprarse A SI MISMO"
110               Exit Function
120           End If
              
130           If NameIndex(dstName) > 0 Then
140               ErrorMsg = "El personaje " & dstName & " se encuentra online. Se le avisará que lo quisiste COMPRAR"
150               WriteConsoleMsg NameIndex(dstName), "El personaje " & .Name & " ha intentado comprarte y no ha podido porque estás conectado.", FontTypeNames.FONTTYPE_INFO
160               Exit Function
170           End If
              
180           If Not CBool(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "InList")) Then
190               ErrorMsg = "El personaje " & dstName & " no está en la lista del MERCADO o bien ya lo compraron. Revisa nuevamente la lista."
200               Exit Function
210           End If
              
220           If CBool(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "Change")) Then
230               ErrorMsg = "El personaje " & dstName & " está en MODO CAMBIO"
240               MAO.LogMao "El personaje " & .Name & " ha intentado comprar un personaje en MODO CAMBIO con NICK: " & dstName
250               Exit Function
260           End If
              
              
270           Gld = val(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "Gld"))
                      
280           If .Stats.Gld < Gld Then
290               ErrorMsg = "Para comprar este personaje debes tener en tu billetera " & Gld & " monedas de oro."
300              Exit Function
310           End If
              
320           Dsp = val(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "Dsp"))
              
330           If Dsp > 0 Then
340               If Not TieneObjetos(880, Dsp, UserIndex) Then
350                   ErrorMsg = "Para comprar este personaje debes tener en tu inventario " & Dsp & " DSP"
360                   Exit Function
370               End If
380           End If

          
390       End With
          
400       Check_Buy_Pj = True

   On Error GoTo 0
   Exit Function

Check_Buy_Pj_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_Buy_Pj of Módulo MAO in line " & Erl
End Function
Public Sub Send_Invitation(ByVal UserIndex As Integer, _
                            ByVal dstName As String)
                                  
          Dim ErrorMsg As String
          Dim SlotOffer As Integer
          Dim tUser As Integer
          
   On Error GoTo Send_Invitation_Error

10        If Check_Invitation(UserIndex, dstName, ErrorMsg, SlotOffer) Then
20            If FreeSlotOffer(UserIndex, dstName) Then
                  ' Agregamos a nuestro charfile la oferta realizada
30                UserList(UserIndex).Mercado.lstOfferSent(SlotOffer) = UCase$(dstName)
40                WriteVar CharPath & UCase$(UserList(UserIndex).Name) & ".chr", "MERCADO", "OfferSent" & SlotOffer, UCase$(dstName)
                  
50                WriteConsoleMsg UserIndex, "Has enviado una solicitud de cambio al personaje " & dstName & ". Espera pronta noticias de él", FontTypeNames.FONTTYPE_INFO
60                MAO.LogMao "INVITACIÓN de parte de: " & UserList(UserIndex).Name & " hacia el personaje " & dstName
70            Else
80                WriteConsoleMsg UserIndex, "Hemos notado que el personaje al que enviaste la solicitud ya tiene muchas ofertas sin responder. Contáctate con él", FontTypeNames.FONTTYPE_INFO
90            End If
100       Else
110           WriteConsoleMsg UserIndex, ErrorMsg, FontTypeNames.FONTTYPE_INFO
120       End If
          

   On Error GoTo 0
   Exit Sub

Send_Invitation_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Send_Invitation of Módulo MAO in line " & Erl
          
End Sub

Public Sub Rechace_Invitation(ByVal UserIndex As Integer, _
                                ByVal dstName As String)
                        
   On Error GoTo Rechace_Invitation_Error

10        With UserList(UserIndex)
20            If Not SearchInvitationSent(.Name, dstName, True) Then
30                WriteConsoleMsg UserIndex, "El personaje " & dstName & " no te ha invitado a intercambiar personajes o bien ha cancelado la oferta.", FontTypeNames.FONTTYPE_INFO
40                Exit Sub
50            End If
              
60            KillOffer UserIndex, dstName
70            WriteConsoleMsg UserIndex, "Oferta rechazada con éxito", FontTypeNames.FONTTYPE_INFO
80        End With

   On Error GoTo 0
   Exit Sub

Rechace_Invitation_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Rechace_Invitation of Módulo MAO in line " & Erl
End Sub

Public Sub Cancel_Invitation(ByVal UserIndex As Integer, _
                                ByVal dstName As String)
   On Error GoTo Cancel_Invitation_Error

10        With UserList(UserIndex)
20            If Not SearchInvitation(.Name, dstName) Then
30                WriteConsoleMsg UserIndex, "No puedes cancelar una invitatión si no se realizó.", FontTypeNames.FONTTYPE_INFO
40                KillOfferSent UserIndex, dstName
50                Exit Sub
60            End If
              
              
70            KillOfferSent UserIndex, dstName
80        End With

   On Error GoTo 0
   Exit Sub

Cancel_Invitation_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Cancel_Invitation of Módulo MAO in line " & Erl
End Sub

Public Sub Accept_Invitation(ByVal UserIndex As Integer, _
                                ByVal dstName As String, _
                                ByVal Pin As String)
                                      
            
          Dim Name As String
          Dim dstNameIndex As Integer
          Dim PinChar As String
          
   On Error GoTo Accept_Invitation_Error

10        If StrComp(UCase$(UserList(UserIndex).Name), UCase$(dstName)) = 0 Then
20            MAO.LogMao "El personaje " & dstName & " se ha intentado comprar a si mismo."
30            Exit Sub
40        End If
          
50        PinChar = GetVar(App.Path & "\CHARFILE\" & UCase$(UserList(UserIndex).Name) & ".chr", "INIT", "PIN")
          
60        If SearchInvitationSent(UserList(UserIndex).Name, dstName) Then
70            If PinChar = Pin Then
80                MAO.LogMao "INTERCAMBIO EXITOSO: " & UserList(UserIndex).Name & " POR el personaje " & dstName
                  
90                Remove_Mercado UserList(UserIndex).Name
100               Remove_Mercado dstName
          
110               Name = UserList(UserIndex).Name
120               dstNameIndex = NameIndex(dstName)
                  
130               If dstNameIndex > 0 Then CloseSocket dstNameIndex
140               CloseSocket UserIndex
              
150               ChangeDataInfo Name, dstName

160           Else
170               WriteConsoleMsg UserIndex, "Has ingresado un PIN inválido. Has sido guardado en los logs del mercado y se podrá considerar como posible atentado de robo.", FontTypeNames.FONTTYPE_INFO
180               MAO.LogMao "El personaje " & Name & " ha puesto un pin inválido al querer intercambiar con " & dstName & "."
190           End If
              
              ' Change data info
              'WriteConsoleMsg Userindex, "Intercambio de personaje " & UserList(Userindex).Name & " POR el personaje " & dstName & " hecho exitosamente", FontTypeNames.FONTTYPE_INFO
              
200       Else
210           WriteConsoleMsg UserIndex, "El personaje ha borrado la invitación de cambio. También se te borrará a ti.", FontTypeNames.FONTTYPE_INFO
220           KillOffer UserIndex, dstName
230       End If
            

   On Error GoTo 0
   Exit Sub

Accept_Invitation_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Accept_Invitation of Módulo MAO in line " & Erl

End Sub

Private Sub ChangeDataInfo(ByVal Name As String, ByVal dstName As String)
          Dim Email(1) As String
          Dim Passwd(1) As String
          Dim Pin(1) As String
          
   On Error GoTo ChangeDataInfo_Error

10        Email(0) = GetVar(CharPath & UCase$(Name) & ".chr", "CONTACTO", "EMAIL")
20        Passwd(0) = GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "PASSWORD")
30        Pin(0) = GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "PIN")
          
40        Email(1) = GetVar(CharPath & UCase$(dstName) & ".chr", "CONTACTO", "EMAIL")
50        Passwd(1) = GetVar(CharPath & UCase$(dstName) & ".chr", "INIT", "PASSWORD")
60        Pin(1) = GetVar(CharPath & UCase$(dstName) & ".chr", "INIT", "PIN")
          
70        WriteVar CharPath & UCase$(Name) & ".chr", "CONTACTO", "EMAIL", Email(1)
80        WriteVar CharPath & UCase$(Name) & ".chr", "INIT", "PASSWORD", Passwd(1)
90        WriteVar CharPath & UCase$(Name) & ".chr", "INIT", "PIN", Pin(1)
          
100       WriteVar CharPath & UCase$(dstName) & ".chr", "CONTACTO", "EMAIL", Email(0)
110       WriteVar CharPath & UCase$(dstName) & ".chr", "INIT", "PASSWORD", Passwd(0)
120       WriteVar CharPath & UCase$(dstName) & ".chr", "INIT", "PIN", Pin(0)

   On Error GoTo 0
   Exit Sub

ChangeDataInfo_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure ChangeDataInfo of Módulo MAO in line " & Erl
          
End Sub


Private Function SearchInvitation(ByVal Name As String, ByVal dstName As String) As Boolean
          Dim i As Long
          
   On Error GoTo SearchInvitation_Error

10        SearchInvitation = False
20        For i = 1 To MAX_OFFER
30            If StrComp(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "Offer" & i), UCase$(Name)) = 0 Then
40                SearchInvitation = True
50                Exit For
60            End If
70        Next i

   On Error GoTo 0
   Exit Function

SearchInvitation_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure SearchInvitation of Módulo MAO in line " & Erl
          
End Function
Private Function SearchInvitationSent(ByVal Name As String, ByVal dstName As String, Optional ByVal Kill As Boolean = False) As Boolean
          Dim i As Long
          
   On Error GoTo SearchInvitationSent_Error

10        SearchInvitationSent = False
20        For i = 1 To MAX_OFFER
30            If StrComp(GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "OfferSent" & i), UCase$(Name)) = 0 Then
40                If Kill Then WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "OfferSent" & i, "0"
50                SearchInvitationSent = True
60                Exit For
70            End If
80        Next i

   On Error GoTo 0
   Exit Function

SearchInvitationSent_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure SearchInvitationSent of Módulo MAO in line " & Erl
          
End Function
Private Function FreeSlotOffer(ByVal UserIndex As Integer, ByVal dstName As String) As Boolean
          
          Dim tUser As Integer
          Dim i As Long
          
   On Error GoTo FreeSlotOffer_Error

10        FreeSlotOffer = -1
20        tUser = NameIndex(dstName)
          
30        FreeSlotOffer = False
          
40        If SearchInvitation(UserList(UserIndex).Name, dstName) Then
50            WriteConsoleMsg UserIndex, "Ya has realizado una oferta al personaje " & dstName & ". Aunque no aparezca en tu lista él usuario la tiene que borrar.", FontTypeNames.FONTTYPE_INFO
60            Exit Function
70        End If
          
80        If tUser > 0 Then
90            With UserList(tUser)
100               For i = 1 To MAX_OFFER
110                   If .Mercado.lstOffer(i) = "0" Then
120                       .Mercado.lstOffer(i) = UCase$(UserList(UserIndex).Name)
130                       WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "Offer" & i, UCase$(UserList(UserIndex).Name)
140                       Exit For
150                   End If
160               Next i
170           End With
180       Else
190           For i = 1 To MAX_OFFER
200               If (GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "Offer" & i)) = "0" Then
210                   WriteVar CharPath & UCase$(dstName) & ".chr", "MERCADO", "Offer" & i, UCase$(UserList(UserIndex).Name)
220                   Exit For
230               End If
240           Next i
250       End If
          
260       FreeSlotOffer = True

   On Error GoTo 0
   Exit Function

FreeSlotOffer_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure FreeSlotOffer of Módulo MAO in line " & Erl
End Function
Private Function Check_Invitation(ByVal UserIndex As Integer, ByVal dstName As String, ByRef ErrorMsg As String, ByRef Slot As Integer) As Boolean

   On Error GoTo Check_Invitation_Error

10        Check_Invitation = False
          
20        With UserList(UserIndex)
30            If StrComp(UCase$(.Name), UCase$(dstName)) = 0 Then
40                ErrorMsg = "¿What is your problem?"
50                MAO.LogMao "El personaje " & .Name & " se intentó comprar a si mismo."
60                Exit Function
70            End If
              
80            If Not FileExist(CharPath & UCase$(dstName) & ".chr") Then
90                ErrorMsg = "El personaje al que intentas invitar a realizar el cambio NO EXISTE"
100               MAO.LogMao "El personaje " & .Name & " intento comprar un PJ INEXISTENTE con NICK: " & dstName
110               Exit Function
120           End If
              
130           If GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "InList") = "0" Then
140               ErrorMsg = "El personaje al que intentas invitar a realizar el cambio NO ESTÁ EN EL MERCADO."
150               MAO.LogMao "El personaje " & .Name & " intentó comprar un PJ que no está en la lista con NICK: " & dstName
160               Exit Function
170           End If
              
180           If GetVar(CharPath & UCase$(dstName) & ".chr", "MERCADO", "Change") = "0" Then
190               ErrorMsg = "El personaje al que intentas invitar a realizar el cambio NO ESTÁ EN MODO CAMBIO."
200               MAO.LogMao "El personaje " & .Name & " podría haber intentado enviar cambio a un PJ EN VENTA. POSIBLE CHEAT"
210               Exit Function
220           End If
              
230           If SlotOffer(UserIndex, UCase$(dstName)) Then
240               ErrorMsg = "Ya has ofrecido una invitación al personaje " & dstName & "."
250               Exit Function
260           End If
              
270           Slot = FreeSlotOfferSent(UserIndex)
              
280           If Slot = -1 Then
290               ErrorMsg = "No tienes espacio para realizar más ofertas. Recuerda que son guardadas en tu personaje por seguridad."
300               Exit Function
310           End If
320       End With
          
330       Check_Invitation = True

   On Error GoTo 0
   Exit Function

Check_Invitation_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_Invitation of Módulo MAO in line " & Erl
End Function
Private Function FreeSlotOfferSent(ByVal UserIndex As Integer) As Integer
   On Error GoTo FreeSlotOfferSent_Error

10        FreeSlotOfferSent = -1
          
          ' ¿Hay lugar disponible para hacer una oferta más?
          Dim i As Long
          
20        With UserList(UserIndex)
30            For i = 1 To MAX_OFFER
40                If StrComp(.Mercado.lstOfferSent(i), "0") = 0 Then
50                    FreeSlotOfferSent = i
60                    Exit For
70                End If
80            Next i
90        End With

   On Error GoTo 0
   Exit Function

FreeSlotOfferSent_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure FreeSlotOfferSent of Módulo MAO in line " & Erl
End Function
Private Function SlotOffer(ByVal UserIndex As Integer, ByVal dstName As String, Optional ByVal Kill As Boolean = False) As Boolean
          
          ' ¿El personaje ya hizo una oferta al personaje?
          Dim i As Long
          
   On Error GoTo SlotOffer_Error

10        With UserList(UserIndex)
20            For i = 1 To MAX_OFFER
30                If StrComp(.Mercado.lstOfferSent(i), dstName) = 0 Then
40                    SlotOffer = True
50                    Exit For
60                End If
70            Next i
80        End With

   On Error GoTo 0
   Exit Function

SlotOffer_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure SlotOffer of Módulo MAO in line " & Erl
          
End Function

Private Function KillOffer(ByVal UserIndex As Integer, ByVal dstName As String) As Boolean
          ' ¿El personaje recibio una oferta y fue borrada? O bien, el personaje borra la oferta recibida (Rechazar)
          
          Dim i As Long
          
   On Error GoTo KillOffer_Error

10        KillOffer = False
20        With UserList(UserIndex)
30            For i = 1 To MAX_OFFER
40                If StrComp(.Mercado.lstOffer(i), UCase$(dstName)) = 0 Then
50                    .Mercado.lstOffer(i) = "0"
60                    WriteVar CharPath & UCase$(.Name) & ".chr", "MERCADO", "OFFER" & i, "0"
70                    KillOffer = True
80                    Exit For
90                End If
100           Next i
110       End With

   On Error GoTo 0
   Exit Function

KillOffer_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure KillOffer of Módulo MAO in line " & Erl
End Function

Private Function KillOfferSent(ByVal UserIndex As Integer, ByVal dstName As String) As Boolean
          
          Dim i As Long
          
   On Error GoTo KillOfferSent_Error

10        KillOfferSent = False
20        With UserList(UserIndex)
30            For i = 1 To MAX_OFFER
40                If StrComp(.Mercado.lstOfferSent(i), UCase$(dstName)) = 0 Then
50                    .Mercado.lstOfferSent(i) = "0"
60                    WriteVar CharPath & UCase$(.Name) & ".chr", "MERCADO", "OFFERSENT" & i, "0"
70                    KillOfferSent = True
80                    Exit For
90                End If
100           Next i
110       End With

   On Error GoTo 0
   Exit Function

KillOfferSent_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure KillOfferSent of Módulo MAO in line " & Erl
End Function
Private Function Check_Publication(ByVal UserIndex As Integer, _
                                    ByVal Email As String, _
                                    ByVal Passwd As String, _
                                    ByVal Pin As String, _
                                    ByRef ErrorMsg As String, _
                                    Optional ByVal dstName As String = vbNullString, _
                                    Optional ByVal Gld As Long = 0, _
                                    Optional ByVal Dsp As Long = 0) As Boolean
   On Error GoTo Check_Publication_Error

10        Check_Publication = False
          
          Dim LoadPasswd As String
          
20        With UserList(UserIndex)
30            If Gld < 0 Then
40                Call Ban(.Name, "Sistema", "Intento de dupeo.")
50                Call WriteErrorMsg(UserIndex, "Has sido baneado por el Sistema AntiCheat.")
60                Call FlushBuffer(UserIndex)
70                Call CloseSocket(UserIndex)
80                Exit Function
90            End If
              
100           If Dsp < 0 Then
110               Call Ban(.Name, "Sistema", "Intento de dupeo.")
120               Call WriteErrorMsg(UserIndex, "Has sido baneado por el Sistema AntiCheat.")
130               Call FlushBuffer(UserIndex)
140               Call CloseSocket(UserIndex)
150               Exit Function
160           End If
              
170           If StrComp(UCase$(.Pin), UCase$(Pin)) <> 0 Then
180               ErrorMsg = "¡¡ATENCIÓN!! El pín que ingresaste no pertenece al personaje."
190               Exit Function
200           End If
              
210           If StrComp(UCase$(.Email), UCase$(Email)) <> 0 Then
220               ErrorMsg = "¡¡ATENCIÓN!! El email que ingresaste no pertenece al personaje."
230               Exit Function
240           End If
              
250           LoadPasswd = UCase$(GetVar(CharPath & UCase$(.Name) & ".chr", "INIT", "Password"))
              
260           If StrComp(UCase$(Passwd), UCase$(LoadPasswd)) <> 0 Then
270               ErrorMsg = "¡¡ATENCIÓN!! La contraseña que has ingresado no pertenece al personaje."
280               Exit Function
290           End If
              
300           If dstName <> vbNullString Then
310               If Not FileExist(CharPath & UCase$(dstName) & ".chr", vbNormal) Then
320                   ErrorMsg = "¡¡ATENCIÓN!! El personaje al que deseas depositar lo que concrete la venta NO EXISTE"
330                   Exit Function
340               End If
350           End If
              
360           If .Mercado.InList > 0 Then
370               ErrorMsg = "¡¡ATENCIÓN!! El personaje que intentas publicar ya está en MERCADO DESTERIUM. Retiralo si no estás conforme."
380               Exit Function
390           End If
              
400       End With
          
410       Check_Publication = True

   On Error GoTo 0
   Exit Function

Check_Publication_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_Publication of Módulo MAO in line " & Erl
End Function

Private Function Add_Mercado(ByVal UserIndex As Integer, _
                        ByVal Name As String, _
                        ByRef Slot As Integer) As Boolean
   On Error GoTo Add_Mercado_Error

10        Slot = SlotMercado()
          
20        If Slot <> -1 Then
30            lstMercado(Slot) = UCase$(Name)
40            WriteVar App.Path & "\DAT\MAO.DAT", "INIT", "Pj" & Slot, UCase$(Name)
50            Add_Mercado = True
60        Else
70            Add_Mercado = False
80            WriteConsoleMsg UserIndex, "No hay mas espacio disponible en el Mercado de DesteriumAO", FontTypeNames.FONTTYPE_WARNING
90        End If

   On Error GoTo 0
   Exit Function

Add_Mercado_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Add_Mercado of Módulo MAO in line " & Erl
          
End Function

Private Function SlotMercado(Optional ByVal Name As String = vbNullString) As Integer
          
          ' @ Con argumentos = Buscamos el slot donde se encuentra el NICKNAME.
          ' @ Sin argumentos = Buscamos un nuevo slot para guardar el personaje en el MERCADO.
          
          Dim i As Long
          
   On Error GoTo SlotMercado_Error

10        SlotMercado = -1
          
20        If Name = vbNullString Then
30            For i = 1 To MAX_MAO
40                If lstMercado(i) = "0" Then
50                    SlotMercado = i
60                    Exit For
70                End If
80            Next i
90        Else
100           For i = 1 To MAX_MAO
110               If StrComp(lstMercado(i), UCase$(Name)) = 0 Then
120                   SlotMercado = i
130                   Exit For
140               End If
150           Next i
160       End If

   On Error GoTo 0
   Exit Function

SlotMercado_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure SlotMercado of Módulo MAO in line " & Erl
End Function

Public Sub ResetMercado(ByVal UserIndex As Integer)

          ' @ Reseteamos el mercado interno del char.
          Dim i As Integer
          
   On Error GoTo ResetMercado_Error

10        With UserList(UserIndex).Mercado
20            .Change = 0
30            .Dsp = 0
40            .dstName = "0"
50            .Gld = 0
60            .InList = 0
              
70            For i = 1 To MAX_OFFER
80                .lstOffer(i) = "0"
90                .lstOfferSent(i) = "0"
100           Next i
          
110       End With

   On Error GoTo 0
   Exit Sub

ResetMercado_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure ResetMercado of Módulo MAO in line " & Erl
End Sub


Sub LogMao(ByVal tempStr As String)
          Dim mifile As Integer
10        mifile = FreeFile
20        Open App.Path & "\logs\Mercado.log" For Append Shared As #mifile
30        Print #mifile, Date & " " & time & " :" & tempStr
40        Close #mifile

End Sub

Public Sub SendInfoCharMAO(ByVal UserIndex As Integer, ByVal UserName As String)
           ' • INFORMACIÓN DE UN PERSONAJE DEL MAO.
           
          Dim strTemp As String
          Dim Char As String
          
   On Error GoTo SendInfoCharMAO_Error

10        Char = CharPath & UCase$(UserName) & ".chr"
          
          If Not FileExist(Char, vbArchive) Then Exit Sub
          
20        With UserList(UserIndex)
30            If Not (GetVar(Char, "MERCADO", "InList")) > 0 And Not SearchInvitationSent(.Name, UserName) Then
40                Call LogMao("El personaje " & .Name & " ha intentado ver la información de un personaje que NO ESTA MAO ni que envió OFERTA. Nick: " & UserName)
50                Exit Sub
60            End If
              
70            strTemp = "INFORMACIÓN DE " & UCase$(UserName) & vbCrLf
80            strTemp = strTemp & "Clase/Raza: " & ListaClases(val(GetVar(Char, "INIT", "CLASE"))) & "/" & ListaRazas(val(GetVar(Char, "INIT", "RAZA"))) & vbCrLf
90            strTemp = strTemp & "Vida: " & val(GetVar(Char, "STATS", "MAXHP")) & vbCrLf
100           strTemp = strTemp & "Maná: " & val(GetVar(Char, "STATS", "MAXMAN")) & vbCrLf
110           strTemp = strTemp & "MinHit/MaxHit: " & val(GetVar(Char, "STATS", "MINHIT")) & "/" & val(GetVar(Char, "STATS", "MINHIT")) & vbCrLf
120           strTemp = strTemp & "Nivel: " & val(GetVar(Char, "STATS", "ELV")) & vbCrLf
130           strTemp = strTemp & "Oro: " & IIf(val(GetVar(Char, "FLAGS", "ORO")), "SI", "NO") & vbCrLf
140           strTemp = strTemp & "Premium: " & IIf(val(GetVar(Char, "FLAGS", "PREMIUM")) > 0, "SI", "NO") & vbCrLf
150           strTemp = strTemp & "Plata: " & IIf(val(GetVar(Char, "FLAGS", "PLATA")) > 0, "SI", "NO") & vbCrLf
160           strTemp = strTemp & "Bronce: " & IIf(val(GetVar(Char, "FLAGS", "BRONCE")) > 0, "SI", "NO") & vbCrLf
170           strTemp = strTemp & "Retos ganados / Retos perdidos: " & val(GetVar(Char, "RETOS", "RetosGanados")) & "/" & val(GetVar(Char, "RETOS", "RetosPerdidos")) & vbCrLf
180           strTemp = strTemp & "Monedas de oro TOTAL: " & (val(GetVar(Char, "STATS", "BANCO")) + val(GetVar(Char, "STATS", "GLD")))
190           strTemp = strTemp & "Famas utilizadas: " & val(GetVar(Char, "FLAGS", "BONOSHP"))
200           WriteConsoleMsg UserIndex, strTemp, FontTypeNames.FONTTYPE_INFO
210       End With
          

   On Error GoTo 0
   Exit Sub

SendInfoCharMAO_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure SendInfoCharMAO of Módulo MAO in line " & Erl
           
End Sub
