Attribute VB_Name = "wskapiAO"
'**************************************************************
' wskapiAO.bas
'
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
 
Option Explicit

''
' Modulo para manejar Winsock
'

#If UsarQueSocket = 1 Then


'Si la variable esta en TRUE , al iniciar el WsApi se crea
'una ventana LABEL para recibir los mensajes. Al detenerlo,
'se destruye.
'Si es FALSE, los mensajes se envian al form frmMain (o el
'que sea).
#Const WSAPI_CREAR_LABEL = True

Private Const SD_BOTH As Long = &H2

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const WS_CHILD = &H40000000
Public Const GWL_WNDPROC = (-4)

Private Const SIZE_RCVBUF As Long = 8192
Private Const SIZE_SNDBUF As Long = 8192

''
'Esto es para agilizar la busqueda del slot a partir de un socket dado,
'sino, la funcion BuscaSlotSock se nos come todo el uso del CPU.
'
' @param Sock sock
' @param slot slot
'
Public Type tSockCache
    Sock As Long
    Slot As Long
End Type

Public WSAPISock2Usr As New Collection

' ====================================================================================
' ====================================================================================

Public OldWProc As Long
Public ActualWProc As Long
Public hWndMsg As Long

' ====================================================================================
' ====================================================================================

Public SockListen As Long
Public LastSockListen As Long

#End If

' ====================================================================================
' ====================================================================================


Public Sub IniciaWsApi(ByVal hwndParent As Long)
#If UsarQueSocket = 1 Then

10    Call LogApiSock("IniciaWsApi")
20    Debug.Print "IniciaWsApi"

#If WSAPI_CREAR_LABEL Then
30    hWndMsg = CreateWindowEx(0, "STATIC", "AOMSG", WS_CHILD, 0, 0, 0, 0, hwndParent, 0, App.hInstance, ByVal 0&)
#Else
40    hWndMsg = hwndParent
#End If 'WSAPI_CREAR_LABEL

50    OldWProc = SetWindowLong(hWndMsg, GWL_WNDPROC, AddressOf WndProc)
60    ActualWProc = GetWindowLong(hWndMsg, GWL_WNDPROC)

      Dim desc As String
70    Call StartWinsock(desc)

#End If
End Sub

Public Sub LimpiaWsApi()
#If UsarQueSocket = 1 Then

10    Call LogApiSock("LimpiaWsApi")

20    If WSAStartedUp Then
30        Call EndWinsock
40    End If

50    If OldWProc <> 0 Then
60        SetWindowLong hWndMsg, GWL_WNDPROC, OldWProc
70        OldWProc = 0
80    End If

#If WSAPI_CREAR_LABEL Then
90    If hWndMsg <> 0 Then
100       DestroyWindow hWndMsg
110   End If
#End If

#End If
End Sub

Public Function BuscaSlotSock(ByVal S As Long) As Long
#If UsarQueSocket = 1 Then

10    On Error GoTo hayerror
          
20        If WSAPISock2Usr.Count <> 0 Then
30            BuscaSlotSock = WSAPISock2Usr.Item(CStr(S))
40        Else
50            BuscaSlotSock = -1
60        End If
70    Exit Function
          
hayerror:
80        BuscaSlotSock = -1
#End If

End Function

Public Sub AgregaSlotSock(ByVal Sock As Long, ByVal Slot As Long)
10    Debug.Print "AgregaSockSlot"
#If (UsarQueSocket = 1) Then

20    If WSAPISock2Usr.Count > MaxUsers Then
30        Call CloseSocket(Slot)
40        Exit Sub
50    End If

60    WSAPISock2Usr.Add CStr(Slot), CStr(Sock)

      'Dim Pri As Long, Ult As Long, Med As Long
      'Dim LoopC As Long
      '
      'If WSAPISockChacheCant > 0 Then
      '    Pri = 1
      '    Ult = WSAPISockChacheCant
      '    Med = Int((Pri + Ult) / 2)
      '
      '    Do While (Pri <= Ult) And (Ult > 1)
      '        If Sock < WSAPISockChache(Med).Sock Then
      '            Ult = Med - 1
      '        Else
      '            Pri = Med + 1
      '        End If
      '        Med = Int((Pri + Ult) / 2)
      '    Loop
      '
      '    Pri = IIf(Sock < WSAPISockChache(Med).Sock, Med, Med + 1)
      '    Ult = WSAPISockChacheCant
      '    For LoopC = Ult To Pri Step -1
      '        WSAPISockChache(LoopC + 1) = WSAPISockChache(LoopC)
      '    Next LoopC
      '    Med = Pri
      'Else
      '    Med = 1
      'End If
      'WSAPISockChache(Med).Slot = Slot
      'WSAPISockChache(Med).Sock = Sock
      'WSAPISockChacheCant = WSAPISockChacheCant + 1

#End If
End Sub

Public Sub BorraSlotSock(ByVal Sock As Long)
#If (UsarQueSocket = 1) Then
      Dim cant As Long

10    cant = WSAPISock2Usr.Count
20    On Error Resume Next
30    WSAPISock2Usr.Remove CStr(Sock)

40    Debug.Print "BorraSockSlot " & cant & " -> " & WSAPISock2Usr.Count

#End If
End Sub



Public Function WndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#If UsarQueSocket = 1 Then

10    On Error Resume Next

          Dim ret As Long
          Dim tmp() As Byte
          Dim S As Long
          Dim E As Long
          Dim n As Integer
          Dim UltError As Long
          
20        Select Case msg
              Case 1025
30                S = wParam
40                E = WSAGetSelectEvent(lParam)
                  
50                Select Case E
                      Case FD_ACCEPT
60                        If S = SockListen Then
70                            Call EventoSockAccept(S)
80                        End If
                      
                  '    Case FD_WRITE
                  '        N = BuscaSlotSock(s)
                  '        If N < 0 And s <> SockListen Then
                  '            'Call apiclosesocket(s)
                  '            call WSApiCloseSocket(s)
                  '            Exit Function
                  '        End If
                  '
                  
                  '        Call IntentarEnviarDatosEncolados(N)
                  '
                  '        Dale = UserList(N).ColaSalida.Count > 0
                  '        Do While Dale
                  '            Ret = WsApiEnviar(N, UserList(N).ColaSalida.Item(1), False)
                  '            If Ret <> 0 Then
                  '                If Ret = WSAEWOULDBLOCK Then
                  '                    Dale = False
                  '                Else
                  '                    'y aca que hacemo' ?? help! i need somebody, help!
                  '                    Dale = False
                  '                    Debug.Print "ERROR AL ENVIAR EL DATO DESDE LA COLA " & Ret & ": " & GetWSAErrorString(Ret)
                  '                End If
                  '            Else
                  '            '    Debug.Print "Dato de la cola enviado"
                  '                UserList(N).ColaSalida.Remove 1
                  '                Dale = (UserList(N).ColaSalida.Count > 0)
                  '            End If
                  '        Loop
              
90                    Case FD_READ
100                       n = BuscaSlotSock(S)
110                       If n < 0 And S <> SockListen Then
                              'Call apiclosesocket(s)
120                           Call WSApiCloseSocket(S)
130                           Exit Function
140                       End If
                          
                          'create appropiate sized buffer
150                       ReDim Preserve tmp(SIZE_RCVBUF - 1) As Byte
                          
160                       ret = recv(S, tmp(0), SIZE_RCVBUF, 0)
                          ' Comparo por = 0 ya que esto es cuando se cierra
                          ' "gracefully". (mas abajo)
170                       If ret < 0 Then
180                           UltError = Err.LastDllError
190                           If UltError = WSAEMSGSIZE Then
200                               Debug.Print "WSAEMSGSIZE"
210                               ret = SIZE_RCVBUF
220                           Else
230                               Debug.Print "Error en Recv: " & GetWSAErrorString(UltError)
240                               Call LogApiSock("Error en Recv: N=" & n & " S=" & S & " Str=" & GetWSAErrorString(UltError))
                                  
                                  'no hay q llamar a CloseSocket() directamente,
                                  'ya q pueden abusar de algun error para
                                  'desconectarse sin los 10segs. CREEME.
250                               Call CloseSocketSL(n)
260                               Call Cerrar_Usuario(n)
270                               Exit Function
280                           End If
290                       ElseIf ret = 0 Then
300                           Call CloseSocketSL(n)
310                           Call Cerrar_Usuario(n)
320                       End If
                          
330                       ReDim Preserve tmp(ret - 1) As Byte
                          
340                       Call EventoSockRead(n, tmp)
                      
350                   Case FD_CLOSE
360                       n = BuscaSlotSock(S)
370                       If S <> SockListen Then Call apiclosesocket(S)
                          
380                       If n > 0 Then
390                           Call BorraSlotSock(S)
400                           UserList(n).ConnID = -1
410                           UserList(n).ConnIDValida = False
420                           Call EventoSockClose(n)
430                       End If
440               End Select
              
450           Case Else
460               WndProc = CallWindowProc(OldWProc, hWnd, msg, wParam, lParam)
470       End Select
#End If
End Function

'Retorna 0 cuando se envió o se metio en la cola,
'retorna <> 0 cuando no se pudo enviar o no se pudo meter en la cola
Public Function WsApiEnviar(ByVal Slot As Integer, ByRef Str As String) As Long
#If UsarQueSocket = 1 Then
          Dim ret As String
          Dim Retorno As Long
          Dim data() As Byte
          
10        ReDim Preserve data(Len(Str) - 1) As Byte

20        data = StrConv(Str, vbFromUnicode)
          
#If SeguridadAlkon Then
30        Call Security.DataSent(Slot, data)
#End If
          
40        Retorno = 0
          
50        If UserList(Slot).ConnID <> -1 And UserList(Slot).ConnIDValida Then
60            ret = Send(ByVal UserList(Slot).ConnID, data(0), ByVal UBound(data()) + 1, ByVal 0)
70            If ret < 0 Then
80                ret = Err.LastDllError
90                If ret = WSAEWOULDBLOCK Then
                      
#If SeguridadAlkon Then
100                   Call Security.DataStored(Slot)
#End If
                      
                      ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
110                   Call UserList(Slot).outgoingData.WriteASCIIStringFixed(Str)
120               End If
130           End If
140       ElseIf UserList(Slot).ConnID <> -1 And Not UserList(Slot).ConnIDValida Then
150           If Not UserList(Slot).Counters.Saliendo Then
160               Retorno = -1
170           End If
180       End If
          
190       WsApiEnviar = Retorno
#End If
End Function

Public Sub LogApiSock(ByVal Str As String)
#If (UsarQueSocket = 1) Then

10    On Error GoTo Errhandler

      Dim nfile As Integer
20    nfile = FreeFile ' obtenemos un canal
30    Open App.Path & "\logs\wsapi.log" For Append Shared As #nfile
40    Print #nfile, Date & " " & time & " " & Str
50    Close #nfile

60    Exit Sub

Errhandler:

#End If
End Sub

Public Sub EventoSockAccept(ByVal SockID As Long)
#If UsarQueSocket = 1 Then
      '==========================================================
      'USO DE LA API DE WINSOCK
      '========================
          
          Dim NewIndex As Integer
          Dim ret As Long
          Dim Tam As Long, sa As sockaddr
          Dim NuevoSock As Long
          Dim i As Long
          Dim tStr As String
          
10        Tam = sockaddr_size
          
          '=============================================
          'SockID es en este caso es el socket de escucha,
          'a diferencia de socketwrench que es el nuevo
          'socket de la nueva conn
          
      'Modificado por Maraxus
          'Ret = WSAAccept(SockID, sa, Tam, AddressOf CondicionSocket, 0)
20        ret = accept(SockID, sa, Tam)

30        If ret = INVALID_SOCKET Then
40            i = Err.LastDllError
50            Call LogCriticEvent("Error en Accept() API " & i & ": " & GetWSAErrorString(i))
60            Exit Sub
70        End If
          
80        If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
90            Call WSApiCloseSocket(NuevoSock)
100           Exit Sub
110       End If

          'If Ret = INVALID_SOCKET Then
          '    If Err.LastDllError = 11002 Then
          '        ' We couldn't decide if to accept or reject the connection
          '        'Force reject so we can get it out of the queue
          '        Ret = WSAAccept(SockID, sa, Tam, AddressOf CondicionSocket, 1)
          '        Call LogCriticEvent("Error en WSAAccept() API 11002: No se pudo decidir si aceptar o rechazar la conexión.")
          '    Else
          '        i = Err.LastDllError
          '        Call LogCriticEvent("Error en WSAAccept() API " & i & ": " & GetWSAErrorString(i))
          '        Exit Sub
          '    End If
          'End If

120       NuevoSock = ret
          
          'Seteamos el tamaño del buffer de entrada
130       If setsockopt(NuevoSock, SOL_SOCKET, SO_RCVBUFFER, SIZE_RCVBUF, 4) <> 0 Then
140           i = Err.LastDllError
150           Call LogCriticEvent("Error al setear el tamaño del buffer de entrada " & i & ": " & GetWSAErrorString(i))
160       End If
          'Seteamos el tamaño del buffer de salida
170       If setsockopt(NuevoSock, SOL_SOCKET, SO_SNDBUFFER, SIZE_SNDBUF, 4) <> 0 Then
180           i = Err.LastDllError
190           Call LogCriticEvent("Error al setear el tamaño del buffer de salida " & i & ": " & GetWSAErrorString(i))
200       End If

          'If SecurityIp.IPSecuritySuperaLimiteConexiones(sa.sin_addr) Then
              'tStr = "Limite de conexiones para su IP alcanzado."
              'Call send(ByVal NuevoSock, ByVal tStr, ByVal Len(tStr), ByVal 0)
              'Call WSApiCloseSocket(NuevoSock)
              'Exit Sub
          'End If
          
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '   BIENVENIDO AL SERVIDOR!!!!!!!!
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          
          'Mariano: Baje la busqueda de slot abajo de CondicionSocket y limite x ip
210       NewIndex = NextOpenUser ' Nuevo indice
          
220       If NewIndex <= MaxUsers Then
              
              'Make sure both outgoing and incoming data buffers are clean
230           Call UserList(NewIndex).incomingData.ReadASCIIStringFixed(UserList(NewIndex).incomingData.length)
240           Call UserList(NewIndex).outgoingData.ReadASCIIStringFixed(UserList(NewIndex).outgoingData.length)

#If SeguridadAlkon Then
250           Call Security.NewConnection(NewIndex)
#End If
              
260           UserList(NewIndex).ip = GetAscIP(sa.sin_addr)
              'Busca si esta banneada la ip
270           For i = 1 To BanIps.Count
280               If BanIps.Item(i) = UserList(NewIndex).ip Then
                      'Call apiclosesocket(NuevoSock)
290                   Call WriteErrorMsg(NewIndex, "Su IP se encuentra bloqueada en este servidor.")
300                   Call FlushBuffer(NewIndex)
                      'Call SecurityIp.IpRestarConexion(sa.sin_addr)
310                   Call WSApiCloseSocket(NuevoSock)
320                   Exit Sub
330               End If
340           Next i
              
350           If NewIndex > LastUser Then LastUser = NewIndex
              
360           UserList(NewIndex).ConnID = NuevoSock
370           UserList(NewIndex).ConnIDValida = True
              
380           Call AgregaSlotSock(NuevoSock, NewIndex)
390       Else
              Dim Str As String
              Dim data() As Byte
              
400           Str = Protocol.PrepareMessageErrorMsg("El servidor se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")
              
410           ReDim Preserve data(Len(Str) - 1) As Byte
              
420           data = StrConv(Str, vbFromUnicode)
              
#If SeguridadAlkon Then
430           Call Security.DataSent(Security.NO_SLOT, data)
#End If
              
440           Call Send(ByVal NuevoSock, data(0), ByVal UBound(data()) + 1, ByVal 0)
450           Call WSApiCloseSocket(NuevoSock)
460       End If
          
#End If
End Sub

Public Sub EventoSockRead(ByVal Slot As Integer, ByRef Datos() As Byte)
#If UsarQueSocket = 1 Then
          Dim i As Long
          
10        If Slot = 0 Then
20            LogError "Hubo un SLOT CERO en EventoSockRead"
30        End If
          
40    With UserList(Slot)
            'Encriptacion dinamica by elSanto
            

        'encriptacion dinamica by elsanto
50        Call .incomingData.WriteBlock(Datos)

60        If .ConnID <> -1 Then
70            .LastPacketComplete = True
              
80            Do While .LastPacketComplete
90                .LastPacketComplete = HandleIncomingData(Slot)
100           Loop
110       Else
120           Exit Sub
130       End If
       
         
140   End With
       
#End If
End Sub

Public Sub EventoSockClose(ByVal Slot As Integer)
#If UsarQueSocket = 1 Then
          
          'Es el mismo user al que está revisando el centinela??
          'Si estamos acá es porque se cerró la conexión, no es un /salir, y no queremos banearlo....
10        If Centinela.RevisandoUserIndex = Slot Then _
              Call modCentinela.CentinelaUserLogout
          
#If SeguridadAlkon Then
20        Call Security.UserDisconnected(Slot)
#End If
          
30        If UserList(Slot).flags.UserLogged Then
40            Call CloseSocketSL(Slot)
50            Call Cerrar_Usuario(Slot)
60        Else
70            Call CloseSocket(Slot)
80        End If
#End If
End Sub


Public Sub WSApiReiniciarSockets()
#If UsarQueSocket = 1 Then
      Dim i As Long
          'Cierra el socket de escucha
10        If SockListen >= 0 Then Call apiclosesocket(SockListen)
          
          'Cierra todas las conexiones
20        For i = 1 To MaxUsers
30            If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
40                Call CloseSocket(i)
50            End If
              
              'Call ResetUserSlot(i)
60        Next i
          
70        For i = 1 To MaxUsers
80            Set UserList(i).incomingData = Nothing
90            Set UserList(i).outgoingData = Nothing
100       Next i
          
          ' No 'ta el PRESERVE :p
110       ReDim UserList(1 To MaxUsers)
120       For i = 1 To MaxUsers
130           UserList(i).ConnID = -1
140           UserList(i).ConnIDValida = False
              
150           Set UserList(i).incomingData = New clsByteQueue
160           Set UserList(i).outgoingData = New clsByteQueue
170       Next i
          
180       LastUser = 1
190       NumUsers = 0
          
200       Call LimpiaWsApi
210       Call Sleep(100)
220       Call IniciaWsApi(frmMain.hWnd)
230       SockListen = ListenForConnect(Puerto, hWndMsg, "")


#End If
End Sub

Public Sub WSApiCloseSocket(ByVal Socket As Long)
#If UsarQueSocket = 1 Then
10    Call WSAAsyncSelect(Socket, hWndMsg, ByVal 1025, ByVal (FD_CLOSE))
20    Call ShutDown(Socket, SD_BOTH)
#End If
End Sub

Public Function CondicionSocket(ByRef lpCallerId As WSABUF, ByRef lpCallerData As WSABUF, ByRef lpSQOS As FLOWSPEC, ByVal Reserved As Long, ByRef lpCalleeId As WSABUF, ByRef lpCalleeData As WSABUF, ByRef Group As Long, ByVal dwCallbackData As Long) As Long
#If UsarQueSocket = 1 Then
          Dim sa As sockaddr
          
          'Check if we were requested to force reject

10        If dwCallbackData = 1 Then
20            CondicionSocket = CF_REJECT
30            Exit Function
40        End If
          
           'Get the address

50        CopyMemory sa, ByVal lpCallerId.lpBuffer, lpCallerId.dwBufferLen

          
60        If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
70            CondicionSocket = CF_REJECT
80            Exit Function
90        End If

100       CondicionSocket = CF_ACCEPT 'En realdiad es al pedo, porque CondicionSocket se inicializa a 0, pero así es más claro....
#End If
End Function
