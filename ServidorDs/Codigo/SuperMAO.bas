Attribute VB_Name = "SuperMAO"
Option Explicit


Private Const SEPARATOR_OFFER As String = "|"
Public Const MAX_MAO_LIST As Byte = 200
Public Const MAX_OFFER As Byte = 30


Private Type tMaoAccount
    Users As String
    Recibidas As String
    Enviadas As String
    Gld As Long
    Dsp As Long
    Bloqued As Byte
End Type

Private Type tMao
    Tittle As String
    Account As String
End Type

Public MaoList(MAX_MAO_LIST) As tMao
Private Sub CreateMAO()
          Dim intFile As Integer
          Dim i As Integer
          
10        intFile = FreeFile

20        Open App.Path & "\DAT\MAO.DAT" For Output As #intFile
30        Print #intFile, "[INIT]"

40        For i = 1 To MAX_MAO_LIST
50            Print #intFile, i & "=-"
60        Next i
70        Close #intFile
End Sub
Public Sub LoadMAO()
        Dim Read As clsIniManager
        Dim A As Long
        Dim Temp As String
        
        If Not FileExist(App.Path & "\DAT\MAO.DAT") Then
            CreateMAO
        End If
          
        Set Read = New clsIniManager
          
        Read.Initialize App.Path & "\DAT\MAO.DAT"
          
        For A = 1 To MAX_MAO_LIST
            Temp = Read.GetValue("INIT", A)
            
            MaoList(A).Tittle = ReadField(1, Temp, Asc("-"))
            MaoList(A).Account = ReadField(2, Temp, Asc("-"))
        Next A
          
        Set Read = Nothing
End Sub
Private Function SlotFree() As Byte
    
    For SlotFree = 1 To MAX_MAO_LIST
        If MaoList(SlotFree).Account = vbNullString Then Exit Function
    Next SlotFree
    
    SlotFree = 255
End Function

Private Sub ResetUserMao(ByVal Slot As Byte)
    
    
    
    With MaoList(Slot)
        SaveDataAccount .Account, "MAO", "ACTIVE", "0"
        SaveDataAccount .Account, "MAO", "USERS", ""
        SaveDataAccount .Account, "MAO", "GLD", "0"
        SaveDataAccount .Account, "MAO", "DSP", "0"
        SaveDataAccount .Account, "MAO", "BLOQUED", "0"
        SaveDataAccount .Account, "MAO", "RECIBIDAS", ""
        SaveDataAccount .Account, "MAO", "ENVIADAS", ""
        
        .Tittle = 0
        .Account = 0
    End With
End Sub
' Comenzar una publicación.
Public Sub Mao_AddList(ByVal userIndex As Integer, _
                        ByVal Gld As Long, _
                        ByVal Dsp As Long, _
                        ByVal Users As String, _
                        ByVal Tittle As String, _
                        ByVal Bloqued As Byte)
    
On Error GoTo Errhandler
    Dim A As Long
    Dim list() As String
    Dim Slot As Long: Slot = SlotFree
        
10
    
    If (Gld < 0) Or (Dsp < 0) Or (Dsp > 32000) Or (Gld > 1000000000) Then Exit Sub
    
    If Len(Users) > 150 Then
        CloseSocket userIndex
        FlushBuffer userIndex
        Exit Sub
    End If
    
    If Len(Tittle) > 300 Then
        CloseSocket userIndex
        FlushBuffer userIndex
        Exit Sub
    End If
    
    If UCase$(Tittle) = "(VACIO)" Then
        CloseSocket userIndex
        FlushBuffer userIndex
        Exit Sub
    End If
        
    
    ' Lista llena
    If Slot = 255 Then
        WriteConsoleMsg userIndex, "La lista de ventas está llena. Espera a que algún lugar quede vacío.", FontTypeNames.FONTTYPE_INFO
        Exit Sub
    End If
30

    ' La cuenta ya está publicada de alguna forma.
    If val(LoadDataAccount(UserList(userIndex).Account, "MAO", "ACTIVE")) > 0 Then
        WriteConsoleMsg userIndex, "Ya has publicado personajes de esta cuenta. Quita la venta anterior para continuar", FontTypeNames.FONTTYPE_INFO
        Exit Sub
    End If
40

    list = Split(Users, "-")
50

    ' Checking account premium
    If UBound(list) > 0 Then
        If Not mCuenta.IsPremiumAccount(UCase$(UserList(userIndex).Account)) Then
            WriteConsoleMsg userIndex, "Solo CUENTAS PREMIUM pueden publicar más de 1 personaje. (Hasta 5))", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
    End If
    
    ' Anti Hacking
    For A = LBound(list) To UBound(list)
        If mCuenta.SearchCharAccount(UserList(userIndex).Account, list(A)) = 0 Then
            LogMao "ANTI HACK» La cuenta " & UserList(userIndex).Account & " con el personaje " & UserList(userIndex).Name & " el cual tiene IP: " & UserList(userIndex).ip & " ha intentado publicar con personajes que no son de él."
            WriteConsoleMsg userIndex, "Lo sentimos, pero no tienes acceso a esos personajes", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
    Next A
60

    ' Agregamos a la lista
    With MaoList(Slot)
        .Tittle = Tittle
        .Account = UserList(userIndex).Account
        
        WriteVar App.Path & "\DAT\MAO.DAT", "INIT", Slot, .Tittle & "-" & .Account
70
        SaveDataAccount .Account, "MAO", "Active", Slot
        SaveDataAccount .Account, "MAO", "USERS", Users
        SaveDataAccount .Account, "MAO", "GLD", Gld
        SaveDataAccount .Account, "MAO", "DSP", Dsp
        SaveDataAccount .Account, "MAO", "BLOQUED", Bloqued
        
    End With
80
    
    WriteConsoleMsg userIndex, "Publicación exitosa. Podrás quitar la misma cuando lo necesites. Recuerda que los datos de los personajes publicados serán visibles a los demás usuarios", FontTypeNames.FONTTYPE_INFO
    
Exit Sub
Errhandler:
LogMao "Error " & Err.Number & " (" & Err.Description & ") in procedure Mao_AddList in line " & Erl
End Sub

' Quitar la publicación.
Public Sub Mao_EndList(ByVal userIndex As Integer)
On Error GoTo Errhandler

    Dim Slot As Byte
10
    Slot = val(LoadDataAccount(UserList(userIndex).Account, "MAO", "Active"))
20
    ' Chequear que se haya hecho alguna publicación.
    If Slot = 0 Then
        WriteConsoleMsg userIndex, "No has realizado ninguna publicación.", FontTypeNames.FONTTYPE_INFO
        Exit Sub
    End If
30
    ' Reset and save
    ResetUserMao Slot
Exit Sub
Errhandler:
LogMao "Error " & Err.Number & " (" & Err.Description & ") in procedure Mao_AddList in line " & Erl
End Sub

' Enviar oferta a la publicación.
Public Sub Mao_SendOffer(ByVal userIndex As Integer, _
                            ByVal SlotMao As Byte, _
                            ByVal Users As String)
    
On Error GoTo Errhandler

    If (SlotMao <= 0) Or (SlotMao > MAX_MAO_LIST) Then Exit Sub
    
    Dim list() As String
    Dim A As Long, B As Long
    Dim Temp As String
    
    If UCase$(MaoList(SlotMao).Account) = UCase$(UserList(userIndex).Account) Then Exit Sub
    
    ' Por si cuando envia la solicitud, la publicación ya no existe.
    If (MaoList(SlotMao).Account = vbNullString) Then
            
        WriteConsoleMsg userIndex, "La publicación ha finalizado en este momento. ¡Te han ganado de mano! Mejor suerte para la próxima.", FontTypeNames.FONTTYPE_INFO
        Exit Sub
    End If
    
    '
    ' Para realizar una oferta: NO se puede ofrecer más de 1 personaje si no tenemos la cuenta PREMIUM.
    '
    If Users <> vbNullString Then
        list = Split(Users, "-")
    
        ' Checking account premium
        If UBound(list) > 0 Then
            If Not mCuenta.IsPremiumAccount(UCase$(UserList(userIndex).Account)) Then
                WriteConsoleMsg userIndex, "Solo CUENTAS PREMIUM pueden ofrecer más de 1 personaje. (Hasta 5))", FontTypeNames.FONTTYPE_INFO
                Exit Sub
            End If
        End If
        
        ' Anti Hacking
        For A = LBound(list) To UBound(list)
            If mCuenta.SearchCharAccount(UserList(userIndex).Account, list(A)) = 0 Then
                LogMao "ANTI HACK» La cuenta " & UserList(userIndex).Account & " con el personaje " & UserList(userIndex).Name & " el cual tiene IP: " & UserList(userIndex).ip & " ha intentado ofrecer personajes que no son de él."
                WriteConsoleMsg userIndex, "Lo sentimos, pero no tienes acceso a esos personajes", FontTypeNames.FONTTYPE_INFO
                Exit Sub
            End If
        Next A
    End If
    

    '
    ' ¿La cuenta ya le ofreció?
    '
    If CheckingOffer(False, MaoList(SlotMao).Account, UserList(userIndex).Account, Temp) Then
        WriteConsoleMsg userIndex, "Ya has ofrecido con tu cuenta.", FontTypeNames.FONTTYPE_INFO
        Exit Sub
    End If
    
    '
    ' Guardamos la solicitud en la cuenta del Mao
    '
    If Temp <> vbNullString Then Temp = Temp & "."
    SaveDataAccount MaoList(SlotMao).Account, "MAO", "RECIBIDAS", Temp & UCase$(UserList(userIndex).Account) & "|" & Users
    
    '
    ' Guardamos nuestra solicitud
    '
    Temp = LoadDataAccount(UserList(userIndex).Account, "MAO", "ENVIADAS")
    If Temp <> vbNullString Then Temp = Temp & "."
    SaveDataAccount UserList(userIndex).Account, "MAO", "ENVIADAS", Temp & UCase$(MaoList(SlotMao).Account) & "|" & Users
    
    WriteConsoleMsg userIndex, "Has enviado solicitud a la publicación " & MaoList(SlotMao).Tittle & ". Espera prontas noticias", FontTypeNames.FONTTYPE_INFO

Exit Sub
Errhandler:

End Sub

Private Function CheckingOffer(ByVal Send As Boolean, _
                                ByVal Account1 As String, _
                                ByVal Account2 As String, _
                                ByRef DataLoad As String) As Boolean
    Dim A As Long
    Dim TempOffer As String
    Dim list() As String
    
    If Send Then
        TempOffer = LoadDataAccount(Account1, "MAO", "ENVIADAS")
    Else
        TempOffer = LoadDataAccount(Account1, "MAO", "RECIBIDAS")
    End If
    
    
    If TempOffer = vbNullString Then Exit Function
    
    DataLoad = TempOffer
    
    list = Split(TempOffer, "|")
    
    For A = LBound(list) To UBound(list)
        If UCase$(list(A)) = Account2 Then
            CheckingOffer = True
            Exit For
        End If
    Next A
End Function



' Aceptar oferta recibida. Acá se intercambian datos y finaliza la operación.
Public Sub Mao_AcceptOffer(ByVal userIndex As Integer, _
                            ByVal SlotOffer As Byte)
    
    
    
    Dim Temp As String, TempAccount As String, TempUsersOffer As String
    Dim list() As String, listPjsMao() As String, listPjsOffer() As String
    Dim tName As String, tUser As Integer
    Dim Gld As Long, Dsp As Long, tmpGld As Long
    Dim Checking As Boolean
    Dim FilePath As String
    
    Dim A As Long
    Dim CantPjs As Byte
    
    Const MONEY_DSP As Integer = 880
    
    Temp = LoadDataAccount(UserList(userIndex).Account, "MAO", "RECIBIDAS"): list = Split(Temp, ".")
    Temp = ReadField(2, list(SlotOffer), Asc("|")): listPjsOffer = Split(Temp, "-"): TempUsersOffer = Temp
    
    TempAccount = ReadField(1, list(SlotOffer), Asc("|"))
    listPjsMao = Split(LoadDataAccount(UserList(userIndex).Account, "MAO", "USERS"), "-")
    
    Gld = val(LoadDataAccount(TempAccount, "MAO", "GLD"))
    Dsp = val(LoadDataAccount(TempAccount, "MAO", "DSP"))
    
    Checking = False
    
    ' El usuario ofrece personajes además de los DSP y ORO
    If TempUsersOffer <> vbNullString Then
        
        
        ' Chequeo que el personaje de la publicación pueda meter personajes extras
        Call SearchSlotAccount(UserList(userIndex).Account, CantPjs)
        
        If (CantPjs - (UBound(listPjsMao) + 1) + (UBound(listPjsOffer) + 1)) > MAX_PJS_ACCOUNT Then
            WriteConsoleMsg userIndex, "No tienes espacio en tu cuenta para aceptar estos personajes.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
        ' Chequeo que el personaje que ofrece pueda meter personajes extras
        Call SearchSlotAccount(TempAccount, CantPjs)
        
        If (CantPjs - (UBound(listPjsOffer) + 1) + UBound(listPjsMao)) > MAX_PJS_ACCOUNT Then
            WriteConsoleMsg userIndex, "La otra cuenta no puede recibir personajes nuevos ya que no cuenta con el espacio suficiente.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
    Else
    
        ' El personaje no recibió oferta de ningun pj, va a recibir dsp y oro pero NO tiene un pj para que se le deposite . _
            TIPICO HIJO DE RE MIL PUTA QUERIENDO BUGUEAR EL SISTEMA
            
        If Dsp Or Gld Then
            Call SearchSlotAccount(UserList(userIndex).Account, CantPjs)
            
            If CantPjs = 1 Then
                WriteConsoleMsg userIndex, "No puedes aceptar la oferta ya que como has pedido DSP/ORO por tu/s personajes no tenes un personaje secundario para recibir dichas monedas.", FontTypeNames.FONTTYPE_INFO
                Exit Sub
            End If
        End If
    End If

    
    ' En caso de que la publicación sea por ORO o DSP nos fijamos si podemos sacarlo de algun personaje. El GLD y el DSP tienen que estar en el mismo personaje para ser válida la compra.
    If Gld Or Dsp Then
        For A = 1 To MAX_PJS_ACCOUNT
            tName = LoadDataAccount(TempAccount, "ACCOUNT", "PERSONAJE" & A)
            
            If tName <> "0" Then
                tUser = NameIndex(tName)
                FilePath = CharPath & UCase$(tName) & ".chr"
                
                If tUser > 0 Then
                    If Not TieneObjetos(MONEY_DSP, Dsp, tUser) Then Checking = False
                    If Not UserList(tUser).Stats.Gld < Gld Then Checking = False
                Else
                    If Not TieneObjetosOffline(MONEY_DSP, Dsp, tName) Then Checking = False
                    tmpGld = val(GetVar(FilePath, "STATS", "GLD")): If tmpGld < Gld Then Checking = False
                End If
                
                If Checking Then Exit For
            End If
        Next A
        
        If Not Checking Then
            WriteConsoleMsg userIndex, "El personaje que envió la solicitud no cuenta con los DSP ni el ORO necesario para concretar la venta.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
        ' Le quitamos a la cuenta que ofreció los DSP y el ORO
        If tUser > 0 Then
            Call QuitarObjetos(MONEY_DSP, Dsp, tUser)
            UserList(tUser).Stats.Gld = UserList(tUser).Stats.Gld - Gld
        Else
            QuitarObjetosOffline MONEY_DSP, Dsp, tName
            WriteVar FilePath, "STATS", "GLD", CStr(tmpGld - Gld)
        End If
    End If

    
    ' Los personajes publicados son quitados de la cuenta
    For A = LBound(listPjsMao) To UBound(listPjsMao)
        tUser = NameIndex(listPjsMao(A))
        
        ' Cerramos la conexión a personajes logeados
        If tUser > 0 Then
            CloseSocket tUser
        End If
        
        mCuenta.UpdateAccountUserName UserList(userIndex).Account, listPjsMao(A), "0"
        
    Next A
    
    ' Recibió oferta de pjs? Los quitamos de uno, y se los ponemos al otro
    If TempUsersOffer <> vbNullString Then
        ' Pjs ofertados se van del ofertador hacia el vendedor
        For A = LBound(listPjsOffer) To UBound(listPjsOffer)
            tUser = NameIndex(listPjsOffer(A))
            
            ' Cerramos la conexión a personajes logeados
            If tUser > 0 Then
                CloseSocket tUser
            End If
            
            mCuenta.UpdateAccountUserName TempAccount, listPjsOffer(A), "0"
            mCuenta.AddCharAccount UserList(userIndex).Account, listPjsOffer(A)
        Next A
    End If
    
    ' El que ofertó recibe los pjs de la publicación
    For A = LBound(listPjsMao) To UBound(listPjsMao)
        mCuenta.AddCharAccount TempAccount, listPjsOffer(A)
    Next A
    
    
    ' / / / tERMINAR
    ' Ahora si con todos los pjs cambiados de lugar, pasamos a darle los DSP y el ORO a la persona que vendió, y para eso buscamos un personaje en la cuenta al cual depositarle.
    For A = 1 To MAX_PJS_ACCOUNT
        tName = LoadDataAccount(UserList(userIndex).Account, "ACCOUNT", "PERSONAJE" & A)
        
        If tName <> "0" Then
            tUser = NameIndex(tName)
            
            If tUser <= 0 Then
                
            Else
                WriteConsoleMsg userIndex, "Has recibido la recompensa pedida por la compra de tu/s personajes. Felicitaciones", FontTypeNames.FONTTYPE_INFO
                UserList(tUser).Stats.Gld = UserList(tUser).Stats.Gld + Gld
                'if not meteritemeninventario(
                'terminar
            End If
        End If
    Next A
    
    
End Sub


