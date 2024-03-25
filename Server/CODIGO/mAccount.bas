Attribute VB_Name = "mAccount"
' Reparado por Lorwik

Option Explicit

Public Const ACCOUNT_FORMAT               As String = ".acc"

Public Const ACCOUNT_MAX_CHARS            As Integer = 10

' Caracteres de la cuenta

Public Const ACCOUNT_MIN_CHARACTER_CHAR   As Byte = 3

Public Const ACCOUNT_MAX_CHARACTER_CHAR   As Byte = 15

Public Const ACCOUNT_MIN_CHARACTER_KEY    As Byte = 20

Public Const ACCOUNT_MIN_CHARACTER_PASSWD As Byte = 8

' Subasta de objetos
Public Const ACCOUNT_MAX_AUCTION_OBJ      As Byte = 10

Public Const ACCOUNT_AUCTION_TIME         As Long = 14400 ' Cuatro horas

' Carga los datos de una cuenta seleccionada
Public Sub LoadDataAccount(ByVal UserIndex As Integer, ByVal Email As String)
                           
    On Error GoTo ErrHandler
              
    Dim Manager As clsIniManager

    Dim A       As Long, B As Long

    Dim TempSTR As String

    Dim ln      As String
    
    Set Manager = New clsIniManager
    
    Manager.Initialize AccountPath & Email & ACCOUNT_FORMAT

    With UserList(UserIndex).Account
        .Email = Email
        .Key = Manager.GetValue("INIT", "KEY")
        .Passwd = Manager.GetValue("INIT", "PASSWD")
        .DateRegister = Manager.GetValue("INIT", "DATEREGISTER")
        .Premium = val(Manager.GetValue("INIT", "PREMIUM"))
        .CharsAmount = val(Manager.GetValue("INIT", "CHARSAMOUNT"))
        .DatePremium = Manager.GetValue("INIT", "DATEPREMIUM")
        
        For A = 1 To ACCOUNT_MAX_CHARS
            .Chars(A).Name = Manager.GetValue("CHARS", A)
            
            If .Chars(A).Name <> vbNullString Then
                Call Login_Char_LoadInfo(UserIndex, A, .Chars(A).Name)

            End If

        Next A
        
        .BancoInvent.NroItems = CInt(Manager.GetValue("BancoInventory", "CantidadItems"))
        
        For A = 1 To MAX_BANCOINVENTORY_SLOTS
            ln = Manager.GetValue("BancoInventory", "Obj" & A)
            .BancoInvent.Object(A).ObjIndex = CInt(ReadField(1, ln, 45))
            .BancoInvent.Object(A).Amount = CInt(ReadField(2, ln, 45))
        
        Next A
        
        .Gld = CLng(Manager.GetValue("INIT", "GLD"))
        .Eldhir = CLng(Manager.GetValue("INIT", "ELDHIR"))

        ' Mercado (Lista de Publicaciones que realizó la Cuenta)
        .MercaderSlot = val(Manager.GetValue("SALE", "LAST"))

    End With
    
    Set Manager = Nothing
    
    Exit Sub

ErrHandler:
    Set Manager = Nothing

End Sub

' Crea la nueva cuenta
Public Function SaveDataNew(ByVal Email As String, ByVal Passwd As String, ByVal Key As String) As Boolean

    '<EhHeader>
    On Error GoTo SaveDataNew_Err

    '</EhHeader>
                            
    Dim Manager As clsIniManager
    
    Set Manager = New clsIniManager
    
    If FileExist(AccountPath & Email & ".acc", vbArchive) Then
        Manager.Initialize AccountPath & Email & ".acc"

    End If
        
    Call Manager.ChangeValue("INIT", "KEY", Key)
    Call Manager.ChangeValue("INIT", "PASSWD", Passwd)

    Manager.DumpFile AccountPath & Email & ACCOUNT_FORMAT

    Set Manager = Nothing
    SaveDataNew = True
    
    '<EhFooter>
    Exit Function

SaveDataNew_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.SaveDataNew " & "at line " & Erl
    
    SaveDataNew = False
    Set Manager = Nothing
        
    '</EhFooter>
End Function

' Guarda los datos de una cuenta
Public Sub SaveDataAccount(ByVal UserIndex As Integer, _
                           ByVal Email As String, _
                           ByVal IP As String)
                            
    On Error GoTo ErrHandler

    Dim Manager   As clsIniManager

    Dim A         As Long

    Dim B         As Long
    
    Dim SaveChars As Boolean
    
    Set Manager = New clsIniManager
    
    If FileExist(AccountPath & Email & ACCOUNT_FORMAT) Then
        Manager.Initialize AccountPath & Email & ACCOUNT_FORMAT

    End If
    
    With UserList(UserIndex).Account
        Call Manager.ChangeValue("INIT", "KEY", .Key)
        Call Manager.ChangeValue("INIT", "PASSWD", .Passwd)
        Call Manager.ChangeValue("INIT", "DATEREGISTER", .DateRegister)
        Call Manager.ChangeValue("INIT", "PREMIUM", .Premium)
        Call Manager.ChangeValue("INIT", "CHARSAMOUNT", .CharsAmount)
        Call Manager.ChangeValue("INIT", "DATEPREMIUM", .DatePremium)
        
        'If SaveChars Then
        'For A = 1 To ACCOUNT_MAX_CHARS
        'Call Manager.ChangeValue("CHARS", A, .Chars(A).Name)
        'Next A
        'End If
        
        Call Manager.ChangeValue("BancoInventory", "CantidadItems", val(.BancoInvent.NroItems))

        For A = 1 To MAX_BANCOINVENTORY_SLOTS
            Call Manager.ChangeValue("BancoInventory", "Obj" & A, .BancoInvent.Object(A).ObjIndex & "-" & .BancoInvent.Object(A).Amount)
        Next A

        Call Manager.ChangeValue("INIT", "GLD", CStr(.Gld))
        Call Manager.ChangeValue("INIT", "ELDHIR", CStr(.Eldhir))
        
        If IP <> vbNullString Then
            Call SaveDataAccount_LastIP(Email, IP, Manager)

        End If
        
    End With
    
    Manager.DumpFile AccountPath & Email & ACCOUNT_FORMAT

    Set Manager = Nothing
    
    Exit Sub
ErrHandler:
    Set Manager = Nothing

End Sub

Private Sub SaveDataAccount_LastIP(ByVal Email As String, _
                                   ByVal IP As String, _
                                   ByRef Manager As clsIniManager)

    '<EhHeader>
    On Error GoTo SaveDataAccount_LastIP_Err

    '</EhHeader>
    
    Dim A As Long
    
    'First time around?
    If Manager.GetValue("INIT", "LASTIP1") = vbNullString Then
        Call Manager.ChangeValue("INIT", "LastIP1", IP & " - " & Date & ":" & Time)
        
        'Is it a different ip from last time?
    ElseIf IP <> Left$(Manager.GetValue("INIT", "LASTIP1"), InStr(1, Manager.GetValue("INIT", "LASTIP1"), " ") - 1) Then

        For A = 5 To 2 Step -1
            Call Manager.ChangeValue("INIT", "LASTIP" & A, Manager.GetValue("INIT", "LastIP" & CStr(A - 1)))
        Next A

        Call Manager.ChangeValue("INIT", "LASTIP1", IP & " - " & Date & ":" & Time)
        
    Else
        'Same ip, just update the date
        Call Manager.ChangeValue("INIT", "LASTIP1", IP & " - " & Date & ":" & Time)

    End If
    
    '<EhFooter>
    Exit Sub

SaveDataAccount_LastIP_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.SaveDataAccount_LastIP " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function LoginAccount(ByVal UserIndex As Integer, _
                             ByVal Email As String, _
                             ByVal Passwd As String) As Boolean

    '<EhHeader>
    On Error GoTo LoginAccount_Err

    '</EhHeader>
    
    Dim N As Integer
    
    If LenB(Email) <= 3 Or LenB(Passwd) <= 3 Then
        Call Protocol.Kick(UserIndex)
        
        Exit Function

    End If
    
    If Not CheckMailString(Email) Then
        Call Protocol.Kick(UserIndex)
        
        Exit Function

    End If
    
    '¿Este IP ya esta conectado?
    If AllowMultiLogins > 0 Then
        If UserList(UserIndex).IpAddress <> vbNullString Then
            If CheckForSameIP(UserIndex, UserList(UserIndex).IpAddress) >= AllowMultiLogins Then
                Call Protocol.Kick(UserIndex, "En este juego se permiten " & AllowMultiLogins & " conexiones simultaneas.")
    
                Exit Function
    
            End If

        End If

    End If

    If Not FileExist(AccountPath & Email & ACCOUNT_FORMAT) Then
        Call Protocol.Kick(UserIndex, "No existe ninguna cuenta bajo es nombre o bien la contraseña es incorrecta.")

        Exit Function

    End If

    #If Testeo = 0 Then

        If GetVar(AccountPath & Email & ACCOUNT_FORMAT, "INIT", "PASSWD") <> Passwd Then
            Call Protocol.Kick(UserIndex, "La contraseña de la cuenta ha sido modificada.")

            Exit Function

        End If

    #End If
        
    If val(GetVar(AccountPath & Email & ACCOUNT_FORMAT, "INIT", "BAN")) > 0 Then

        Dim tStr As String
        
        tStr = GetVar(AccountPath & Email & ACCOUNT_FORMAT, "PENAS", "DATEDAY")
            
        If tStr <> vbNullString Then
            If Format(Now, "dd/mm/yyyy") > tStr Then
                Call WriteVar(AccountPath & Email & ACCOUNT_FORMAT, "INIT", "BAN", "0")

            End If
                
        Else

            Dim Razon As String

            Dim Pena  As String: Pena = GetVar(AccountPath & Email & ACCOUNT_FORMAT, "PENAS", "CANT")

            Razon = GetVar(AccountPath & Email & ACCOUNT_FORMAT, "PENAS", "P" & Pena)
            Call Protocol.Kick(UserIndex, "Tu cuenta está Baneada en este servidor. RAZON: " & Razon)

            Exit Function

        End If

    End If

    Dim MaxLogged As Byte
            
'    If (CheckEmailLogged(LCase$(Email))) > 0 Then
'        Call Protocol.Kick(UserIndex, "La cuenta ha superado la máxima cantidad de usuarios permitidos en ella.")
'
'        Exit Function
'
'    End If

    NumUsers = NumUsers + 1
    UserList(UserIndex).AccountLogged = True
    UserList(UserIndex).Counters.TimeInactive = 0
    
    frmMain.lstDebug.AddItem "LoadDataAccount"
    
    Call LoadDataAccount(UserIndex, Email)

    '#If Classic = 1 Then
    Call WriteLoggedAccount(UserIndex, UserList(UserIndex).Account.Chars)
    '#Else
    '      Call WriteLoggedAccountBattle(UserIndex)
    ' #End If
    
    Call MostrarNumUsers
    
    N = FreeFile
    Open LogPath & "Connect.log" For Append Shared As #N
    Print #N, "La IP " & UserList(UserIndex).Account.Sec.IP_Public & " ha entrado al juego. UserIndex:" & UserIndex & " " & Time & " " & Date
    Close #N
    
    LoginAccount = True

    '<EhFooter>
    Exit Function

LoginAccount_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.LoginAccount " & "at line " & Erl

    '</EhFooter>
End Function

Private Function LoginAccount_Char_Check(ByVal UserIndex As Integer, _
                                         ByVal UserName As String, _
                                         ByVal Slot As Byte) As Boolean

    '<EhHeader>
    On Error GoTo LoginAccount_Char_Check_Err

    '</EhHeader>
    
    If Slot <= 0 Or Slot > ACCOUNT_MAX_CHARS Then
        ' Anti Hacking
        
        Call Protocol.Kick(UserIndex)
        Exit Function

    End If
        
    If Not PersonajeExiste(UserName) Then
        Call Protocol.Kick(UserIndex)
        Exit Function

    End If
        
    If UserName = vbNullString Then
        Call Protocol.Kick(UserIndex)
        Exit Function

    End If
    
    With UserList(UserIndex)
    
        If .flags.UserLogged Then
            Call Protocol.Kick(UserIndex)
            ' Anti Hacking:: Chequeo en el cliente
            Exit Function

        End If

        If Not StrComp(UCase$(.Account.Chars(Slot).Name), UCase$(UserName)) = 0 Then
            Call Protocol.Kick(UserIndex)
            Exit Function

        End If
            
        ' # LION
            
        Dim tUser As Integer

        tUser = NameIndex(UserName)
            
        If tUser > 0 And tUser <> UserIndex Then
            Call Protocol.Kick(User)        ' Cerramos personaje bug
            Call LogError("Solucionamos bug provisoriamente by LoginAccount_Char_Check NICK: " & UserName)
            Exit Function

        End If
    
    End With
    
    LoginAccount_Char_Check = True
    
    '<EhFooter>
    Exit Function

LoginAccount_Char_Check_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.LoginAccount_Char_Check " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub LoginAccount_Char(ByVal UserIndex As Integer, _
                             ByVal UserName As String, _
                             ByVal Key As String, _
                             ByVal Slot As Byte, _
                             ByVal NewChar As Boolean)

    '<EhHeader>
    On Error GoTo LoginAccount_Char_Err

    '</EhHeader>
    
    Dim SlotUserName As Byte
    
    If Not LoginAccount_Char_Check(UserIndex, UCase$(UserName), Slot) Then Exit Sub
    
    ' Actualizo el nombre a como lo puse
    Call WriteVar(AccountPath & UserList(UserIndex).Account.Email & ".acc", "CHARS", CStr(Slot), UserName)

    Call ConnectUser(UserIndex, UserName, NewChar)

    UserList(UserIndex).Counters.TimeInactive = 0
    UserList(UserIndex).Account.SlotLogged = Slot
    '<EhFooter>
    Exit Sub

LoginAccount_Char_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.LoginAccount_Char " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub LoginAccount_ChangeAlias(ByVal UserIndex As Integer, ByVal UserName As String)

    '<EhHeader>
    On Error GoTo LoginAccount_ChangeAlias_Err

    '</EhHeader>

    If Not ValidarNombre(UserName) Then
        Call Protocol.Kick(UserIndex)
            
        Exit Sub

    End If
    
    If Not NombrePermitido(UCase$(UserName)) Then
        Call WriteErrorMsg(UserIndex, "El nombre no está permitido en estas tierras. Elige otro dentro de la fantasía que admite el juego.")

        Exit Sub

    End If
    
    If Not NameCheckReserve(UserList(UserIndex).Account.Email, UCase$(UserName)) Then
        Call WriteErrorMsg(UserIndex, "Parece que el nombre se encuentra reservado para que pueda ser creado únicamente por su dueño...")

        Exit Sub
            
    End If
    
    UserList(UserIndex).Account.Alias = UCase$(UserName)

    '<EhFooter>
    Exit Sub

LoginAccount_ChangeAlias_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.LoginAccount_ChangeAlias " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub ChangeNickChar(ByVal UserIndex As Integer, ByVal UserName As String)

    If Not ValidarNombre(UserName) Then Exit Sub
    
    Dim OldChar      As String
        
    Dim FilePath_Old As String

    FilePath_Old = CharPath & UCase$(UserList(UserIndex).Name) & ".chr"
    
    Dim FilePath As String

    FilePath = CharPath & UCase$(UserName) & ".chr"
    
    Dim FilePath_Copy As String

    FilePath_Copy = Replace(CharPath, "CHARS\CHARFILE", "CHARS\CHARFILE_COPY") & UCase$(UserName) & ".chr"
    
    If PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(UserIndex, "El nombre del personaje ya está siendo utilizado por otro usuario.", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
    
    If Mercader_CheckUsers(UserList(UserIndex).Account.MercaderSlot, UCase$(UserName)) Then
        Call WriteErrorMsg(UserIndex, "¡El personaje se encuentra en una publicación! Remueve la misma antes de cambiar el nombre.")
        Exit Sub

    End If

    If Not TieneObjetos(ACTA_NACIMIENTO, 1, UserIndex) Then Exit Sub
    Call QuitarObjetos(ACTA_NACIMIENTO, 1, UserIndex)
        
    With UserList(UserIndex)
        OldChar = .Name
            
        If .GuildIndex > 0 Then
            GuildsInfo(.GuildIndex).Members(.GuildSlot).Name = UCase$(UserName)
            GuildsInfo(.GuildIndex).Members(.GuildSlot).Char.Name = UCase$(UserName)
            
            Call LogError("Personaje " & OldChar & " paso a llamarse " & .Name & " y cambio de lider SLOT-USER-GUILD: " & .GuildSlot & ", GuildIndex: " & .GuildIndex)
            
            Call Guilds_Save(.GuildIndex)
            
            Call SaveUser(UserList(UserIndex), CharPath & UCase$(OldChar) & ".chr")
            
            Call LogError("Personaje " & OldChar & " paso a llamarse " & .Name & " y cambio de lider SLOT-USER-GUILD: " & .GuildSlot & ", GuildIndex: " & .GuildIndex)

        End If
       
        Call FileCopy(FilePath_Old, FilePath_Copy)
              
        .Name = UserName
        .secName = .Name
        .Account.Chars(.Account.SlotLogged).Name = .Name
        
        Call WriteVar(CharPath & UCase$(UserName) & ".chr", "GUILD", "GUILDRANGE", CStr(.GuildRange))
        Call WriteVar(CharPath & UCase$(UserName) & ".chr", "GUILD", "GUILDINDEX", CStr(.GuildIndex))
        Call WriteVar(CharPath & UCase$(UserName) & ".chr", "INIT", "ACCOUNTSLOT", CStr(.Account.SlotLogged))
        Call WriteVar(CharPath & UCase$(UserName) & ".chr", "INIT", "ACCOUNTNAME", CStr(UserList(UserIndex).Account.Email))
        Call WriteVar(AccountPath & UserList(UserIndex).Account.Email & ".acc", "CHARS", CStr(.Account.SlotLogged), UserName)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeNick(.Char.charindex, UserName))
        Call WriteConsoleMsg(UserIndex, "¡Ahora pasaste a llamarte " & .Name & "!", FontTypeNames.FONTTYPE_INFOGREEN)
        Call WriteLoggedAccount_DataChar(UserIndex, .Account.SlotLogged, .Account.Chars(.Account.SlotLogged))
        Call mAccount.SaveDataAccount(UserIndex, .Account.Email, .IpAddress)
        Call Kill(FilePath_Old)
              
        Call Logs_Security(eLog.eSecurity, eLogSecurity.eAntiHack, "CAMBIO DE NICK» " & UCase$(OldChar) & " pasó a llamarse " & UCase$(UserName))

    End With
    
End Sub

Public Function Account_FreeChar(ByVal UserIndex As Integer) As Byte

    '<EhHeader>
    On Error GoTo Account_FreeChar_Err

    '</EhHeader>
    Dim A As Long
    
    For A = 1 To ACCOUNT_MAX_CHARS

        If UserList(UserIndex).Account.Chars(A).Name = vbNullString Then
            Account_FreeChar = A
            Exit Function

        End If

    Next A
    
    '<EhFooter>
    Exit Function

Account_FreeChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.Account_FreeChar " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub LoginAccount_CharNew(ByVal UserIndex As Integer, ByVal UserName As String, ByVal Class As Byte, ByVal Raze As Byte, ByVal Genero As Byte, ByVal Head As Integer)

    '<EhHeader>
    On Error GoTo LoginAccount_CharNew_Err

    '</EhHeader>
    
    Dim Slot As Byte
            
    Slot = Account_FreeChar(UserIndex)
            
    If Slot = 0 Or Slot > mAccount.ACCOUNT_MAX_CHARS Then
        ' Anti Hacking
        Call Protocol.Kick(UserIndex)
        Exit Sub

    End If
            
    If Class = 0 Or Class > NUMCLASES Then
        Call Protocol.Kick(UserIndex)
        ' Anti Hacking
        Exit Sub

    End If
    
    If Raze = 0 Or Raze > NUMRAZAS Then
        Call Protocol.Kick(UserIndex)
        ' Anti Hacking
        Exit Sub

    End If

    If Genero = 0 Or Genero > 2 Then
        Call Protocol.Kick(UserIndex)
        ' Anti Hacking
        Exit Sub

    End If
    
    If Not ValidarNombre(UserName) Then
        Call Protocol.Kick(UserIndex)
            
        Exit Sub

    End If
    
    With UserList(UserIndex)
        
        If .Account.Chars(Slot).Name <> vbNullString Then
            Call Protocol.Kick(UserIndex)
            ' Anti Hacking::
            Exit Sub

        End If
        
        If .Account.CharsAmount = ACCOUNT_MAX_CHARS Then
            Call Protocol.Kick(UserIndex)
            ' Anti Hacking:: Chequeo en el cliente
            Exit Sub

        End If
    
        If .flags.UserLogged Then
            Call Protocol.Kick(UserIndex)
            ' Anti Hacking:: Chequeo en el cliente
            Exit Sub

        End If
        
        ' Anti Hack de Cabezas
        'If .Account.Premium > 0 Then
        '  If Not ValidarCabeza(Raze, Genero, Head) Then
        '    Call Protocol.Kick(UserIndex)
        '    Call Logs_Security(eSecurity, eAntiHack, "El usuario " & UserName & " ha seleccionado la cabeza " & Head & " desde la IP " & .IpAddress)
        '  Exit Sub

        '  End If

        '  End If
        
        If PuedeCrearPersonajes = 0 Then
            Call WriteErrorMsg(UserIndex, "Por el momento no se permite la creación de nuevos personajes.")

            Exit Sub

        End If
            
        If aClon.MaxPersonajes(UserList(UserIndex).IpAddress) Then
            Call Protocol.Kick(UserIndex, "Creemos que has creado demasiados personajes. Bajale la espuma a tu chocolate")
            Exit Sub

        End If
        
        If Not NombrePermitido(UCase$(UserName)) Then
            Call WriteErrorMsg(UserIndex, "El nombre no está permitido en estas tierras. Elige otro dentro de la fantasía que admite el juego.")

            Exit Sub

        End If
        
        If PersonajeExiste(UserName) Then
            Call WriteErrorMsg(UserIndex, "El personaje ya existe.")

            Exit Sub

        End If
            
        If Not NameCheckReserve(.Account.Email, UCase$(UserName)) Then
            Call WriteErrorMsg(UserIndex, "Parece que el nombre se encuentra reservado para que pueda ser creado únicamente por su dueño...")

            Exit Sub
            
        End If
        
        '.Account.Chars(Slot).Name = UserName
        'Call WriteVar(AccountPath & .Email & ACCOUNT_FORMAT, "CHARS", Slot, UserName)
        'Call WriteVar(CharPath & UCase$(UserName) & ".chr", "INIT", "ACCOUNTSLOT", Slot)
            
        ' Dim CopyAccount As tAccount
        ' CopyAccount = .Account
        Call ConnectNewUser(UserName, Class, Raze, Genero, Head, UserList(UserIndex))
        Call LoginAccount_SetChar(UserIndex, UserName, Slot, 1)
        'UserList(UserIndex).AccountLogged = True
        '     UserList(UserIndex).Counters.TimeInactive = 0
        '   UserList(UserIndex).Account = CopyAccount
              
        ' Stats iniciales
        Call InitialUserStats(UserList(UserIndex))
                
        ' Set Inicial
        Call ApplySetInitial_Newbie(UserIndex)

        #If Classic = 1 Then
            ' Set Spells
            Call ApplySpellsStats(UserIndex)
        #End If
                
        Call SaveUser(UserList(UserIndex), CharPath & UCase$(UserName) & ".chr")
        Call LoginAccount_Char(UserIndex, UserName, .Account.Key, Slot, True)
        
        ' Setting New
        Login_InfoAccountChars UserIndex, Slot, UserName
        'Call Login_Char_LoadInfo(UserIndex, Slot, UserName)
        Call WriteLoggedAccount(UserIndex, .Account.Chars)
        'Call WriteLoggedAccount_DataChar(UserIndex, Slot, .Account.Chars(Slot))
        Call Logs_Security(eSecurity, eLogSecurity.eNewChar, "Personaje " & UserName & " en la cuenta: " & UserList(UserIndex).Account.Email & ". IP: " & UserList(UserIndex).IpAddress)
            
        ' nro34 es la quest inicial (Newbie)
        #If Classic = 1 Then
            Call Quest_SetUserPrincipa(UserIndex)
                    
        #Else
            'Call Quest_SetUser(UserIndex, 69)
        #End If
            
    End With
    
    '<EhFooter>
    Exit Sub

LoginAccount_CharNew_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.LoginAccount_CharNew " & "at line " & Erl

    '</EhFooter>
End Sub

Private Function NameCheckReserve(ByVal Email As String, _
                                  ByVal UserName As String) As Boolean
        
    If StrComp(UserName, "LION") = 0 Or StrComp(UserName, "LAUTARO") = 0 Then
        If Not StrComp(Email, "marinolauta@gmail.com") = 0 Then
            NameCheckReserve = False
            Exit Function

        End If

    End If
    
    If StrComp(UserName, "KEOL") = 0 Then
        If Not StrComp(Email, "santiagoeschira@gmail.com") = 0 Then
            NameCheckReserve = False
            Exit Function

        End If

    End If
    
    If StrComp(UserName, "TAROT") = 0 Then
        If Not StrComp(Email, "mateoalvarezlogan@gmail.com") = 0 Then
            NameCheckReserve = False
            Exit Function

        End If

    End If
        
    If StrComp(UserName, "ARAGON") = 0 Then
        If Not StrComp(Email, "ferminzeta@hotmail.com") = 0 Then
            NameCheckReserve = False
            Exit Function

        End If

    End If
    
    If StrComp(UserName, "MELKOR") = 0 Then
        If Not StrComp(Email, "montiel.marcoseze@gmail.com") = 0 Then
            NameCheckReserve = False
            Exit Function

        End If

    End If
    
    If StrComp(UserName, "ELENTARI") = 0 Then
        If Not StrComp(Email, "angelesechevarrieta53@gmail.com") = 0 Then
            NameCheckReserve = False
            Exit Function

        End If

    End If
        
    If StrComp(UserName, "SELENE") = 0 Then
        If Not StrComp(Email, "arcoiris_4577@hotmail.com") = 0 Then
            NameCheckReserve = False
            Exit Function

        End If

    End If
        
    If StrComp(UserName, "LITO") = 0 Then
        If Not StrComp(Email, "marinolauta@gmail.com") = 0 Then
            NameCheckReserve = False
            Exit Function

        End If

    End If
    
    If StrComp(UserName, "WISTERIA") = 0 Then
        If Not StrComp(Email, "marinolauta@gmail.com") = 0 Then
            NameCheckReserve = False
            Exit Function

        End If

    End If
    
    If StrComp(UserName, "DUKA") = 0 Then
        If Not StrComp(Email, "marinolauta@gmail.com") = 0 Then
            NameCheckReserve = False
            Exit Function

        End If

    End If
    
    NameCheckReserve = True

End Function

Private Sub LoginAccount_SetChar(ByVal UserIndex As Integer, _
                                 ByVal UserName As String, _
                                 ByVal Slot As Byte, _
                                 ByVal CharsAmount As Integer, _
                                 Optional ByVal KillChar As Boolean = False)

    '<EhHeader>
    On Error GoTo LoginAccount_SetChar_Err

    '</EhHeader>
    
    UserList(UserIndex).Account.CharsAmount = UserList(UserIndex).Account.CharsAmount + CharsAmount
    
    Call WriteVar(AccountPath & UserList(UserIndex).Account.Email & ACCOUNT_FORMAT, "INIT", "CHARSAMOUNT", CStr(UserList(UserIndex).Account.CharsAmount))
    
    If Not KillChar Then
        UserList(UserIndex).Account.Chars(Slot).Name = UserName

        Call WriteVar(CharPath & UCase$(UserName) & ".chr", "INIT", "ACCOUNTSLOT", CStr(Slot))
        Call WriteVar(CharPath & UCase$(UserName) & ".chr", "INIT", "ACCOUNTNAME", CStr(UserList(UserIndex).Account.Email))
        Call WriteVar(AccountPath & UserList(UserIndex).Account.Email & ACCOUNT_FORMAT, "CHARS", Slot, UserName)
    Else
        
        Dim NullChar As tAccountChar

        UserList(UserIndex).Account.Chars(Slot) = NullChar
        Call WriteVar(AccountPath & UserList(UserIndex).Account.Email & ACCOUNT_FORMAT, "CHARS", Slot, vbNullString)
        Kill (CharPath & UCase$(UserName) & ".chr")
        
    End If
    
    '<EhFooter>
    Exit Sub

LoginAccount_SetChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.LoginAccount_SetChar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub LoginAccount_Remove(ByVal UserIndex As Integer, _
                               ByVal Key As String, _
                               ByVal Slot As Byte)

    '<EhHeader>
    On Error GoTo LoginAccount_Remove_Err

    '</EhHeader>
    
    Dim SlotUserName As Byte

    Dim Elv          As Byte
    
    Dim UserName     As String
    
    If Slot <= 0 Or Slot > ACCOUNT_MAX_CHARS Then
        ' Anti Hacking
        
        Call Protocol.Kick(UserIndex)
        Exit Sub

    End If
    
    If UserList(UserIndex).flags.UserLogged Then
        Call Protocol.Kick(UserIndex)
        ' Anti Hacking:: Chequeo en el cliente
        Exit Sub

    End If
    
    UserName = UserList(UserIndex).Account.Chars(Slot).Name
    
    If UserName = vbNullString Then
        Call Protocol.Kick(UserIndex)
        ' Anti Hacking:: Chequeo en el cliente
        Exit Sub

    End If
    
    If UserList(UserIndex).Account.Key <> Key Then
        Call WriteErrorMsg(UserIndex, "¡La clave pin es incorrecta!")
        
        Exit Sub

    End If
        
    Elv = 29
    
    If val(GetVar(CharPath & UCase$(UserName) & ".chr", "STATS", "ELV")) > Elv Then
        Exit Sub
    
    End If
    
    If val(GetVar(CharPath & UCase$(UserName) & ".chr", "GUILD", "GUILDINDEX")) > 0 Then
        Call WriteErrorMsg(UserIndex, "¡El personaje posee clan!")

        Exit Sub
    
    End If
    
    If val(GetVar(CharPath & UCase$(UserName) & ".chr", "FLAGS", "BAN")) > 0 Then
        Call WriteErrorMsg(UserIndex, "¡El personaje se encuentra baneado!")

        Exit Sub
    
    End If
        
    If Mercader_CheckUsers(UserList(UserIndex).Account.MercaderSlot, UCase$(UserName)) Then
        Call WriteErrorMsg(UserIndex, "¡El personaje se encuentra en una publicación!")
        Exit Sub

    End If
            
    Call LoginAccount_SetChar(UserIndex, UserName, Slot, -1, True)
    Call WriteLoggedRemoveChar(UserIndex, Slot)
    Call WriteLoggedAccount(UserIndex, UserList(UserIndex).Account.Chars)
    Call Logs_Security(eSecurity, eAntiHack, "Borrado del Personaje " & UserName & " en la cuenta: " & UserList(UserIndex).Account.Email & ". IP: " & UserList(UserIndex).IpAddress)
    
    '<EhFooter>
    Exit Sub

LoginAccount_Remove_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.LoginAccount_Remove " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Busca un personaje en la cuenta del usuario.
Public Function Account_Search_Char(ByVal UserIndex As Integer) As Boolean
    
    Dim A As Long
    
    With UserList(UserIndex)

        For A = 1 To ACCOUNT_MAX_CHARS

            If UCase$(.Account.Chars(A).Name) = UCase$(UserList(UserIndex).Name) Then
                Account_Search_Char = True
                Exit Function

            End If

        Next A
    
    End With

End Function

Public Sub DisconnectAccount(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo DisconnectAccount_Err

    '</EhHeader>
    
    Dim N         As Integer

    Dim VbNullSec As tAccountSecurity
        
    With UserList(UserIndex)
        .Account.Sec = VbNullSec
        
        If NumUsers > 0 Then NumUsers = NumUsers - 1

        ' Desconectamos el personaje en el que estamos
        If .flags.UserLogged Then
            Call Quit_AddNew(UserIndex, True)

            ' Call CloseSocket(UserIndex)
        End If
        
        Call SaveDataAccount(UserIndex, .Account.Email, vbNullString)
        Call ResetUserAccount(UserIndex)

    End With
    
    Call MostrarNumUsers
    
    N = FreeFile
    Open LogPath & "Connect.log" For Append Shared As #N
    Print #N, "La IP " & UserList(UserIndex).Account.Sec.IP_Public & " ha salido del juego. UserIndex:" & UserIndex & " " & Time & " " & Date
    Close #N
    
    '<EhFooter>
    Exit Sub

DisconnectAccount_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.DisconnectAccount " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub DisconnectForced(ByVal UserIndex As Integer, _
                            ByVal Account As String, _
                            ByVal Key As String)

    '<EhHeader>
    On Error GoTo DisconnectForced_Err

    '</EhHeader>
   
    Dim tAccount As Integer

    Dim TempKey  As String
    
    tAccount = CheckEmailLogged(Account)
    
    If tAccount > 0 Then
      
        TempKey = GetVar(AccountPath & Account & ACCOUNT_FORMAT, "INIT", "KEY")
        
        If TempKey = Key Then
            Call Protocol.Kick(tAccount)
            Call Protocol.Kick(UserIndex, "La cuenta ha sido deslogeada.")

        End If
        
    Else
        Call Protocol.Kick(UserIndex, "Cuenta inválida o bien no está conectada.")

    End If

    '<EhFooter>
    Exit Sub

DisconnectForced_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.DisconnectForced " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Login_Char_LoadInfo(ByVal UserIndex As Integer, _
                               ByVal Slot As Byte, _
                               ByVal Name As String)

    '<EhHeader>
    On Error GoTo Login_Char_LoadInfo_Err

    '</EhHeader>
    Dim A        As Long

    Dim TempChar As tAccountChar

    Dim MapTemp  As String

    Dim FilePath As String

    Dim Temp     As String
        
    FilePath = CharPath & UCase$(Name) & ".chr"
        
    With UserList(UserIndex).Account
        TempChar = .Chars(Slot)
        TempChar.Name = Name
            
        If (TempChar.Name <> vbNullString) Then
            MapTemp = GetVar(FilePath, "INIT", "POSITION")
            TempChar.Blocked = val(GetVar(FilePath, "FLAGS", "BLOCKED"))
            TempChar.Map = val(ReadField(1, MapTemp, Asc("-")))
            TempChar.posX = val(ReadField(2, MapTemp, Asc("-")))
            TempChar.posY = val(ReadField(3, MapTemp, Asc("-")))
                
            TempChar.Body = val(GetVar(FilePath, "INIT", "BODY"))
            TempChar.Head = val(GetVar(FilePath, "INIT", "HEAD"))
            TempChar.Helm = val(GetVar(FilePath, "INIT", "CASCO"))
            TempChar.Shield = val(GetVar(FilePath, "INIT", "ESCUDO"))
            TempChar.Weapon = val(GetVar(FilePath, "INIT", "ARMA"))
                
            TempChar.Ban = val(GetVar(FilePath, "FLAGS", "BAN"))
            TempChar.Elv = val(GetVar(FilePath, "STATS", "ELV"))
                
            TempChar.Class = val(GetVar(FilePath, "INIT", "CLASE"))
            TempChar.Raze = val(GetVar(FilePath, "INIT", "RAZA"))
            
            TempChar.Faction = val(GetVar(FilePath, "FACTION", "STATUS"))
            TempChar.FactionRange = val(GetVar(FilePath, "FACTION", "RANGE"))
                
            ' Flags Muerto
            If val(GetVar(FilePath, "FLAGS", "MUERTO")) = 1 Then
                If val(GetVar(FilePath, "FLAGS", "NAVEGANDO")) = 1 Then
                    TempChar.Body = iFragataFantasmal
                Else
                    TempChar.Body = iCuerpoMuerto(False)

                End If

            End If
                
            ' Buscamos una faccion
            If TempChar.Faction = 0 Then
                If val(GetVar(FilePath, "REP", "PROMEDIO")) < 0 Then
                    TempChar.Faction = 3
                Else
                    TempChar.Faction = 4

                End If

            End If
                
            Temp = GetVar(FilePath, "GUILD", "GUILDINDEX")
                
            If val(Temp) > 0 Then
                TempChar.Guild = GuildsInfo(val(Temp)).Name

            End If

            If val(GetVar(FilePath, "FLAGS", "NAVEGANDO")) = 1 Then
                TempChar.Head = 0

            End If
                
        End If
            
        .Chars(Slot) = TempChar

    End With

    '<EhFooter>
    Exit Sub

Login_Char_LoadInfo_Err:
    LogError Err.description & vbCrLf & "in Login_Char_LoadInfo " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub Login_InfoAccountChars(ByVal UserIndex As Integer, _
                                  ByVal Slot As Byte, _
                                  ByVal Name As String)

    '<EhHeader>
    On Error GoTo Login_Char_LoadInfo_Err

    '</EhHeader>
    Dim A        As Long

    Dim TempChar As tAccountChar

    Dim MapTemp  As String

    Dim FilePath As String

    Dim Temp     As String
        
    FilePath = CharPath & UCase$(Name) & ".chr"
        
    With UserList(UserIndex)
        TempChar = .Account.Chars(Slot)
        TempChar.Name = Name
            
        If (TempChar.Name <> vbNullString) Then

            TempChar.Map = .Pos.Map
            TempChar.posX = .Pos.X
            TempChar.posY = .Pos.Y
                
            TempChar.Body = .Char.Body
                
            TempChar.Head = .Char.Head
            TempChar.Helm = .Char.CascoAnim
            TempChar.Shield = .Char.ShieldAnim
            TempChar.Weapon = .Char.WeaponAnim
                
            TempChar.Ban = .flags.Ban
            TempChar.Elv = .Stats.Elv
                
            TempChar.Class = .Clase
            TempChar.Raze = .Raza
            
            TempChar.Faction = .Faction.Status
            TempChar.FactionRange = .Faction.Range
                
            ' Flags Muerto
            If .flags.Muerto = 1 Then
                If .flags.Navegando = 1 Then
                    TempChar.Body = iFragataFantasmal
                Else
                    TempChar.Body = iCuerpoMuerto(False)

                End If

            End If
                
            ' Buscamos una faccion
            If .Faction.Status = 0 Then
                If .Reputacion.promedio < 0 Then
                    TempChar.Faction = 3
                Else
                    TempChar.Faction = 4

                End If

            End If
                
            If .GuildIndex > 0 Then
                TempChar.Guild = .GuildIndex

            End If

            If .flags.Navegando = 1 Then
                TempChar.Head = 0

            End If
                
        End If
            
        .Account.Chars(Slot) = TempChar

    End With

    '<EhFooter>
    Exit Sub

Login_Char_LoadInfo_Err:
    LogError Err.description & vbCrLf & "in Login_Char_LoadInfo " & "at line " & Erl

    '</EhFooter>
End Sub

Sub ResetUserAccount(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo ResetUserAccount_Err

    '</EhHeader>

    UserList(UserIndex).Account = NullAccount
    UserList(UserIndex).AccountLogged = False
    
    '<EhFooter>
    Exit Sub

ResetUserAccount_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.ResetUserAccount " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Function SearchFreeChar(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo SearchFreeChar_Err

    '</EhHeader>

    Dim A As Long
    
    With UserList(UserIndex).Account
    
        For A = 1 To ACCOUNT_MAX_CHARS

            If .Chars(A).Name = vbNullString Then
                SearchFreeChar = A

                Exit Function

            End If

        Next A
        
    End With
    
    '<EhFooter>
    Exit Function

SearchFreeChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.SearchFreeChar " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function Account_SearchFreeChar_Offline(ByVal FilePath As String) As Byte

    '<EhHeader>
    On Error GoTo Account_SearchFreeChar_Offline_Err

    '</EhHeader>

    Dim A    As Long

    Dim Temp As String
    
    For A = 1 To ACCOUNT_MAX_CHARS
        Temp = GetVar(FilePath, "CHARS", A)
            
        If Temp = vbNullString Then
            Account_SearchFreeChar_Offline = A

            Exit Function

        End If

    Next A
    
    '<EhFooter>
    Exit Function

Account_SearchFreeChar_Offline_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.Account_SearchFreeChar_Offline " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function CheckUserLogged(ByVal UserName As String) As Boolean

    '<EhHeader>
    On Error GoTo CheckUserLogged_Err

    '</EhHeader>
    Dim i As Long
    
    For i = 1 To LastUser

        If (UCase$(UserList(i).Name) = UserName) Then
            CheckUserLogged = True
            Exit Function

        End If

    Next i
    
    CheckUserLogged = False
    
    '<EhFooter>
    Exit Function

CheckUserLogged_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.CheckUserLogged " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function CheckEmailLogged(ByVal Email As String) As Integer

    '<EhHeader>
    On Error GoTo CheckEmailLogged_Err

    '</EhHeader>
    Dim i As Long
    
    For i = 1 To LastUser

        If (LCase$(UserList(i).Account.Email) = Email) Then
            CheckEmailLogged = i
            Exit Function

        End If

    Next i
    
    '<EhFooter>
    Exit Function

CheckEmailLogged_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.CheckEmailLogged " & "at line " & Erl
        
    '</EhFooter>
End Function

Private Function Email_Is_Valid(ByVal Email As String) As Boolean

    '<EhHeader>
    On Error GoTo Email_Is_Valid_Err

    '</EhHeader>
    
    On Error GoTo ErrHandler
    
    Dim TempDominio As String '@

    Dim TempGmail   As String '.

    Dim Valid(4)    As String

    Dim TempInt     As Long
    
    Dim A           As Long

    Valid(0) = "gmail.com"
    Valid(1) = "outlook.com"
    Valid(2) = "outlook.com.ar"
    Valid(3) = "hotmail.com"
    Valid(4) = "yahoo.com"
    ' Default: marinolauta@gmail.com
    TempDominio = ReadField(2, Email, Asc("@")) ' gmail.com
    TempGmail = ReadField(1, TempDominio, Asc(".")) ' gmail
    
    For A = LBound(Valid) To UBound(Valid)

        If StrComp(Valid(A), TempDominio) = 0 Then
            
            Email_Is_Valid = True
            Exit Function

        End If

    Next A
    
    Exit Function
ErrHandler:
    Email_Is_Valid = False
    
    '<EhFooter>
    Exit Function

Email_Is_Valid_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAccount.Email_Is_Valid " & "at line " & Erl
        
    '</EhFooter>
End Function
