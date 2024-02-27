Attribute VB_Name = "mCuenta"
Option Explicit

' AYUDA
' CREAR NUEVA CUENTA: mCuenta.CreateAccount Account, Passwd, Pin, Email
' BORRAR CUENTA : mCuenta.KillAccount Account, Passwd, Pin, Email
' AGREGAR PERSONAJE A CUENTA: mCuenta.AddCharAccount UserName, Account, Passwd, Pin, Email
' RECUPERAR CONTRASEÑA : mCuenta.RecoverAccount Account, Pin, Email
' BORRAR PERSONAJE : mCuenta.KillCharAccount UserName, Account, Passwd, Pin, Email
' CAMBIAR CONTRASEÑA : mCuenta.ChangePasswdAccount Account, Pin, Email, OldPasswd, NewPasswd

Private Const DIR_ACCOUNT = "\ACCOUNT\"
Private Const FORMAT_ACCOUNT = ".DS"
Public Const MAX_PJS_ACCOUNT = 8

' Lista de Recuperación por tiempo
Private Type tRecover
    Account As String
    Email As String
End Type

Private Recovers(1 To 20) As tRecover
Public Const MAX_OBJ_PREMIUM As Byte = 60

''''''''''''''''''''''''''''''''''''''

Private Type tObj
    ObjIndex As Long
    Amount As Integer
End Type

Private Type tBank
    Login As Byte
    Obj(1 To MAX_OBJ_PREMIUM) As tObj
End Type

' Configuración de la Cuenta
Private Type tCuenta
    UserName(1 To MAX_PJS_ACCOUNT) As String
    Email As String
    Pin As String
    Passwd As String
    Bank As tBank
End Type

'''''''''''''''''''''''''''''''''''''''''


Private Const MAX_AMOUNT_PREMIUM = 100000


' ¿Existe la CUENTA?
Public Function ExistAccount(ByVal Account As String) As Boolean

    If FileExist(App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT) Then
        ExistAccount = True
        Exit Function
    End If
    
End Function

Public Function IsPremiumAccount(ByVal Account As String) As Boolean
    Dim Temp
    Temp = LoadDataAccount(Account, "ACCOUNT", "PREMIUM")
    IsPremiumAccount = IIf((Temp = vbNullString), True, False)
End Function

' Personaje tiene cuenta VIP?
Public Function IsPremium(ByVal UserIndex As Integer) As Boolean
    IsPremium = LoadDataAccount(UserList(UserIndex).Account, "ACCOUNT", "PREMIUM")
    
    If IsPremium Then
        UserList(UserIndex).IsPremium = True
        WriteConsoleMsg UserIndex, "Tu cuenta es Premium. Sigue disfrutando de los beneficios que te brinda DesteriumAO.", FontTypeNames.FONTTYPE_CITIZEN
    Else
        UserList(UserIndex).IsPremium = False
        WriteConsoleMsg UserIndex, "Tu cuenta NO Premium. ¡Enterate de los beneficios!", FontTypeNames.FONTTYPE_CITIZEN
    End If
End Function

' Chequeamos alguna información de la cuenta
Private Function CheckData(ByVal DataUser As String, _
                                ByVal Main As String, _
                                ByVal Account As String) As Boolean
    
    Dim FilePath As String
    Dim Temp As String
    
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    If DataUser = GetVar(FilePath, "ACCOUNT", Main) Then
        CheckData = True
        Exit Function
    End If
End Function

' Obtenemos la contraseña de una Cuenta
Private Function GetPasswd(ByVal Account As String) As String
    Dim FilePath As String
    
    If Not ExistAccount(Account) Then Exit Function
    
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    GetPasswd = GetVar(FilePath, "ACCOUNT", "PASSWD")
End Function

' Guardamos información de la Cuenta
Public Function SaveDataAccount(ByVal Account As String, _
                            ByVal Main As String, _
                            ByVal Var As String, _
                            ByVal Value As String)
    
    Dim FilePath As String
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    WriteVar FilePath, Main, Var, Value
End Function

' Cargamos información de la cuenta
Public Function LoadDataAccount(ByVal Account As String, _
                            ByVal Main As String, _
                            ByVal Var As String) As String
    
    Dim FilePath As String
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    LoadDataAccount = GetVar(FilePath, Main, Var)
End Function

' Buscamos el Slot vacio en la Cuenta
Public Function SearchSlotAccount(ByVal Account As String, ByRef CantPjs As Byte) As Byte
    Dim LoopC As Integer
    Dim Temp As String
    Dim FilePath As String
    Dim Slot As Byte
    
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    For LoopC = 1 To MAX_PJS_ACCOUNT
        If GetVar(FilePath, "ACCOUNT", "PERSONAJE" & LoopC) = "0" Then
            If SearchSlotAccount = 0 Then SearchSlotAccount = LoopC
        Else
            CantPjs = CantPjs + 1
        
        End If
    Next LoopC
End Function
' Buscamos el Personaje
Public Function SearchCharAccount(ByVal Account As String, ByVal UserName As String) As Byte
    Dim LoopC As Integer
    Dim Temp As String
    Dim FilePath As String
    
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    For LoopC = 1 To MAX_PJS_ACCOUNT
        If UCase$(GetVar(FilePath, "ACCOUNT", "PERSONAJE" & LoopC)) = UCase$(UserName) Then
            SearchCharAccount = LoopC
            Exit For
        End If
    Next LoopC
End Function


' Creamos la Cuenta
Public Sub CreateAccount(ByVal UserIndex As Integer, _
                            ByVal Account As String, _
                            ByVal Passwd As String, _
                            ByVal Pin As String, _
                            ByVal Email As String)
                            
    Dim FilePath As String
    Dim LoopC As Integer
    
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    
    If Not AsciiValidos(Account) Then
        Call WriteErrorMsg(UserIndex, "Nombre inválido.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    If ExistAccount(Account) Then
        ' El nombre de la cuenta ya está en uso.
        Exit Sub
    End If
    
    SaveDataAccount Account, "ACCOUNT", "PASSWD", Passwd
    SaveDataAccount Account, "ACCOUNT", "EMAIL", Email
    SaveDataAccount Account, "ACCOUNT", "PIN", Pin
    
    For LoopC = 1 To MAX_PJS_ACCOUNT
        SaveDataAccount Account, "ACCOUNT", "Personaje" & LoopC, "0"
    Next LoopC
    
    For LoopC = 1 To MAX_OBJ_PREMIUM
        SaveDataAccount Account, "BANK", "OBJ" & LoopC, "0-0-0"
    Next LoopC
    
    Protocol.WriteErrorMsg UserIndex, "La cuenta ha sido creada exitosamente"
End Sub

' Logeamos la Cuenta
Public Sub LoginAccount(ByVal UserIndex As Integer, _
                            ByVal Account As String, _
                            ByVal Passwd As String)
                                                 
    Dim FileAccount As String
    Dim FileChar As String
    Dim Chars(1 To MAX_PJS_ACCOUNT) As tCuentaUser
    Dim LoopC As Integer
    
    FileAccount = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    If Not ExistAccount(Account) Or Not CheckData(Passwd, "PASSWD", Account) Then
        WriteErrorMsg UserIndex, "La cuenta no existe o la contraseña es inválida."
        Exit Sub
    End If
    
    For LoopC = 1 To MAX_PJS_ACCOUNT
        With Chars(LoopC)
            .Name = GetVar(FileAccount, "ACCOUNT", "PERSONAJE" & LoopC)
            
            
            If .Name <> "0" Then
                FileChar = CharPath & UCase$(.Name) & ".chr"
                .Ban = val(GetVar(FileChar, "FLAGS", "BAN"))
                .clase = val(GetVar(FileChar, "INIT", "CLASE"))
                .raza = val(GetVar(FileChar, "INIT", "RAZA"))
                .ELV = val(GetVar(FileChar, "STATS", "ELV"))
            Else
                .Ban = 0
                .clase = 0
                .raza = 0
                .ELV = 0
            End If
        End With
    Next LoopC
    
    WriteAccount_Data UserIndex, Chars
End Sub

' Logiamos el Personaje de la cuenta
Public Sub LoginCharAccount(ByVal UserIndex As Integer, _
                                ByVal Account As String, _
                                ByVal Passwd As String, _
                                ByVal UserName As String)
                            
    Dim FilePath As String

    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    If Not ExistAccount(Account) Then Exit Sub
    If Not CheckData(Passwd, "PASSWD", Account) Then Exit Sub
    
    If BANCheck(UserName) Then
        Call WriteErrorMsg(UserIndex, "Se te ha prohibido la entrada a Desterium AO. Baneado por " & ban_Reason(UserName))
        Exit Sub
    End If
    
    If GetVar(CharPath & UserName & ".chr", "INIT", "ACCOUNT") = vbNullString Then
        WriteVar CharPath & UserName & ".chr", "INIT", "ACCOUNT", UCase$(Account)
    End If
    
    If UCase$(Account) <> UCase$(GetVar(CharPath & UserName & ".chr", "INIT", "ACCOUNT")) Then Exit Sub
    
    ConnectUser UserIndex, UserName
    'UserList(UserIndex).Account = Account
End Sub

' Agregamos personajes ya existenes
Public Sub AddTemporal(ByVal UserIndex As Integer, _
                        ByVal Account As String, _
                        ByVal AccountPw As String, _
                        ByVal UserName As String, _
                        ByVal Password As String, _
                        ByVal Email As String, _
                        ByVal Pin As String)

    Dim FilePath As String
    Dim Leer As New clsIniManager
    
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    If Not ExistAccount(Account) Then Exit Sub
    If Not PersonajeExiste(UserName) Then Exit Sub
    If Not CheckData(AccountPw, "PASSWD", Account) Then Exit Sub
    
    If UCase$(Password) <> UCase$(GetVar(CharPath & UserName & ".CHR", "INIT", "PASSWORD")) Then
        WriteErrorMsg UserIndex, "La contraseña no coincide con la del personaje."
        Exit Sub
    End If
    
    If UCase$(Pin) <> UCase$(GetVar(CharPath & UserName & ".CHR", "INIT", "PIN")) Then
        WriteErrorMsg UserIndex, "La Clave PIN no coincide con la del personaje."
        Exit Sub
    End If
    
    If UCase$(Email) <> UCase$(GetVar(CharPath & UserName & ".CHR", "CONTACTO", "EMAIL")) Then
        WriteErrorMsg UserIndex, "El Email no coincide con el del personaje."
        Exit Sub
    End If
    
    If AddCharAccount(Account, UserName) Then
    'JONI DE NIX
    Dim i As Integer
        
            For i = 1 To MAX_GUILDS_DISOLVED
               WriteVar CharPath & UserName & ".CHR", "DISOLVED", "DISOLVED" & i, 0
           Next i
        
        WriteVar CharPath & UserName & ".CHR", "COUNTERS", "TIMETELEP", 0
        WriteVar CharPath & UserName & ".CHR", "INVENTORY", "ANILLONPCSLOT", 0
        WriteVar CharPath & UserName & ".CHR", "STATS", "CANJE", 0
        WriteVar CharPath & UserName & ".CHR", "CONTACTO", "EMAIL", "PERSONAJE AGREGADO A CUENTA " & Account
        
        WriteErrorMsg UserIndex, "Has agregado el personaje CORRECTAMENTE."

        LoginAccount UserIndex, Account, AccountPw
    Else
        WriteErrorMsg UserIndex, "No tienes más espacio."
    End If
    
    
End Sub
' Creamos un Nuevo Personaje
Public Sub CreateCharAccount(ByVal UserIndex As Integer, _
                                ByVal Account As String, _
                                ByVal Passwd As String, _
                                ByVal UserName As String, _
                                ByVal UserClase As Byte, _
                                ByVal UserRaza As Byte, _
                                ByVal UserSexo As Byte)
                                
    Dim FilePath As String
    
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    If Not ExistAccount(Account) Then Exit Sub
    If Not CheckData(Passwd, "PASSWD", Account) Then Exit Sub
    
    If Not AsciiValidos(UserName) Or LenB(UserName) = 0 Then
        WriteErrorMsg UserIndex, "Nombre inválido."
        Exit Sub
    End If
    
    If PersonajeExiste(UserName) Then
        WriteErrorMsg UserIndex, "El personaje ya existe."
        Exit Sub
    End If

    If UserList(UserIndex).flags.UserLogged Then
        Call LogCheating("El usuario " & UserList(UserIndex).Name & " ha intentado crear a " & UserName & " desde la IP " & UserList(UserIndex).ip)
              
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
              
        Exit Sub
    End If
    
    'Tiró los dados antes de llegar acá??
    If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = 0 Then
        Call WriteErrorMsg(UserIndex, "Debe tirar los dados antes de poder crear un personaje.")
        Exit Sub
    End If
    
    
    UserList(UserIndex).Account = Account
    
    ConnectNewUser UserIndex, UserName, UserClase, UserRaza, UserSexo
    AddCharAccount Account, UserName
End Sub

' Borramos la Cuenta ¡NO SE USA!
Public Sub KillAccount(ByVal Account As String, _
                            ByVal Passwd As String, _
                            ByVal Pin As String, _
                            ByVal Email As String)
    
    Dim FilePath As String
    
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    If Not ExistAccount(Account) Then
        ' La cuenta que deseas eliminar no existe.
        Exit Sub
    End If
    
    If Not CheckData(Passwd, "PASSWD", Account) Then Exit Sub
    If Not CheckData(Pin, "PIN", Account) Then Exit Sub
    If Not CheckData(Email, "EMAIL", Account) Then Exit Sub
    
    Kill (App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT)
End Sub

' Agregamos el personaje a nuestra cuenta
Public Function AddCharAccount(ByVal Account As String, _
                            ByVal UserName As String) As Boolean
    Dim Slot As Byte
    Slot = SearchSlotAccount(Account, 1)
    
    If Slot = 0 Then
        AddCharAccount = False
        ' No tienes más espacio para más personajes
        Exit Function
    End If
    
    SaveDataAccount Account, "ACCOUNT", "PERSONAJE" & Slot, UserName
    
    AddCharAccount = True
End Function

' Borramos el personaje de la cuenta
Public Sub KillCharAccount(ByVal UserIndex As Integer, _
                            ByVal Account As String, _
                            ByVal Passwd As String, _
                            ByVal Index As Byte)
    
    Dim FilePath As String
    Dim Temp As String
    Dim UserName As String
    
    UserName = LoadDataAccount(Account, "ACCOUNT", "PERSONAJE" & Index)
    FilePath = App.Path & "\CHARFILE\" & UCase$(UserName) & ".chr"
    
    If Not CheckData(Passwd, "PASSWD", Account) Then Exit Sub
    
    If val(GetVar(FilePath, "FLAGS", "BAN")) = 1 Then
        WriteErrorMsg UserIndex, "No puedes borrar personajes baneados"
        Exit Sub
    End If
    
    SaveDataAccount Account, "ACCOUNT", "PERSONAJE" & Index, "0"
    Kill FilePath
    LoginAccount UserIndex, Account, Passwd
    
End Sub

' Recuperamos la Cuenta
Public Sub RecoverAccount(ByVal UserIndex As Integer, _
                            ByVal Account As String, _
                            ByVal Pin As String, _
                            ByVal Email As String)

    Dim Temp As String
    
    If Not ExistAccount(Account) Then Exit Sub
    If Not CheckData(Pin, "PIN", Account) Then Exit Sub
    If Not CheckData(Email, "EMAIL", Account) Then Exit Sub
    
    'If ExistRecover(UCase$(Account)) Then
        'WriteErrorMsg UserIndex, "Ya hemos recibido tu solicitud. Aguarda unos momentos y te llegará un mail con las indicaciones para entrar al juego."
        'Exit Sub
    'End If
    
    Temp = GeneratePasswd
    SaveDataAccount Account, "ACCOUNT", "PASSWD", Temp
    
    WriteErrorMsg UserIndex, "Cuenta: " & Account & vbCrLf & " Contraseña nueva: " & Temp
    'AddRecoverPasswd Account, Email
End Sub

Public Sub UpdateAccountUserName(ByVal Account As String, _
                                    ByVal UserName As String, _
                                    Optional ByVal NewNick As String = "0")

    If Not ExistAccount(Account) Then Exit Sub
    
    Dim Slot As Byte
    
    Slot = SearchCharAccount(Account, UserName)
    
    SaveDataAccount Account, "ACCOUNT", "PERSONAJE" & Slot, NewNick
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema de recuperación DE CUENTAS
' Cambiamos la contraseña de la Cuenta
Public Sub ChangePasswdAccount(ByVal Account As String, _
                                ByVal Pin As String, _
                                ByVal Email As String, _
                                ByVal OldPasswd As String, _
                                ByVal NewPasswd As String)
    
    If Not CheckData(OldPasswd, "PASSWD", Account) Then Exit Sub
    If Not CheckData(Email, "EMAIL", Account) Then Exit Sub
    If Not CheckData(Pin, "PIN", Account) Then Exit Sub
    
    SaveDataAccount Account, "ACCOUNT", "PASSWD", NewPasswd
    ' Tu contraseña ha sido modificada a NewPasswd
End Sub


Private Function ExistRecover(ByVal Account As String) As Boolean
    Dim LoopC As Long
    
    For LoopC = LBound(Recovers) To UBound(Recovers)
        If Recovers(LoopC).Account = Account Then
            ExistRecover = True
            Exit Function
        End If
    Next LoopC
End Function
Private Function AddRecoverPasswd(ByVal Account As String, ByVal Email As String) As Boolean
    Dim LoopC As Long
    
    For LoopC = LBound(Recovers) To UBound(Recovers)
        If Recovers(LoopC).Account = vbNullString Then
            Recovers(LoopC).Email = Email
            Recovers(LoopC).Account = UCase$(Account)
            AddRecoverPasswd = True
            Exit Function
        End If
    Next LoopC
End Function
Public Sub CheckRecoverPasswd()
    Dim LoopC As Long
    
    For LoopC = LBound(Recovers) To UBound(Recovers)
        With Recovers(LoopC)
            If .Account <> vbNullString Then
                SendMail .Account, .Email
                .Account = vbNullString
                .Email = vbNullString
                Exit Sub
            End If
        End With
    Next LoopC
End Sub

Private Function GeneratePasswd() As String
    Randomize
    GeneratePasswd = Int(Rnd(1) * 10000) & Int(Rnd(1) * 10) & Int(Rnd(1) * 1000)
End Function
Private Function SendMail(ByVal AccountName As String, ByVal Email As String) As Boolean
      
    ' Variable de objeto Cdo.Message
    Dim Obj_Email As CDO.message
            
    ' Crea un Nuevo objeto CDO.Message
    Set Obj_Email = New CDO.message

    ' Indica el servidor Smtp para poder enviar el Mail ( puede ser el nombre _
      del servidor o su dirección IP )
    Obj_Email.Configuration.Fields(cdoSMTPServer) = "mail.desterium.com"
      
    Obj_Email.Configuration.Fields(cdoSendUsingMethod) = 2
      
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CLng(2525)
  
    ' Indica el tipo de autentificación con el servidor de correo _
     El valor 0 no requiere autentificarse, el valor 1 es con autentificación
    Obj_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/" & _
                "configuration/smtpauthenticate") = Abs(True)
      
      
        ' Tiempo máximo de espera en segundos para la conexión
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
  
      
    ' Id de usuario del servidor Smtp ( en el caso de gmail, debe ser la dirección de correro _
     mas el @gmail.com )
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "recoverpasswd@desterium.com"
  
    ' Password de la cuenta
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "p_@zJSB^sGJD"
  
    ' Indica si se usa SSL para el envío. En el caso de Gmail requiere que esté en True
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
      
  
    ' *********************************************************************************
    ' Estructura del mail
    '**********************************************************************************
    ' Dirección del Destinatario
    Obj_Email.To = Email
      
    ' Dirección del remitente
    Obj_Email.from = "recoverpasswd@desterium.com"
      
    ' Asunto del mensaje
    Obj_Email.Subject = "Nueva contraseña de la cuenta '" & AccountName & "'"
      
    ' Cuerpo del mensaje
    Obj_Email.TextBody = "Gracias por recuperar tu cuenta. Recuerda borrar este mensaje una vez anotados los nuevos datos." & _
                         vbCrLf & vbCrLf & _
                         "Cuenta: " & AccountName & vbCrLf & _
                         "Contraseña: " & GetPasswd(AccountName) & vbCrLf & vbCrLf & _
                         "Recuerda que no podrás respondernos mediante este mensaje." & vbCrLf & _
                         "Utiliza nuestro soporte exclusivo ingresando a www.desterium.com" & vbCrLf & _
                         "Atte Staff DSAO"
    

    ' Actualiza los datos antes de enviar
    Obj_Email.Configuration.Fields.Update
      
    On Error Resume Next
    
    ' Envía el email
    Obj_Email.Send
      
      
    If Not Err.Number = 0 Then
       LogError "ERROR EMAIL: Error al recuperar contraseña de " & AccountName & " que tiene el mail " & Email & "."
    End If
      
    ' Descarga la referencia
    If Not Obj_Email Is Nothing Then
        Set Obj_Email = Nothing
    End If
      
    On Error GoTo 0
  
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' FINAL Sistema de recuperación DE CUENTAS

' Agregamos un Objeto al Banco Premium.
Private Function Bank_AddItem(ByVal UserIndex As Integer, _
                                ByVal ObjIndex As Integer, _
                                ByVal Amount As Long, _
                                ByVal LvlItem As Byte) As Boolean
    Dim LoopC As Long
    Dim Temp As Long
    
    With UserList(UserIndex)
        'For LoopC = 1 To MAX_OBJ_PREMIUM
                
            ' Encontramos un Objeto REPETIDO y la cantidad es menor a 100.000
           ' If (.Account_Bank(LoopC).ObjIndex = ObjIndex And _
                (.Account_Bank(LoopC).LvlItem = LvlItem) And _
                (.Account_Bank(LoopC).Amount + Amount) < MAX_AMOUNT_PREMIUM) Then
               ' .Account_Bank(LoopC).Amount = .Account_Bank(LoopC).Amount + Amount
                
                'W 'riteUpdateBankPremium UserIndex, LoopC
                'Bank_AddItem = True
                'Exit Function
           ' End If
            
            ' Encontramos un Slot VACIO para meter nuestro Objeto.
            'If (.Account_Bank(LoopC).ObjIndex = 0 And Bank_AddItem = False) Then
             ''   Bank_AddItem = True
                'Temp = LoopC
            'End If
            
       'Next LoopC
        
        'If Bank_AddItem Then
           ' .Account_Bank(Temp).Amount = Amount
           ' .Account_Bank(Temp).ObjIndex = ObjIndex
           '' .Account_Bank(Temp).LvlItem = LvlItem
            
           ' WriteUpdateBankPremium UserIndex, Temp
       ' End If
    End With
End Function


' Removemos un objeto de la boveda Premium a partir de un Slot y una Cantidad
Public Sub Bank_RemoveItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Amount As Long)
    
    Dim Obj As Obj
    
    With UserList(UserIndex)
        'Obj.ObjIndex = .Account_Bank(Slot).ObjIndex
        'Obj.LvlItem = .Account_Bank(Slot).LvlItem
       ' Obj.Amount = Amount
        
        'If .Account_Bank(Slot).ObjIndex <= 0 Or .Account_Bank(Slot).Amount <= 0 Then
            ' ERROR: HACKEO DE BANCO
        '    Exit Sub
       ' End If
        
        'If (Amount > .Account_Bank(Slot).Amount) Then
            ' ERROR: NO PUEDES QUITAR MÁS DE LO QUE TIENES.
            Exit Sub
        'End If
        
        'If Not MeterItemEnInventario(UserIndex, Obj) Then
            ' ERROR: No tienes espacio en tu inventario
         '   Exit Sub
      '  End If
        
        .Account_Bank(Slot).Amount = .Account_Bank(Slot).Amount - Amount
        
        If .Account_Bank(Slot).Amount = 0 Then
            .Account_Bank(Slot).ObjIndex = 0
           ' .Account_Bank(Slot).LvlItem = 0
        End If
        
      '  WriteUpdateBankPremium UserIndex, Slot
    End With
End Sub


' Depositamos en la Boveda Premium
'Public Sub Bank_DepositeItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Amount As Long)
   ' Dim Obj As Obj
    
    'With UserList(UserIndex)

        
      '  Obj.ObjIndex = .Invent.Object(Slot).ObjIndex
      '  Obj.Amount = Amount
       ' Obj.LvlItem = .Invent.Object(Slot).LvlItem
        
       ' If Not TieneObjetos(Obj.ObjIndex, Obj.Amount, UserIndex) Then
            'ERROR : HACKEO
        '    Exit Sub
       ' End If
        
       ' If Not Bank_AddItem(UserIndex, Obj.ObjIndex, Obj.Amount, Obj.LvlItem) Then
        '    WriteConsoleMsg UserIndex, "No tienes más espacio en tu Boveda Premium.", FontTypeNames.FONTTYPE_WARNING
        '    Exit Sub
       ' End If
        
      '  QuitarObjetos Obj.ObjIndex, Obj.Amount, UserIndex
      '  WriteChangeInventorySlot UserIndex, Slot
        
   ' End With
'End Sub
' Guardamos el Banco Premium
' 1: Cuando se nos cierra el usuario guardamos la información
' 2: Cuando cerramos el comercio Premium guardamos la información
' ¡¡RESETEAMOS INFORMACIÓN SIEMPRE Y SI ES NULA NO LA VOLVEMOS A GUARDAR!!
'Public Sub Bank_Save(ByVal UserIndex As Integer)
 '   Dim LoopC As Integer
    
 '   For LoopC = 1 To MAX_OBJ_PREMIUM
  '      With UserList(UserIndex)
  '          SaveDataAccount .Account, "BANK", "Obj" & LoopC, .Account_Bank(LoopC).ObjIndex & "-" & .Account_Bank(LoopC).Amount & "-" & .Account_Bank(LoopC).LvlItem
  '          .Account_Bank(LoopC).ObjIndex = 0
   ''         .Account_Bank(LoopC).Amount = 0
   '         .Account_Bank(LoopC).LvlItem = 0
  '      End With
  '  Next LoopC
'End Sub

' Cargamos el Banco Premium en el personaje.
'Public Sub Bank_Load(ByVal UserIndex As Integer)
 '   Dim LoopC As Integer
   ' Dim strTemp As String
    
   ' UserList(UserIndex).flags.ComerciandoPremium = True
   ' UserList(UserIndex).flags.Comerciando = True
    
    'For LoopC = 1 To MAX_OBJ_PREMIUM
        'With UserList(UserIndex)
      ''      strTemp = LoadDataAccount(.Account, "BANK", "OBJ" & LoopC)
      '     .Account_Bank(LoopC).ObjIndex = val(ReadField(1, strTemp, Asc("-")))
       '     .Account_Bank(LoopC).Amount = val(ReadField(2, strTemp, Asc("-")))
      '      .Account_Bank(LoopC).LvlItem = val(ReadField(3, strTemp, Asc("-")))
      ''  End With
   ' Next LoopC
'End Sub

'Public Sub Bank_Reset(ByVal UserIndex As Integer)
'
'    Dim LoopC As Long
    
 '''   With UserList(UserIndex)
      '  For LoopC = 1 To MAX_OBJ_PREMIUM
 '           .Account_Bank(LoopC).ObjIndex = 0
 ''           .Account_Bank(LoopC).Amount = 0
'            .Account_Bank(LoopC).LvlItem = 0
    '    Next LoopC
    
 '   End With
'End Sub

'Public Sub Flags_Cuenta_Reset(ByVal UserIndex As Integer)
    
    ' Reseteo de variables de la cuenta y sistemas Premiums relacionados con la misma
    
    'Dim LoopC As Integer
    
    'With UserList(UserIndex)
        'For LoopC = 1 To MAX_OBJ_PREMIUM
            '.Account_Bank(LoopC).Amount = 0
            '.Account_Bank(LoopC).LvlItem = 0
           ' .Account_Bank(LoopC).ObjIndex = 0
'        Next LoopC
        
'        .IsPremium = False
'    End With
'End Sub



