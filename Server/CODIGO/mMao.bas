Attribute VB_Name = "mMao"
' Author @ Lautaro
' Sistema de Mercado Estilo Tierras del Sur
' 03/07/2022 03:20hs-05:31hs | 14:10-19:26

' DISEÑO
' Esto se guarda en la cuenta
'[SALE]
'Last = 5
'1=SLOT-1-LION
'2=SLOT-2-GORO-LION
'3=SLOT-1-GORO
'4=SLOT-1-KIM DOTCOM
'5=SLOT-3-LION-KIM DOTCOM-GORO

'[SALEOFFER]
' (Carga el Slot del Mercado en la oferta)
'Last = 5
'1=SLOT-ACCOUNT-ORO-1-LION
'2=SLOT-ACCOUNT-ORO-2-GORO-LION
'3=SLOT-ACCOUNT-ORO-1-GORO
'4=SLOT-ACCOUNT-ORO-1-KIM DOTCOM
'5=SLOT-ACCOUNT-ORO-3-LION-KIM DOTCOM-GORO

' En el Mercader.DAT
'[INIT]
'Last = 5

'[SALE]
'1=ACCOUNT-ORO-1-LION
'2=ACCOUNT-ORO-2-GORO-LION
'3=ACCOUNT-ORO-1-GORO
'4=ACCOUNT-ORO-1-KIM DOTCOM
'5=ACCOUNT-ORO-3-LION-KIM DOTCOM-GORO

Option Explicit

Public Const MERCADER_MAX_LIST   As Integer = 255 '

Public Const MERCADER_MAX_GLD    As Long = 2000000000 ' 2.000.000.000

Public Const MERCADER_MAX_DSP    As Long = 100000 '100.000

Public Const MERCADER_GLD_SALE   As Long = 1500 ' 3.000 pide de base de Monedas de oro.

Public Const MERCADER_MIN_LVL    As Byte = 15    ' Pide Nivel 15 para poder ser publicado.

Public Const MERCADER_MAX_OFFER  As Byte = 50

Public Const MERCADER_OFFER_TIME As Long = 120000

Public Type tMercaderObj

    ObjIndex As Integer
    Amount As Integer

End Type

Public Type tMercaderCharInfo

    Name As String
    
    Body As Integer
    Head As Integer
    Weapon As Integer
    Shield As Integer
    Helm As Integer
    
    Elv As Byte
    Exp As Long
    Elu As Long
    
    Hp As Integer
    Constitucion As Byte
    
    Class As Byte
    Raze As Byte
    
    Faction As Byte
    FactionRange As Byte
    FragsCiu As Integer
    FragsCri As Integer
    FragsOther As Integer
    
    Gld As Long
    GuildIndex As Integer
    
    Object() As tMercaderObj
    Bank() As tMercaderObj
    Spells() As Byte
    Skills() As Byte

End Type

Public Type tMercaderChar

    Desc As String
    Dsp As Long
    Gld As Long
    Account As String
    Count As Byte
    NameU() As String
    Info() As tMercaderCharInfo

End Type

Public Type tMercader

    Chars As tMercaderChar
    
    LastOffer As Byte
    Offer(1 To MERCADER_MAX_OFFER) As tMercaderChar
    OfferTime(1 To MERCADER_MAX_OFFER) As Long
    Slot As Integer

End Type

Public MercaderList(1 To MERCADER_MAX_LIST) As tMercader

' Path del Mercado.
Private Function FilePath()
    FilePath = DatPath & "mercader.ini"

End Function

' Guardamos la información del Mercado
Private Sub Mercader_Save(ByVal Slot As Byte)

    '<EhHeader>
    On Error GoTo Mercader_Save_Err

    '</EhHeader>
    Dim Manager As clsIniManager
    
    Set Manager = New clsIniManager
        
    If FileExist(FilePath) Then
        Manager.Initialize (FilePath)

    End If
            
    With MercaderList(Slot)
        Call Manager.ChangeValue("SALE" & Slot, "ACCOUNT", .Chars.Account)
        Call Manager.ChangeValue("SALE" & Slot, "GLD", CStr(.Chars.Gld))
        
        Call Manager.ChangeValue("SALE" & Slot, "CHAR", CStr(.Chars.Count))
        
        If .Chars.Count > 0 Then
            Call Manager.ChangeValue("SALE" & Slot, "CHARS", Mercader_Generate_Text_Chars(.Chars.NameU))
        Else
            Call Manager.ChangeValue("SALE" & Slot, "CHARS", vbNullString)

        End If
        
        Manager.DumpFile FilePath

    End With
    
    Set Manager = Nothing
    '<EhFooter>
    Exit Sub

Mercader_Save_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_Save " & "at line " & Erl

    '</EhFooter>
End Sub

' Guardamos la información del Mercado en la Cuenta del Personaje
Private Sub Mercader_SaveUser(ByVal UserIndex As Integer, _
                              ByVal Slot As Integer, _
                              ByRef Mercader As tMercaderChar)

    '<EhHeader>
    On Error GoTo Mercader_SaveUser_Err

    '</EhHeader>
                              
    Dim Manager As clsIniManager

    Dim A       As Long

    Dim FileP   As String
    
    Set Manager = New clsIniManager
    
    FileP = AccountPath & UserList(UserIndex).Account.Email & ".acc"
    
    Call Manager.Initialize(FileP)
    
    Call Manager.ChangeValue("SALE", "LAST", CStr(Slot))
    
    Call Manager.DumpFile(FileP)
    
    Set Manager = Nothing
    
    '<EhFooter>
    Exit Sub

Mercader_SaveUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_SaveUser " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Cargamos la Lista de Mercado
Public Sub Mercader_Load()

    '<EhHeader>
    On Error GoTo Mercader_Load_Err

    '</EhHeader>
    Dim Manager As clsIniManager

    Dim A       As Long, B As Long
    
    Set Manager = New clsIniManager
    
    If FileExist(FilePath, vbArchive) Then
        Manager.Initialize FilePath

    End If
    
    Dim Temp    As String

    Dim TempA() As String
    
    For A = 1 To MERCADER_MAX_LIST

        With MercaderList(A)
            .Chars.Account = Manager.GetValue("SALE" & A, "ACCOUNT")
            .Chars.Gld = val(Manager.GetValue("SALE" & A, "GLD"))
            .Chars.Count = val(Manager.GetValue("SALE" & A, "CHAR"))
            
            If .Chars.Count > 0 Then
                Temp = Manager.GetValue("SALE" & A, "CHARS")
                TempA = Split(Temp, "-")
                
                ReDim .Chars.NameU(1 To .Chars.Count) As String
                ReDim .Chars.Info(1 To .Chars.Count) As tMercaderCharInfo
                      
                For B = 1 To .Chars.Count
                    .Chars.NameU(B) = TempA(B - 1)
                    .Chars.Info(B) = Mercader_SetChar(B, .Chars.NameU(B), 0)
                Next
    
            End If
            
        End With

    Next A
    
    Set Manager = Nothing
    '<EhFooter>
    Exit Sub

Mercader_Load_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_Load " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Busca un slot libre en el mercado (Máximo=MERCADER_MAX_LIST)
Private Function Mercader_FreeSlot(Optional ByVal Premium As Byte = 0) As Integer

    '<EhHeader>
    On Error GoTo Mercader_FreeSlot_Err

    '</EhHeader>
    Dim A As Long
    
    For A = 1 To MERCADER_MAX_LIST

        If MercaderList(A).Chars.Count = 0 Then
            Mercader_FreeSlot = A
            Exit Function

        End If

    Next A

    '<EhFooter>
    Exit Function

Mercader_FreeSlot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_FreeSlot " & "at line " & Erl
        
    '</EhFooter>
End Function

' Chequeo de 'Hack' al postear personajes erroneos.
' Devuelve la lista de nicks cargados de la cuenta.
' Calculo de cuanto le sale en Monedas de oRO automatico. Realizado con Formula de * según cant + precio base
Public Function Mercader_CheckingChar(ByVal UserIndex As Integer, _
                                      ByRef Chars() As Byte, _
                                      ByRef SumaLvls As Long) As Boolean

    '<EhHeader>
    On Error GoTo Mercader_CheckingChar_Err

    '</EhHeader>
    Dim sChars() As String
    
    Dim A        As Long

    Dim UserName As String
    
    Dim TempLvls As Long

    Dim Temp     As Long
    
    For A = LBound(Chars) To UBound(Chars)

        If Chars(A) = 1 Then
            UserName = UCase$(UserList(UserIndex).Account.Chars(A).Name)
                
            If UserName = vbNullString Then

                Exit Function ' No hay nada en ese slot

            End If
            
            Temp = val(GetVar(CharPath & UserName & ".chr", "STATS", "ELV"))
            
            If Temp < MERCADER_MIN_LVL Then

                Exit Function ' No tiene el nivel correspondiente

            End If
            
            TempLvls = TempLvls + Temp
            
            Temp = val(GetVar(CharPath & UserName & ".chr", "FLAGS", "BAN"))
            
            If Temp > 0 Then

                Exit Function ' El personaje está baneado. Cliente message

            End If

        End If

    Next A
    
    SumaLvls = TempLvls
    Mercader_CheckingChar = True
    
    '<EhFooter>
    Exit Function

Mercader_CheckingChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_CheckingChar " & "at line " & Erl
        
    '</EhFooter>
End Function

' Chequeo de 'Hack' al postear personajes erroneos.
Public Function Mercader_CheckingChar_Offer(ByRef Chars() As String, _
                                            ByVal sAccount As String) As Boolean

    '<EhHeader>
    On Error GoTo Mercader_CheckingChar_Offer_Err

    '</EhHeader>
    
    Dim A    As Long

    Dim Temp As String
    
    For A = LBound(Chars) To UBound(Chars)
        
        Temp = GetVar(CharPath & Chars(A) & ".chr", "INIT", "ACCOUNTNAME")
        
        If Not StrComp(Temp, sAccount) = 0 Then

            Exit Function ' No está más en la cuenta

        End If
        
    Next A
    
    Mercader_CheckingChar_Offer = True
    
    '<EhFooter>
    Exit Function

Mercader_CheckingChar_Offer_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_CheckingChar_Offer " & "at line " & Erl
        
    '</EhFooter>
End Function

' Comprueba si la publicación es válida
' FORMULA =  BASE DE ORO * CANTIDAD DE PJS * SUMA DE NIVELES
'
Public Function Mercader_CheckingNew(ByVal UserIndex As Integer, _
                                     ByRef Chars() As Byte, _
                                     ByRef Mercader As tMercaderChar, _
                                     ByRef SaleCost As Long, _
                                     ByVal Blocked As Byte) As Boolean

    '<EhHeader>
    On Error GoTo Mercader_CheckingNew_Err

    '</EhHeader>
    Dim Suma_Lvls As Long

    Dim A         As Long

    Dim tUser     As Integer
    
    If Mercader.Gld > MERCADER_MAX_GLD Or Mercader.Gld < 1 Then

        Exit Function ' No se permite tanto ORO. Mensaje informativo en el Cliente

    End If
    
    If Not Mercader_CheckingChar(UserIndex, Chars, Suma_Lvls) Then Exit Function

    ' No tiene suficiente Oro en cuenta para realizar la publicación.

    If UserList(UserIndex).Account.Premium > 2 Then
        SaleCost = 0
    Else
        SaleCost = MERCADER_GLD_SALE * Suma_Lvls
            
        If UserList(UserIndex).Account.Gld < SaleCost Then
            ' Mensaje en el Cliente
            Exit Function

        End If

    End If
        
    Dim LastChar As Byte
    
    ' Setting Chars String
    For A = LBound(Chars) To UBound(Chars)

        If Chars(A) = 1 Then
            LastChar = LastChar + 1
            ReDim Preserve Mercader.NameU(1 To LastChar) As String
            ReDim Preserve Mercader.Info(1 To LastChar) As tMercaderCharInfo
                  
            Mercader.NameU(LastChar) = UCase$(UserList(UserIndex).Account.Chars(A).Name)
            Mercader.Count = Mercader.Count + 1
            
            If Blocked = 1 Then
                tUser = NameIndex(Mercader.NameU(LastChar))
                
                If tUser > 0 Then
                    Call WriteErrorMsg(tUser, "Tu personaje pasará a estar bloqueado debido a una Publicación/Oferta.")
                    Call WriteDisconnect(tUser)
                    Call FlushBuffer(tUser)
                    Call CloseSocket(tUser)

                End If

            End If
            
            Mercader.Info(LastChar) = Mercader_SetChar(LastChar, Mercader.NameU(LastChar), Blocked)

        End If

    Next A
        
    ' Posible Hack de los PREMIUM, que no creo que paguen para estafar pero por las dudas!
    If LastChar > 1 Then
        If UserList(UserIndex).Account.Premium < 2 Then
            Exit Function

        End If

    End If
    
    Mercader_CheckingNew = True
    '<EhFooter>
    Exit Function

Mercader_CheckingNew_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_CheckingNew " & "at line " & Erl

    '</EhFooter>
End Function

' Nueva combinación de PJS
Public Sub Mercader_AddList(ByVal UserIndex As Integer, _
                            ByRef Chars() As Byte, _
                            ByRef Mercader As tMercaderChar, _
                            ByVal Blocked As Byte)

    '<EhHeader>
    On Error GoTo Mercader_AddList_Err

    '</EhHeader>
    Dim Slot     As Long

    Dim SaleCost As Long

    Dim SlotUser As Long
        
    Slot = Mercader_FreeSlot(UserList(UserIndex).Account.Premium)
    
    If Slot = 0 Then
        Call WriteErrorMsg(UserIndex, "¡No hay más espacio en el Mercado Central!")
    Else

        If UserList(UserIndex).Account.MercaderSlot > 0 Then
            Call WriteErrorMsg(UserIndex, "¡Tienes una publicación en curso! Primero deberás quitarla para poder realizar otra.")
            Exit Sub

        End If
        
        If Mercader_CheckingNew(UserIndex, Chars, Mercader, SaleCost, Blocked) Then
            If SaleCost = 0 And UserList(UserIndex).Account.Premium < 3 Then
                ' Intentó publicar 0 Personajes, no deberia.
                Exit Sub
    
            End If
                    
            MercaderList(Slot).Chars = Mercader
            MercaderList(Slot).Chars.Account = UserList(UserIndex).Account.Email
            UserList(UserIndex).Account.MercaderSlot = Slot
            UserList(UserIndex).Account.Gld = UserList(UserIndex).Account.Gld - SaleCost
                  
            Call Mercader_SaveUser(UserIndex, Slot, Mercader)
            Call Mercader_Save(Slot)
            
            If UserList(UserIndex).Account.Premium > 2 Then
                Call Mercader_MessageDiscord(MercaderList(Slot).Chars.NameU)

            End If
                
            Call WriteErrorMsg(UserIndex, "Publicación Exitosa. Un correo ha sido enviado con la información de la publicación.")
            Call WriteUpdateStatusMAO(UserIndex, 1)
            Call WriteAccountInfo(UserIndex)
            Call Logs_Security(eLog.eGeneral, eMercader, "Cuenta: " & UserList(UserIndex).Account.Email & " con IP: " & UserList(UserIndex).IpAddress & " ha realizado una PUBLICACION. PIDE ORO: " & MercaderList(Slot).Chars.Gld)

        End If

    End If
    
    '<EhFooter>
    Exit Sub

Mercader_AddList_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_AddList " & "at line " & Erl

    '</EhFooter>
End Sub

' Agregamos una Nueva Oferta a la cuenta
' PROTAGONISTA: El comprador
' Formula para aceptar la OFERTA PONER ABAJO: MAXIMO DE PJS POR CUENTA - LOS OCUPADOS ACTUALES - LOS PJS QUE OFERTO YO
Public Sub Mercader_AddOffer(ByVal UserIndex As Integer, _
                             ByRef Chars() As Byte, _
                             ByVal MercaderSlot As Byte, _
                             ByRef Mercader As tMercaderChar, _
                             ByVal Blocked As Byte)

    '<EhHeader>
    On Error GoTo Mercader_AddOffer_Err

    '</EhHeader>

    Dim tUser     As Integer

    Dim FilePath  As String

    Dim SlotOffer As Integer
    
    Dim SaleCost  As Long
    
    With MercaderList(MercaderSlot)

        If .Chars.Gld > Mercader.Gld Or .Chars.Gld > UserList(UserIndex).Account.Gld Then

            Exit Sub ' La publicación dice que pide un mínimo de Oro y el usuario quiere ofrecer menos

        End If
        
        If .LastOffer = MERCADER_MAX_OFFER Then
            Call WriteErrorMsg(UserIndex, "Parece que el usuario ha recibido demasiadas ofertas y debe seleccionar algunas. Pídele que limpie su lista.")
            Exit Sub

        End If
    
        SlotOffer = Mercader_SlotOffer(MercaderSlot, UserList(UserIndex).Account.Email)
    
        If SlotOffer = -1 Then
            Call WriteErrorMsg(UserIndex, "¡Ya has ofrecido a esta publicación!")
            Exit Sub

        End If
        
        If Mercader_CheckingNew(UserIndex, Chars, Mercader, SaleCost, Blocked) Then
            If SaleCost = 0 And Mercader.Gld = 0 Then

                Exit Sub ' No puede ofrecer NADA a la publicación. Está hackeando el sistema

            End If
                  
            .Offer(SlotOffer) = Mercader
            .OfferTime(SlotOffer) = GetTime
            .LastOffer = .LastOffer + 1
            
            Call WriteErrorMsg(UserIndex, "Tu oferta ha sido enviada. ¡Espera prontas noticias del creador de la publicación!")
            
            tUser = CheckEmailLogged(MercaderList(MercaderSlot).Chars.Account)
            
            If tUser > 0 Then
                If UserList(tUser).flags.UserLogged Then
                    Call WriteConsoleMsg(tUser, "Has recibido una nueva oferta por tu publicación. Dirigete a la Boveda para decidir si quieres aceptarla o no.", FontTypeNames.FONTTYPE_INFOGREEN)

                End If

            End If
                
            'Call WriteSendMercaderOffer(MercaderSlot, SlotOffer, MERCADER_OFFER_TIME)
            Call WriteUpdateStatusMAO(UserIndex, 1)
                
            Call Logs_Security(eLog.eGeneral, eMercader, "Cuenta: " & UserList(UserIndex).Account.Email & " con IP: " & UserList(UserIndex).IpAddress & " ha realizado una oferta a " & MercaderList(MercaderSlot).Chars.Account & ". Ofrece ORO: " & .Offer(SlotOffer).Gld)

        End If
       
    End With

    '<EhFooter>
    Exit Sub

Mercader_AddOffer_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_AddOffer " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Busca un nombre de personaje en la publicación de la cuenta
Public Function Mercader_CheckUsers(ByVal Mercader As Integer, _
                                    ByVal UserName As String) As Boolean

    '<EhHeader>
    On Error GoTo Mercader_CheckUsers_Err

    '</EhHeader>
    Dim A As Long

    If Mercader > 0 Then

        With MercaderList(Mercader)
            
            For A = 1 To .Chars.Count
            
                If UCase$(.Chars.NameU(A)) = UserName Then
                    Mercader_CheckUsers = True
                    Exit Function

                End If
                
            Next A
            
        End With
    
    End If

    '<EhFooter>
    Exit Function

Mercader_CheckUsers_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_CheckUsers " & "at line " & Erl
        
    '</EhFooter>
End Function

' Leemos la Oferta seleccionada
' Protagonista: El que publico (vendedor)
Public Sub Mercader_AcceptOffer(ByVal UserIndex As Integer, _
                                ByVal MercaderSlot As Integer, _
                                ByVal SlotOffer As Byte)

    '<EhHeader>
    On Error GoTo Mercader_AcceptOffer_Err

    '</EhHeader>
                                
    Dim tUser             As Integer

    Dim FilePath          As String

    Dim Gld               As Long

    Dim MercaderNull      As tMercader

    Dim Temp              As String

    Dim SumaLvls          As Long

    Dim SlotsDisponibles  As Long

    Dim SlotsDisponiblesB As Long

    Dim TempChars         As Long

    Dim A                 As Long
        
    Dim NullOffer         As tMercaderChar
        
    FilePath = AccountPath & MercaderList(MercaderSlot).Offer(SlotOffer).Account & ".acc"
            
    ' La oferta caducó
    If (GetTime - MercaderList(MercaderSlot).OfferTime(SlotOffer)) >= MERCADER_OFFER_TIME Then
        MercaderList(MercaderSlot).Offer(SlotOffer) = NullOffer
        MercaderList(MercaderSlot).OfferTime(SlotOffer) = 0
        Call WriteErrorMsg(UserIndex, "¡La oferta caducó!")
        Exit Sub

    End If
        
    ' Caso hipotetico
    ' Personaje nro1 = Publica 5 personajes. Tiene un total de 10
    ' Personaje nro2= Ofrece 5 personajes. Tiene un total de 10
    ' Personajenro1 = 10 - 5 +5 = 10
    SlotsDisponibles = (UserList(UserIndex).Account.CharsAmount - MercaderList(MercaderSlot).Chars.Count + MercaderList(MercaderSlot).Offer(SlotOffer).Count)
          
    If SlotsDisponibles > ACCOUNT_MAX_CHARS Then
        Call WriteErrorMsg(UserIndex, "No tienes espacio para recibir la Oferta.")
        Exit Sub

    End If
    
    tUser = CheckEmailLogged(MercaderList(MercaderSlot).Offer(SlotOffer).Account)
    
    ' El pibe de la oferta
    If tUser > 0 Then
        SlotsDisponibles = (UserList(tUser).Account.CharsAmount - MercaderList(MercaderSlot).Offer(SlotOffer).Count) + MercaderList(MercaderSlot).Chars.Count
              
        Gld = UserList(tUser).Account.Gld
    Else
        Gld = val(GetVar(FilePath, "INIT", "GLD"))
        TempChars = val(GetVar(FilePath, "INIT", "CHARSAMOUNT"))
            
        SlotsDisponibles = (TempChars - MercaderList(MercaderSlot).Offer(SlotOffer).Count) + MercaderList(MercaderSlot).Chars.Count
          
    End If
    
    If SlotsDisponibles > ACCOUNT_MAX_CHARS Then ' MercaderList(MercaderSlot).Chars.Count
        Call WriteErrorMsg(UserIndex, "Parece que la persona no tiene espacio para recibir nuevos personajes.")
        Exit Sub

    End If
    
    ' El pibe de la Oferta no tiene el oro que tenia en un principio
    If MercaderList(MercaderSlot).Offer(SlotOffer).Gld > Gld Then
        Call WriteErrorMsg(UserIndex, "El usuario te ha ofrecido oro y luego lo ha utilizado por lo cual no tiene para pagarte. ¡Lamentamos lo sucedido!")
        
        If tUser > 0 Then
            If UserList(tUser).flags.UserLogged = True Then
                Call WriteConsoleMsg(tUser, "Parece ser que no tienes el oro suficiente y la oferta enviada recientemente no puede ser aceptada.", FontTypeNames.FONTTYPE_INFORED)

            End If

        End If
        
        Exit Sub

    End If

    If MercaderList(MercaderSlot).Offer(SlotOffer).Count > 0 Then
        If Not Mercader_CheckingChar_Offer(MercaderList(MercaderSlot).Offer(SlotOffer).NameU, MercaderList(MercaderSlot).Offer(SlotOffer).Account) Then
            Call WriteErrorMsg(UserIndex, "La solicitud ha expirado por alguna razón. Comprueba que la persona siga disponiendo de la Oferta")

            Exit Sub    ' El pibe de la oferta no tiene mas los personajes

        End If

    End If
    
    ' Quitamos el Oro de la cuenta y lo agregamos a la otra
    If MercaderList(MercaderSlot).Offer(SlotOffer).Gld > 0 Then
        If tUser > 0 Then
            UserList(tUser).Account.Gld = UserList(tUser).Account.Gld - MercaderList(MercaderSlot).Offer(SlotOffer).Gld
        Else
            Call WriteVar(FilePath, "INIT", "GLD", CStr(Gld - MercaderList(MercaderSlot).Offer(SlotOffer).Gld))
    
        End If
        
        UserList(UserIndex).Account.Gld = UserList(UserIndex).Account.Gld + MercaderList(MercaderSlot).Offer(SlotOffer).Gld
        Call WriteErrorMsg(UserIndex, "Se ha depositado en tu cuenta algunas Monedas de Oro por una Venta que acaba de ser confirmada.")

    End If
    
    ' Le quitamos los Pjs al flaco de la venta. Está online
    Call Mercader_UpdateCharsAccount(UserIndex, MercaderList(MercaderSlot).Chars.NameU, True)
    
    ' Le metemos los pjs de la oferta en caso de que haya y no sea solo oro
    If MercaderList(MercaderSlot).Offer(SlotOffer).Count > 0 Then

        ' Quitamos los personajes de la oferta de la cuenta
        If tUser > 0 Then
            Call Mercader_UpdateCharsAccount(tUser, MercaderList(MercaderSlot).Offer(SlotOffer).NameU, True)
              
        Else
            Call Mercader_RemoveCharsAccount_Offline(MercaderList(MercaderSlot).Offer(SlotOffer).Account, MercaderList(MercaderSlot).Offer(SlotOffer).NameU, True)

        End If
            
        Call Mercader_UpdateCharsAccount(UserIndex, MercaderList(MercaderSlot).Offer(SlotOffer).NameU, False)
        
    End If
        
    ' Agregamos los personajes que compró
    If tUser > 0 Then
        Call Mercader_UpdateCharsAccount(tUser, MercaderList(MercaderSlot).Chars.NameU, False)
    Else
        Call Mercader_RemoveCharsAccount_Offline(MercaderList(MercaderSlot).Offer(SlotOffer).Account, MercaderList(MercaderSlot).Chars.NameU, False)

    End If
    
    ' Quitamos la publicacion necesaria en la oferta
    If MercaderList(MercaderSlot).Offer(SlotOffer).Count > 0 Then

        For A = 1 To MercaderList(MercaderSlot).Offer(SlotOffer).Count
            Call Mercader_SearchPublications_User(MercaderList(MercaderSlot).Offer(SlotOffer).Account, MercaderList(MercaderSlot).Offer(SlotOffer).NameU(A))
        Next A

    End If
        
    Call Logs_Security(eLog.eGeneral, eMercader, "Cuenta: " & UserList(UserIndex).Account.Email & " con IP: " & UserList(UserIndex).IpAddress & " ha confirmado la oferta de " & MercaderList(MercaderSlot).Offer(SlotOffer).Count & ". por ORO: " & MercaderList(MercaderSlot).Offer(SlotOffer).Gld)
         
    ' Quitamos la publicación de la VENTA
    MercaderList(MercaderSlot) = MercaderNull
        
    'Call WriteSendMercaderOffer(MercaderSlot, SlotOffer, 0)
    Call Mercader_Remove(MercaderSlot, UserList(UserIndex).Account.Email)
        
    Call WriteLoggedAccount(UserIndex, UserList(UserIndex).Account.Chars)
    Call mAccount.SaveDataAccount(UserIndex, UserList(UserIndex).Account.Email, UserList(UserIndex).IpAddress)
    Call Mercader_Save(MercaderSlot)
        
    '<EhFooter>
    Exit Sub

Mercader_AcceptOffer_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_AcceptOffer " & "at line " & Erl

    '</EhFooter>
End Sub

' Removemos los personajes de la Cuenta
Public Sub Mercader_UpdateCharsAccount(ByVal UserIndex As Integer, _
                                       ByRef Chars() As String, _
                                       ByVal Killed As Boolean)

    '<EhHeader>
    On Error GoTo Mercader_UpdateCharsAccount_Err

    '</EhHeader>
    Dim A           As Long, B As Long

    Dim NullChar    As tAccountChar

    Dim tUser       As Integer
        
    Dim CharsAmount As Byte
        
    CharsAmount = UserList(UserIndex).Account.CharsAmount
    ' Mientras no haya completado la cantidad de chars a poner
              
    For B = LBound(Chars) To UBound(Chars)
        For A = 1 To ACCOUNT_MAX_CHARS
            tUser = NameIndex(Chars(B))
            
            If tUser > 0 Then
                Call WriteDisconnect(tUser)
                Call FlushBuffer(tUser)
                Call CloseSocket(tUser)

            End If
            
            If Killed Then
                If StrComp(UCase$(UserList(UserIndex).Account.Chars(A).Name), UCase$(Chars(B))) = 0 Then
                    UserList(UserIndex).Account.Chars(A) = NullChar
                    Call WriteVar(AccountPath & UserList(UserIndex).Account.Email & ".acc", "CHARS", CStr(A), vbNullString)
                    Call WriteVar(CharPath & Chars(B) & ".chr", "FLAGS", "BLOCKED", "0")
                    Call WriteVar(CharPath & Chars(B) & ".chr", "INIT", "ACCOUNTNAME", vbNullString)
                    Call WriteVar(CharPath & Chars(B) & ".chr", "INIT", "ACCOUNTSLOT", "0")
                    CharsAmount = CharsAmount - 1
                    Exit For

                End If

            Else

                If Len(UserList(UserIndex).Account.Chars(A).Name) = 0 Then
                    UserList(UserIndex).Account.Chars(A).Name = Chars(B)
                    Call Login_Char_LoadInfo(UserIndex, A, Chars(B))
                    Call WriteVar(AccountPath & UserList(UserIndex).Account.Email & ".acc", "CHARS", CStr(A), Chars(B))
                    Call WriteVar(CharPath & Chars(B) & ".chr", "INIT", "ACCOUNTNAME", UserList(UserIndex).Account.Email)
                    Call WriteVar(CharPath & Chars(B) & ".chr", "INIT", "ACCOUNTSLOT", CStr(A))
                    CharsAmount = CharsAmount + 1
                    Exit For

                End If
                
            End If
                
        Next A

    Next B

    UserList(UserIndex).Account.CharsAmount = CharsAmount
    
    '<EhFooter>
    Exit Sub

Mercader_UpdateCharsAccount_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_UpdateCharsAccount " & "at line " & Erl

    '</EhFooter>
End Sub

' Removemos los personajes de la Cuenta Offline
Public Sub Mercader_RemoveCharsAccount_Offline(ByVal Account As String, _
                                               ByRef Chars() As String, _
                                               ByVal Killed As Boolean)

    '<EhHeader>
    On Error GoTo Mercader_RemoveCharsAccount_Offline_Err

    '</EhHeader>
    Dim A           As Long, B As Long

    Dim FilePath    As String

    Dim CharsAmount As Byte
        
    FilePath = AccountPath & Account & ".acc"
          
    CharsAmount = val(GetVar(FilePath, "INIT", "CHARSAMOUNT"))

    For B = LBound(Chars) To UBound(Chars)
        For A = 1 To ACCOUNT_MAX_CHARS
                  
            If Killed Then
                If StrComp(UCase$(GetVar(FilePath, "CHARS", A)), Chars(B)) = 0 Then
                    Call WriteVar(FilePath, "CHARS", A, vbNullString)
                    Call WriteVar(CharPath & Chars(B) & ".chr", "FLAGS", "BLOCKED", "0")
                    CharsAmount = CharsAmount - 1
                    Exit For

                End If
                
            Else
                                
                If GetVar(FilePath, "CHARS", A) = vbNullString Then
                    Call WriteVar(FilePath, "CHARS", A, Chars(B))
                    Call WriteVar(CharPath & Chars(B) & ".chr", "INIT", "ACCOUNTNAME", Account)
                    Call WriteVar(CharPath & Chars(B) & ".chr", "INIT", "ACCOUNTSLOT", CStr(A))
                    CharsAmount = CharsAmount + 1
                    Exit For

                End If

            End If

        Next A
    Next B

    Call WriteVar(FilePath, "INIT", "CHARSAMOUNT", CStr(CharsAmount))
    
    '<EhFooter>
    Exit Sub

Mercader_RemoveCharsAccount_Offline_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_RemoveCharsAccount_Offline " & "at line " & Erl

    '</EhFooter>
End Sub

' Limpiamos un Slot del Mercado
Public Sub Mercader_Remove(ByVal Slot As Integer, ByVal Account As String)

    '<EhHeader>
    On Error GoTo Mercader_Remove_Err

    '</EhHeader>
    Dim MercaderNull As tMercader

    Dim Temp         As String

    Dim A            As Long
        
    Dim tUser        As Integer
        
    For A = 1 To MercaderList(Slot).Chars.Count

        If val(GetVar(CharPath & UCase$(MercaderList(Slot).Chars.NameU(A)) & ".chr", "FLAGS", "BLOCKED")) > 0 Then
            Call WriteVar(CharPath & UCase$(MercaderList(Slot).Chars.NameU(A)) & ".chr", "FLAGS", "BLOCKED", "0")

        End If

    Next A
            
    tUser = CheckEmailLogged(Account)
        
    If tUser > 0 Then
        UserList(tUser).Account.MercaderSlot = 0
        Call WriteVar(AccountPath & Account & ".acc", "SALE", "LAST", "0")
    Else
        Call WriteVar(AccountPath & MercaderList(Slot).Chars.Account & ".acc", "SALE", "LAST", "0")

    End If
          
    MercaderList(Slot) = MercaderNull
    Call Mercader_Save(Slot)
    '<EhFooter>
    Exit Sub

Mercader_Remove_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_Remove " & "at line " & Erl

    '</EhFooter>
End Sub

' Buscamos las publicaciones donde el usuario tenga pjs y las sacamos.
Public Sub Mercader_SearchPublications_User(ByVal Account As String, _
                                            ByVal User As String, _
                                            Optional ByVal BanAccount As Boolean = False)

    '<EhHeader>
    On Error GoTo Mercader_SearchPublications_User_Err

    '</EhHeader>

    Dim A            As Long, B As Long, C As Long

    Dim SlotMercader As Integer

    Dim tUser        As Integer
        
    If User <> vbNullString Then
        tUser = NameIndex(User)

    End If
        
    If tUser > 0 Then
        SlotMercader = UserList(tUser).Account.MercaderSlot
    Else
        SlotMercader = val(GetVar(AccountPath & Account & ".acc", "SALE", "LAST"))

    End If
        
    If SlotMercader = 0 Then Exit Sub
    If MercaderList(SlotMercader).Chars.Count = 0 Then Exit Sub ' Ya fue removida, debido a que otro personaje se involucraba

    ' Control de Baneo de Cuenta entera
    If BanAccount Then
        Mercader_Remove SlotMercader, Account
        Exit Sub

    End If
        
    With MercaderList(SlotMercader)
    
        For A = 1 To .Chars.Count

            If StrComp(.Chars.NameU(A), User) = 0 Then
                Mercader_Remove SlotMercader, Account
                      
                Exit Sub

            End If

        Next A
        
    End With

    '<EhFooter>
    Exit Sub

Mercader_SearchPublications_User_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_SearchPublications_User " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Reiniciamos un Slot de Oferta
' Genera la lista de personajes separados con "-"
Public Function Mercader_Generate_Text_Chars(ByRef Users() As String) As String

    '<EhHeader>
    On Error GoTo Mercader_Generate_Text_Chars_Err

    '</EhHeader>
    Dim A    As Long

    Dim Temp As String

    For A = LBound(Users) To UBound(Users)

        If Users(A) <> vbNullString Then
            Temp = Temp & Users(A) & "-"

        End If

    Next A
    
    Temp = Left$(Temp, Len(Temp) - 1)
    
    Mercader_Generate_Text_Chars = Temp
    '<EhFooter>
    Exit Function

Mercader_Generate_Text_Chars_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_Generate_Text_Chars " & "at line " & Erl
        
    '</EhFooter>
End Function

' Genera un Slot Libre para la Oferta
Private Function Mercader_SlotOffer(ByVal MercaderSlot As Integer, _
                                    ByVal Account As String) As Integer

    '<EhHeader>
    On Error GoTo Mercader_SlotOffer_Err

    '</EhHeader>
    Dim A    As Long

    Dim Temp As Integer
    
    For A = 1 To MERCADER_MAX_OFFER

        With MercaderList(MercaderSlot).Offer(A)
            
            If .Account = vbNullString And Temp = 0 Then
                Temp = A

            End If
            
            If StrComp(.Account, Account) = 0 Then
                Mercader_SlotOffer = -1
                Exit Function

            End If

        End With

    Next A
   
    Mercader_SlotOffer = Temp
    '<EhFooter>
    Exit Function

Mercader_SlotOffer_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_SlotOffer " & "at line " & Erl

    '</EhFooter>
End Function

' # Enviamos un mensaje al canal de discord [SOLO TIERS 3]
Public Sub Mercader_MessageDiscord(ByRef Chars() As String)
    
    On Error GoTo ErrHandler
    
    Dim A     As Long

    Dim Users As String

    For A = LBound(Chars) To UBound(Chars)

        If Chars(A) <> vbNullString Then
            Users = Users & Mercader_GenerateText(Chars(A), True)
            
            If A < UBound(Chars) Then
                Users = Users & " | "

            End If

        End If

    Next A
    
    WriteMessageDiscord CHANNEL_MERCADER, "**Nueva publicación:** " & Users
    
    Exit Sub
    
ErrHandler:

End Sub

Private Function Mercader_SetChar(ByVal SlotChar As Byte, _
                                  ByVal Char As String, _
                                  ByVal Blocked As Byte, _
                                  Optional ByVal IsOffer As Boolean = False) As tMercaderCharInfo

    '<EhHeader>
    On Error GoTo Mercader_SetChar_Err

    '</EhHeader>
    
    Dim Manager As clsIniManager

    Set Manager = New clsIniManager
        
    Dim Charfile As String: Charfile = CharPath & Char & ".chr"

    Dim Temp     As tMercaderCharInfo

    Dim promedio As Long

    Dim A        As Long

    Dim ln       As String
        
    Manager.Initialize Charfile

    With Temp
        .Class = val(Manager.GetValue("INIT", "CLASE"))
        .Raze = val(Manager.GetValue("INIT", "RAZA"))
        .Elv = val(Manager.GetValue("STATS", "ELV"))
        .Exp = val(Manager.GetValue("STATS", "EXP"))
        .Elu = val(Manager.GetValue("STATS", "ELU"))
        .Hp = val(Manager.GetValue("STATS", "MAXHP"))
        .Constitucion = val(Manager.GetValue("ATRIBUTOS", "AT" & eAtributos.Constitucion))
        
        .Body = val(Manager.GetValue("INIT", "BODY"))
        .Head = val(Manager.GetValue("INIT", "HEAD"))
        .Weapon = val(Manager.GetValue("INIT", "ARMA"))
        .Helm = val(Manager.GetValue("INIT", "CASCO"))
        .Shield = val(Manager.GetValue("INIT", "ESCUDO"))
        
        .Gld = val(Manager.GetValue("STATS", "GLD"))
        .GuildIndex = val(Manager.GetValue("GUILD", "GUILDINDEX"))
        
        .Faction = val(Manager.GetValue("FACTION", "STATUS"))
        .FactionRange = val(Manager.GetValue("FACTION", "RANGE"))
        
        Dim Tempito As Long
        
        If .Faction = 0 Then
            Tempito = val(Manager.GetValue("REP", "PROMEDIO"))
            
            If Tempito < 0 Then
                .Faction = 4
            Else
                .Faction = 3

            End If
            
        End If
        
        .FragsCri = val(Manager.GetValue("FACTION", "FRAGSCRI"))
        .FragsCiu = val(Manager.GetValue("FACTION", "FRAGSCIU"))
                
        ReDim .Bank(1 To MAX_BANCOINVENTORY_SLOTS) As tMercaderObj
              
        For A = 1 To MAX_BANCOINVENTORY_SLOTS
            ln = Manager.GetValue("BANCOINVENTORY", "OBJ" & A)
            .Bank(A).ObjIndex = CInt(ReadField(1, ln, 45))
            .Bank(A).Amount = CInt(ReadField(2, ln, 45))
        Next A

        ReDim .Object(1 To MAX_INVENTORY_SLOTS) As tMercaderObj

        For A = 1 To MAX_INVENTORY_SLOTS
            ln = Manager.GetValue("INVENTORY", "OBJ" & A)
            .Object(A).ObjIndex = val(ReadField(1, ln, 45))
            .Object(A).Amount = val(ReadField(2, ln, 45))
        Next A
            
        ReDim .Spells(1 To 35) As Byte
                
        For A = 1 To 35
            .Spells(A) = val(Manager.GetValue("HECHIZOS", "H" & A))
        Next A
                
        ReDim .Skills(1 To NUMSKILLS) As Byte
                
        For A = 1 To NUMSKILLS
            .Skills(A) = val(Manager.GetValue("SKILLS", "SK" & A))
        Next A
                
    End With
    
    If Blocked = 1 Then
        Call Manager.ChangeValue("FLAGS", "BLOCKED", "1")

        If IsOffer Then
            Call Manager.ChangeValue("FLAGS", "OFFERTIME", Format$(Now, "dd/mm/yyyy hh:mm:ss"))

        End If
              
        Call Manager.DumpFile(Charfile)

    End If
    
    Mercader_SetChar = Temp
    Set Manager = Nothing
    '<EhFooter>
    Exit Function

Mercader_SetChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_SetChar " & "at line " & Erl & " NICK: " & Char
        
    '</EhFooter>
End Function

Public Function Mercader_GenerateText(ByVal Char As String, _
                                      Optional ByVal IsDiscord As Boolean = False) As String

    '<EhHeader>
    On Error GoTo Mercader_GenerateText_Err

    '</EhHeader>
    
    Dim Reader As clsIniManager

    Set Reader = New clsIniManager
    
    Dim Class As eClass

    Dim Raze         As eRaza

    Dim Elv          As Byte

    Dim Exp          As Long

    Dim Elu          As Long

    Dim Ups          As Single

    Dim Penas        As Byte

    Dim Hp           As Integer

    Dim Constitution As Byte
    
    Dim Charfile     As String: Charfile = CharPath & Char & ".chr"

    Dim Text         As String
    
    Dim TextUps      As String
    
    Reader.Initialize Charfile
    
    Class = val(Reader.GetValue("INIT", "CLASE"))
    Raze = val(Reader.GetValue("INIT", "RAZA"))
    Elv = val(Reader.GetValue("STATS", "ELV"))
    Exp = val(Reader.GetValue("STATS", "EXP"))
    Elu = val(Reader.GetValue("STATS", "ELU"))
    Hp = val(Reader.GetValue("STATS", "MAXHP"))
    Constitution = val(Reader.GetValue("ATRIBUTOS", "AT" & eAtributos.Constitucion))
    Ups = Hp - getVidaIdeal(Elv, Class, Constitution)
    
    If Ups > 0 Then
        TextUps = "+" & Ups
    ElseIf Ups < 0 Then
        TextUps = Ups
    ElseIf Ups = 0 Then
        TextUps = "Prom"

    End If
        
    If IsDiscord Then
        Text = "**" & UCase$(Char) & "**." & ListaClases(Class) & "." & ListaRazas(Raze) & "." & Elv & "**(" & TextUps & ")**"
    Else
        Text = UCase$(Char) & "." & ListaClases(Class) & "." & ListaRazas(Raze) & "." & Elv & "(" & TextUps & ")"

    End If

    If Elv <> STAT_MAXELV Then
        Text = Text & "(" & Round(CDbl(Exp) * CDbl(100) / CDbl(Elu), 2) & "%)"

    End If
    
    Mercader_GenerateText = Text
    
    Set Reader = Nothing
    
    Exit Function

Mercader_GenerateText_Err:
    Set Reader = Nothing
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_GenerateText " & "at line " & Erl
        
    '</EhFooter>
End Function
