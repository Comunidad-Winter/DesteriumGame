Attribute VB_Name = "mShop"
Option Explicit

Public Const MAX_TRANSACCION As Byte = 100

Public Type tShopWaiting

    Email As String
    Promotion As Byte
    Bank As String

End Type

Public Type tShop

    Name As String
    Gld As Long
    Dsp As Long
    Desc As String
    ObjIndex As Integer
    ObjAmount As Integer
    Points As Integer
    
End Type

Public Shop()                       As tShop

Public ShopWaiting(MAX_TRANSACCION) As tShopWaiting

Public ShopLast                     As Integer

Public Type tShopChars

    Name As String
    Elv As Byte
    Constitucion As Byte
    Class As Byte
    Raze As Byte
    Head As Integer
    Hp As Integer
    Man As Integer
    Dsp As Integer
    Porc As Byte

End Type

Public ShopCharLast As Integer

Public ShopChars()  As tShopChars

Public Sub Shop_Load_Chars_Index(ByRef Char As tShopChars)
    
    Dim ManagerChar As clsIniManager, FilePath_Char As String

    Set ManagerChar = New clsIniManager
    
    Dim Exp As Long, Elu As Long
    
    With Char
        Set ManagerChar = New clsIniManager
        FilePath_Char = CharPath & .Name & ".chr"
            
        ManagerChar.Initialize FilePath_Char
                
        .Elv = val(ManagerChar.GetValue("STATS", "ELV"))
                  
        If .Elv <> STAT_MAXELV Then
            Exp = val(ManagerChar.GetValue("STATS", "EXP"))
            Elu = val(ManagerChar.GetValue("STATS", "ELU"))
            .Porc = Int(Exp) * CDbl(100) / CDbl(Elu)

        End If
        
        .Head = val(ManagerChar.GetValue("INIT", "HEAD"))
        .Class = val(ManagerChar.GetValue("INIT", "CLASE"))
        .Raze = val(ManagerChar.GetValue("INIT", "RAZA"))
        .Hp = val(ManagerChar.GetValue("STATS", "MAXHP"))
        .Man = val(ManagerChar.GetValue("STATS", "MAXMAN"))
        
        Set ManagerChar = Nothing

    End With

End Sub

Public Sub Shop_Load_Chars()

    '<EhHeader>
    On Error GoTo Shop_Load_Chars_Err

    '</EhHeader>
    
    Dim A        As Long

    Dim FilePath As String, FilePath_Char As String

    Dim Manager  As clsIniManager, ManagerChar As clsIniManager

    Dim Temp     As String

    Dim Exp      As Long, Elu As Long
        
    FilePath = DatPath & "CHARS.ini"
    
    Set Manager = New clsIniManager

    Manager.Initialize FilePath
    
    ShopCharLast = val(Manager.GetValue("INIT", "LAST"))

    ReDim ShopChars(0 To ShopCharLast) As tShopChars
    
    For A = 1 To ShopCharLast

        With ShopChars(A)
            Temp = Manager.GetValue("CHARS", A)
            .Name = ReadField(1, Temp, 45)
            .Dsp = val(ReadField(2, Temp, 45))
                  
            Set ManagerChar = New clsIniManager
            FilePath_Char = CharPath & .Name & ".chr"
            
            ManagerChar.Initialize FilePath_Char
                
            .Elv = val(ManagerChar.GetValue("STATS", "ELV"))
                  
            If .Elv <> STAT_MAXELV Then
                Exp = val(ManagerChar.GetValue("STATS", "EXP"))
                Elu = val(ManagerChar.GetValue("STATS", "ELU"))
                .Porc = Int(Exp) * CDbl(100) / CDbl(Elu)
            Else
                .Porc = A

            End If
                  
            .Head = val(ManagerChar.GetValue("INIT", "HEAD"))
            .Class = val(ManagerChar.GetValue("INIT", "CLASE"))
            .Raze = val(ManagerChar.GetValue("INIT", "RAZA"))
            .Hp = val(ManagerChar.GetValue("STATS", "MAXHP"))
            .Man = val(ManagerChar.GetValue("STATS", "MAXMAN"))
            
            Set ManagerChar = Nothing

        End With
    
    Next A

    Set Manager = Nothing
    
    '<EhFooter>
    Exit Sub

Shop_Load_Chars_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mShop.Shop_Load_Chars " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub SaveShopChars()

    '<EhHeader>
    On Error GoTo Shop_Save_Char_Err

    '</EhHeader>
    
    Dim FilePath As String

    Dim A        As Long
        
    Dim Manager  As clsIniManager

    FilePath = DatPath & "CHARS.ini"
    
    Set Manager = New clsIniManager
        
    Call Manager.ChangeValue("INIT", "LAST", CStr(ShopCharLast))
          
    For A = 1 To ShopCharLast

        With ShopChars(A)
            Call Manager.ChangeValue("CHARS", CStr(A), .Name & "-" & CStr(.Dsp))
    
        End With

    Next A
        
    Manager.DumpFile FilePath
    
    Set Manager = Nothing
    
    '<EhFooter>
    Exit Sub

Shop_Save_Char_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mShop.Shop_Save_Char " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub Shop_Load()

    '<EhHeader>
    On Error GoTo Shop_Load_Err

    '</EhHeader>
    Dim A        As Long

    Dim FilePath As String

    Dim Manager  As clsIniManager

    Dim Temp     As String
    
    FilePath = DatPath & "SHOP.ini"
    
    Set Manager = New clsIniManager
    
    Manager.Initialize FilePath
    
    ShopLast = val(Manager.GetValue("INIT", "LAST"))

    ReDim Shop(1 To ShopLast) As tShop
    
    For A = 1 To ShopLast

        With Shop(A)
            .Name = Manager.GetValue(A, "NAME")
            .Desc = Manager.GetValue(A, "DESC")
            .Gld = val(Manager.GetValue(A, "GLD"))
            .Dsp = val(Manager.GetValue(A, "DSP"))
            
            Temp = Manager.GetValue(A, "OBJINDEX")
            .ObjIndex = val(ReadField(1, Temp, 45))
            .ObjAmount = val(ReadField(2, Temp, 45))
            
            .Points = val(Manager.GetValue(A, "POINTS"))

        End With
    
    Next A
    
    Set Manager = Nothing
    
    Call DataServer_Generate_Shop
    '<EhFooter>
    Exit Sub

Shop_Load_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mShop.Shop_Load " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Transacciones que el Admin debe habilitar

Private Function Transaccion_FreeSlot() As Integer

    '<EhHeader>
    On Error GoTo Transaccion_FreeSlot_Err

    '</EhHeader>

    Dim A As Long
    
    Transaccion_FreeSlot = -1
        
    For A = 0 To MAX_TRANSACCION

        If ShopWaiting(A).Email = vbNullString Then
            Transaccion_FreeSlot = A
            Exit Function

        End If

    Next A

    '<EhFooter>
    Exit Function

Transaccion_FreeSlot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mShop.Transaccion_FreeSlot " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub Transaccion_Add(ByVal UserIndex As Integer, ByRef Waiting As tShopWaiting)

    '<EhHeader>
    On Error GoTo Transaccion_Add_Err

    '</EhHeader>
    
    Dim Slot As Integer
    
    Slot = Transaccion_FreeSlot
    
    If Slot = -1 Then
        Call WriteErrorMsg(UserIndex, "Ocurrió un error grave con la transacción. Espera unos momentos y vuelve a intentar.")
        Exit Sub

    End If
        
    'If Not AsciiValidos(Waiting.Bank) Then
    'Call WriteErrorMsg(UserIndex, "Evita números, simbolos y tildes. Escribe el nombre de la persona (la que figura en la cuenta que ingreso el pago)")
    'Exit Sub
    ' End If
        
    If Not CheckMailString(Waiting.Email) Then
        Call WriteErrorMsg(UserIndex, "Email inválido. Corrobora que no tenga espacios ni caracteres inválidos.")
        Exit Sub

    End If

    If Waiting.Promotion < 0 Or Waiting.Promotion > 5 Then Exit Sub
        
    ShopWaiting(Slot) = Waiting
    FrmShop.lstShop.AddItem Slot & "|" & ShopWaiting(Slot).Email & "|" & ShopWaiting(Slot).Promotion
          
    Call WriteErrorMsg(UserIndex, "¡Has confirmado una nueva transacción. Te pedimos que si aún no has enviado el dinero, lo hagas de manera inmediata así la validación es en los próximos minutos...")
    '<EhFooter>
    Exit Sub

Transaccion_Add_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mShop.Transaccion_Add " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Transaccion_Accept(ByVal Slot As Byte)

    '<EhHeader>
    On Error GoTo Transaccion_Accept_Err

    '</EhHeader>
    
    ' // Le damos lo que pagó
    Dim tUser    As Integer

    Dim FilePath As String

    Dim TempDSP  As Long
        
    tUser = CheckEmailLogged(ShopWaiting(Slot).Email)
        
    If tUser > 0 Then
            
        With UserList(tUser)
            .Account.Eldhir = .Account.Eldhir + Cantidad_Dsp(ShopWaiting(Slot).Promotion)
                
            Call WriteConsoleMsg(tUser, "¡Has recibido la suma de " & Cantidad_Dsp(ShopWaiting(Slot).Promotion) & " DSP. ¡Disfrutalas!", FontTypeNames.FONTTYPE_INFOGREEN)
            Call WriteAccountInfo(tUser)

        End With

    Else
        FilePath = AccountPath & ShopWaiting(Slot).Email & ".acc"
            
        If Not FileExist(FilePath, vbArchive) Then
            Call MsgBox("¡No existe la cuenta!")
            Exit Sub

        End If
            
        TempDSP = val(GetVar(FilePath, "INIT", "ELDHIR"))
            
        Call WriteVar(FilePath, "INIT", "ELDHIR", CStr(TempDSP + Cantidad_Dsp(ShopWaiting(Slot).Promotion)))

    End If
        
    ' // Generamos LOG
    Call Logs_Security(eLog.eGeneral, eLogSecurity.eShop, "Carga de DSP: " & Cantidad_Dsp(ShopWaiting(Slot).Promotion) & " DSP en la cuenta de " & ShopWaiting(Slot).Email & ".")
    
    ' Quitamos de la lista
    Call Transaccion_Clear(Slot)

    FrmShop.lblRef.Caption = "Carga de DSP: " & Cantidad_Dsp(ShopWaiting(Slot).Promotion) & " DSP en la cuenta de " & ShopWaiting(Slot).Email & "."
    '<EhFooter>
    Exit Sub

Transaccion_Accept_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mShop.Transaccion_Accept " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Transaccion_Clear(ByVal Slot As Byte)

    '<EhHeader>
    On Error GoTo Transaccion_Clear_Err

    '</EhHeader>
    
    Dim NullShop As tShopWaiting
    
    ShopWaiting(Slot) = NullShop
    
    FrmShop.lstShop.RemoveItem FrmShop.lstShop.ListIndex
    
    '<EhFooter>
    Exit Sub

Transaccion_Clear_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mShop.Transaccion_Clear " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Function Cantidad_Dsp(ByVal Promotion As Byte) As Long

    Select Case Promotion

        Case 0
            Cantidad_Dsp = 250

        Case 1
            Cantidad_Dsp = 500

        Case 2
            Cantidad_Dsp = 1000

        Case 3
            Cantidad_Dsp = 2000
            
        Case 4
            Cantidad_Dsp = 4000
            
        Case 5
            Cantidad_Dsp = 8000
            
        Case Else
            Cantidad_Dsp = 0
            
    End Select

End Function

Public Function ApplyDiscount(ByVal UserIndex As Integer, ByVal Price As Long)

    '<EhHeader>
    On Error GoTo ApplyDiscount_Err

    '</EhHeader>
    Select Case UserList(UserIndex).Account.Premium
    
        Case 0
            ApplyDiscount = Price

        Case 1
            ApplyDiscount = Price - (Price * 0.05)

        Case 2
            ApplyDiscount = Price - (Price * 0.07)

        Case 3
            ApplyDiscount = Price - (Price * 0.1)

    End Select

    '<EhFooter>
    Exit Function

ApplyDiscount_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mShop.ApplyDiscount " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Function

Public Sub ConfirmItem(ByVal UserIndex As Integer, _
                       ByVal ID As Integer, _
                       ByVal SelectedValue As Byte)

    '<EhHeader>
    On Error GoTo ConfirmItem_Err

    '</EhHeader>
    
    Dim Obj    As Obj

    Dim Random As Byte

    Dim Sound  As Integer
    
    With Shop(ID)
            
        Obj.Amount = .ObjAmount
        Obj.ObjIndex = .ObjIndex
            
        If Obj.ObjIndex <> 880 Then
            If SelectedValue = 0 And .Gld = 0 Then Exit Sub ' Quiere comprar por ORO, y el item no sale ORO
            If SelectedValue = 1 And .Dsp = 0 Then Exit Sub ' Quiere comprar por DSP, y el item no sale DSP
              
            If (.Gld > UserList(UserIndex).Account.Gld) And SelectedValue = 0 Then Exit Sub
            If (.Dsp > UserList(UserIndex).Account.Eldhir) And SelectedValue = 1 Then Exit Sub

        End If

        If Obj.ObjIndex = 880 Then

            ' Solicita Puntos de Torneo [CANJE]
            If .Points > UserList(UserIndex).Stats.Points Then Exit Sub
            UserList(UserIndex).Account.Eldhir = UserList(UserIndex).Account.Eldhir + Obj.Amount
            UserList(UserIndex).Stats.Points = UserList(UserIndex).Stats.Points - .Points
                  
        ElseIf Obj.ObjIndex = 9999 Then
            Exit Sub
            
        ElseIf Obj.ObjIndex = 9998 Then

            Dim MeditationSelected As Byte

            MeditationSelected = val(ReadField(2, .Name, Asc(" ")))
                    
            UserList(UserIndex).MeditationUser(MeditationSelected) = 1
            Call mMeditations.Meditation_Select(UserIndex, MeditationSelected)
                
            Call Logs_Security(eLog.eGeneral, eLogSecurity.eShop, "Usuario: " & UserList(UserIndex).Name & "» COMPRA DE MEDITACION: " & MeditationSelected)
        Else

            If Not MeterItemEnInventario(UserIndex, Obj, True) Then Exit Sub
                  
            Call Logs_Security(eLog.eGeneral, eLogSecurity.eShop, "Usuario: " & UserList(UserIndex).Name & "» COMPRA DE ITEM: " & ObjData(Obj.ObjIndex).Name & " (x" & Obj.Amount & ")")

        End If
            
        If SelectedValue = 0 Then
            UserList(UserIndex).Account.Gld = UserList(UserIndex).Account.Gld - ApplyDiscount(UserIndex, .Gld)
        ElseIf SelectedValue = 1 Then
            UserList(UserIndex).Account.Eldhir = UserList(UserIndex).Account.Eldhir - ApplyDiscount(UserIndex, .Dsp)

        End If
              
        Random = RandomNumber(1, 100)
        
        If Random <= 25 Then
            Sound = eSound.sChestDrop1
        ElseIf Random <= 50 Then
            Sound = eSound.sChestDrop2
        Else
            Sound = eSound.sChestDrop3

        End If
        
        Call SendData(SendTarget.ToOne, UserIndex, PrepareMessagePlayEffect(Sound, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Call WriteAccountInfo(UserIndex)
    
    End With

    '<EhFooter>
    Exit Sub

ConfirmItem_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mShop.ConfirmItem " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub ConfirmTier(ByVal UserIndex As Integer, ByVal Tier As Byte)

    '<EhHeader>
    On Error GoTo ConfirmTier_Err

    '</EhHeader>
    If Tier <= 0 Or Tier > 3 Then Exit Sub
    
    Dim Price As Long
    
    Select Case Tier
    
        Case 1
            Price = 100

        Case 2
            Price = 250

        Case 3
            Price = 450

    End Select
    
    With UserList(UserIndex)

        If .Account.Eldhir < Price Then Exit Sub

        If .Account.Premium > 0 Then
            If .Account.Premium < Tier Then
                Call WriteConsoleMsg(UserIndex, "¡Debes esperar a que se venca el Tier Inferior o hablar con un administrador para realizar un Upgrade!", FontTypeNames.FONTTYPE_INFORED)
                Call WriteErrorMsg(UserIndex, "¡Debes esperar a que se venca el Tier Inferior o hablar con un administrador para realizar un Upgrade!")
                Exit Sub

            End If
                  
            .Account.DatePremium = DateAdd("m", 1, .Account.DatePremium)

            Call Logs_Security(eLog.eGeneral, eLogSecurity.eShop, "Usuario: " & UserList(UserIndex).Name & "» ACTUALIZA A TIER:  " & Tier)
        Else
            .Account.DatePremium = DateAdd("m", 1, Now)
            
            Call Logs_Security(eLog.eGeneral, eLogSecurity.eShop, "Usuario: " & UserList(UserIndex).Name & "» COMPRA DE TIER:  " & Tier)

        End If
            
        .Account.Eldhir = .Account.Eldhir - Price
                   
        .Account.Premium = Tier
            
        Call WriteErrorMsg(UserIndex, "Tiempo PREMIUM actualizado hasta " & .Account.DatePremium & ".")
        Call WriteConsoleMsg(UserIndex, "Tiempo PREMIUM actualizado hasta " & .Account.DatePremium & ".", FontTypeNames.FONTTYPE_USERPREMIUM)
         
        Call SaveDataAccount(UserIndex, .Account.Email, .IpAddress)
        Call SendData(SendTarget.ToOne, UserIndex, PrepareMessagePlayEffect(eSound.sVictory2, .Pos.X, .Pos.Y))

    End With
          
    Call WriteAccountInfo(UserIndex)
    '<EhFooter>
    Exit Sub

ConfirmTier_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mShop.ConfirmTier " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub ConfirmChar(ByVal UserIndex As Integer, ByVal ID As Byte)

    '<EhHeader>
    On Error GoTo ConfirmChar_Err

    '</EhHeader>
    
    If UserList(UserIndex).Account.Eldhir < ShopChars(ID).Dsp Then Exit Sub
    
    Dim NullShopChar As tShopChars

    Dim Chars(0)     As String

    If ShopChars(ID).Elv = 0 Then
        Call WriteErrorMsg(UserIndex, "¡Parece que el personaje ha sido vendido!")
        Call WriteShopChars(UserIndex)
        Exit Sub

    End If

    If (UserList(UserIndex).Account.CharsAmount) = ACCOUNT_MAX_CHARS Then
        Call WriteErrorMsg(UserIndex, "No tienes espacio para recibir nuevos personajes.")
        Exit Sub

    End If
    
    UserList(UserIndex).Account.Eldhir = UserList(UserIndex).Account.Eldhir - ShopChars(ID).Dsp
    
    Chars(0) = ShopChars(ID).Name
    
    Call Logs_Security(eLog.eGeneral, eLogSecurity.eShop, "Cuenta: " & UserList(UserIndex).Account.Email & "» COMPRA EL PERSONAJE:  " & ShopChars(ID).Name & " a " & ShopChars(ID).Dsp & " DSP")
    
    Call Mercader_UpdateCharsAccount(UserIndex, Chars, False)
    Call UpdateShopChars(ID)
          
    Call mAccount.SaveDataAccount(UserIndex, UserList(UserIndex).Account.Email, UserList(UserIndex).IpAddress)
    Call WriteLoggedAccount(UserIndex, UserList(UserIndex).Account.Chars)

    Call WriteShopChars(UserIndex)
    Call WriteErrorMsg(UserIndex, "¡Has comprado el personaje: " & Chars(0) & "!")
    '<EhFooter>
    Exit Sub

ConfirmChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mShop.ConfirmChar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub UpdateShopChars(ByVal Slot As Byte)

    '<EhHeader>
    On Error GoTo UpdateShopChars_Err

    '</EhHeader>

    Dim A      As Long

    Dim Temp() As tShopChars

    ReDim Temp(1 To ShopCharLast) As tShopChars
            
    ' Copia para no repetir
    For A = 1 To ShopCharLast
        Temp(A) = ShopChars(A)
    Next A
    
    ShopCharLast = ShopCharLast - 1
    
    ' Movemos +1 a los usuarios desde esta posición.
    For A = Slot To ShopCharLast
        ShopChars(A) = Temp(A + 1)
    Next A
    
    Call SaveShopChars

    '<EhFooter>
    Exit Sub

UpdateShopChars_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mShop.UpdateShopChars " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Shop_CharAdd(ByRef Char As tShopChars)

    '<EhHeader>
    On Error GoTo Shop_CharAdd_Err

    '</EhHeader>
    
    ShopCharLast = ShopCharLast + 1
    
    ReDim Preserve ShopChars(0 To ShopCharLast) As tShopChars
          
    Call Shop_Load_Chars_Index(Char)
    ShopChars(ShopCharLast) = Char
    Call SaveShopChars

    '<EhFooter>
    Exit Sub

Shop_CharAdd_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mShop.Shop_CharAdd " & "at line " & Erl
        
    '</EhFooter>
End Sub
