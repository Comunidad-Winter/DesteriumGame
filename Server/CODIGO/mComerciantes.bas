Attribute VB_Name = "mComerciantes"
Option Explicit

' Sistema de comerciantes alquilables por los usuarios. Para que puedan tener sus propios comerciantes y vender sus items. Incluso items de donación al valor que vale el item en el juego.
' ARCHIVO COMERCIANTES.DAT ¡SERVIDOR!

Public Type tCommerceChar

    Char As String

End Type

Public Type tComerciantes

    Owner As String
    OwnerDate As String
    
    ValueGLD As Long
    ValueDSP As Long
    
    RewardGLD As Long           ' ORO que gano vendiendo
    RewardDSP As Long           ' DSP que gano vendiendo
    
    MaxItems As Byte
    Pos As WorldPos
    NpcIndex As Integer
    
    Days As Double
    
    BkItems As Inventario
    
End Type

Public ComerciantesLast As Integer

Public Comerciantes()   As tComerciantes

Public CommerceChar     As tCommerceChar

Public Sub Comerciantes_Load()

    '<EhHeader>
    On Error GoTo Comerciantes_Load_Err

    '</EhHeader>
    
    Dim Manager  As clsIniManager

    Dim FilePath As String

    Dim A        As Long, B As Long

    Dim Temp     As String

    Dim NpcIndex As Integer
    
    Set Manager = New clsIniManager
    
    FilePath = DatPath & "comerciantes.dat"
    
    Manager.Initialize FilePath
    
    ComerciantesLast = val(Manager.GetValue("INIT", "LAST"))
            
    If ComerciantesLast > 0 Then
        ReDim Comerciantes(1 To ComerciantesLast) As tComerciantes
    
        For A = 1 To ComerciantesLast

            With Comerciantes(A)
                .MaxItems = val(Manager.GetValue(A, "MAXITEMS"))
                .ValueDSP = val(Manager.GetValue(A, "VALUEDSP"))
                .ValueGLD = val(Manager.GetValue(A, "VALUEGLD"))
                .NpcIndex = val(Manager.GetValue(A, "NPCINDEX"))
                .Owner = Manager.GetValue(A, "OWNER")
                .OwnerDate = Manager.GetValue(A, "OWNERDATE")
                .Days = val(Manager.GetValue(A, "DAYS"))
            
                Temp = Manager.GetValue(A, "POSITION")
            
                .Pos.Map = val(ReadField(1, Temp, 45))
                .Pos.X = val(ReadField(2, Temp, 45))
                .Pos.Y = val(ReadField(3, Temp, 45))
            
                NpcIndex = SpawnNpc(.NpcIndex, .Pos, False, False)
            
                If NpcIndex = 0 Then
                    Call LogError("ERROR CRITICO EN LA CARGA DE COMERCIANTES")
                    Set Manager = Nothing
                    Exit Sub
                Else
                    .NpcIndex = NpcIndex
                    Npclist(NpcIndex).CommerceIndex = A
                    Npclist(NpcIndex).CommerceChar = .Owner
                
                    ' Le carga el inventario que tenía !
                    If Npclist(NpcIndex).Invent.NroItems > 0 Then
                        Call Manager.ChangeValue(A, "INVENTORY_CANT", CStr(Npclist(NpcIndex).Invent.NroItems))
                
                        For B = 1 To Npclist(NpcIndex).Invent.NroItems
                            Temp = Manager.GetValue(A, "INVENTORY_OBJ" & B)
                        
                            Npclist(NpcIndex).Invent.Object(B).ObjIndex = val(ReadField(1, Temp, 45))
                            Npclist(NpcIndex).Invent.Object(B).Amount = val(ReadField(2, Temp, 45))
                        Next B
                
                    End If

                End If
            
            End With
    
        Next A

    End If
    
    Set Manager = Nothing

    '<EhFooter>
    Exit Sub

Comerciantes_Load_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mComerciantes.Comerciantes_Load " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

' Guarda los comerciantes actualizados con el nombre del dueño del comerciante
Public Sub Comerciantes_Save()

    '<EhHeader>
    On Error GoTo Comerciantes_Save_Err

    '</EhHeader>

    Dim Manager  As clsIniManager

    Dim FilePath As String

    Dim A        As Long

    Dim B        As Long
    
    FilePath = DatPath & "Comerciantes.dat"
    
    Set Manager = New clsIniManager
    
    For A = 1 To ComerciantesLast

        With Comerciantes(A)
            Call Manager.ChangeValue(A, "OWNER", .Owner)
            Call Manager.ChangeValue(A, "OWNERDATE", .OwnerDate)

            ' Guarda lo que haya obtenido
            Call Manager.ChangeValue(A, "REWARDDSP", CStr(.RewardDSP))
            Call Manager.ChangeValue(A, "REWARDGLD", CStr(.RewardGLD))
                  
            ' Guarda todos los objetos que tenga en ese momento.
            If Npclist(.NpcIndex).Invent.NroItems > 0 Then
                Call Manager.ChangeValue(A, "INVENTORY_CANT", CStr(Npclist(.NpcIndex).Invent.NroItems))
                
                For B = 1 To Npclist(.NpcIndex).Invent.NroItems
                    Call Manager.ChangeValue(A, "INVENTORY_OBJ" & B, CStr(Npclist(.NpcIndex).Invent.Object(B).ObjIndex) & "-" & CStr(Npclist(.NpcIndex).Invent.Object(B).Amount))
                Next B
                
            End If

        End With

    Next A
    
    Manager.DumpFile FilePath
    
    Set Manager = Nothing

    '<EhFooter>
    Exit Sub

Comerciantes_Save_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mComerciantes.Comerciantes_Save " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

' Comprueba las fechas de los comerciantes. Si tiene que cancelar le devuelve los objetos al pibe.
' Se comprueba cada un minuto ya que no tiene que ser tan preciso. [CAMBIAR DESDE LA LLAMADA]
Public Sub Comerciantes_Loop()

    '<EhHeader>
    On Error GoTo Comerciantes_Loop_Err

    '</EhHeader>

    Dim A               As Long

    Dim NullComerciante As tComerciantes
    
    For A = 1 To ComerciantesLast

        With Comerciantes(A)

            If .Owner <> vbNullString Then
                If Format(Now, "dd/mm/yyyy") > .OwnerDate Then
                    Call Comerciantes_Return_User(A)

                End If

            End If

        End With

    Next A

    '<EhFooter>
    Exit Sub

Comerciantes_Loop_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mComerciantes.Comerciantes_Loop " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

' Devuelve los objetos a la boveda del personaje.
' En caso de que supere los slots disponibles de bovedas, se pierden los items, pero aun asi se los registramos para no ser tan bruscos
Public Sub Comerciantes_Return_User(ByVal ComercianteIndex As Integer)

    '<EhHeader>
    On Error GoTo Comerciantes_Return_User_Err

    '</EhHeader>

    Dim UserName As String

    Dim tUser    As Integer

    Dim FilePath As String
    
    With Comerciantes(ComercianteIndex)
        UserName = .Owner
        
        tUser = NameIndex(UserName)
        
        If tUser > 0 Then
            Call WriteConsoleMsg(tUser, "Te hemos devuelto los objetos que han quedado y están en tu boveda...", FontTypeNames.FONTTYPE_INFOGREEN)
        
        Else
            FilePath = CharPath & UserName & ".chr"
            
        End If
        
        .Owner = vbNullString
        .OwnerDate = vbNullString
        
        Call Comerciantes_Save

    End With

    '<EhFooter>
    Exit Sub

Comerciantes_Return_User_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mComerciantes.Comerciantes_Return_User " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

' Comprueba que el pibe pueda alquilar la tienda
Public Function Commerce_CanUser_Rent(ByVal UserIndex As Integer, _
                                      ByRef Commerce As tComerciantes) As Boolean

    '<EhHeader>
    On Error GoTo Commerce_CanUser_Rent_Err

    '</EhHeader>

    With UserList(UserIndex)

        If .Stats.Gld < Commerce.ValueGLD Then
            Call WriteConsoleMsg(UserIndex, "¡No tienes suficientes Monedas de Oro para alquilar esta tienda.", FontTypeNames.FONTTYPE_INFORED)
            Exit Function

        End If
    
        If .Stats.Eldhir < Commerce.ValueDSP Then
            Call WriteConsoleMsg(UserIndex, "¡No tienes suficientes Monedas Desterium para alquilar esta tienda. Recuerda tenerlas en tu billetera y no en tu cuenta.", FontTypeNames.FONTTYPE_INFORED)
            Exit Function

        End If
            
        If Commerce.Owner <> vbNullString Then
            Call WriteConsoleMsg(UserIndex, "¡El mercader está alquilado y estará disponible el " & Commerce.OwnerDate & ".", FontTypeNames.FONTTYPE_INFORED)
            
            Exit Function

        End If

    End With
    
    Commerce_CanUser_Rent = True

    '<EhFooter>
    Exit Function

Commerce_CanUser_Rent_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mComerciantes.Commerce_CanUser_Rent " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Function

Public Sub Commerce_ViewBalance(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    
    Dim Temp As tComerciantes

    Temp = Comerciantes(Npclist(NpcIndex).CommerceIndex)
    
    Call WriteChatOverHead(UserIndex, "He hecho " & Temp.RewardDSP & " DSP y " & Temp.RewardGLD & " Monedas de Oro", Npclist(NpcIndex).Char.charindex, vbCyan)
    Call WriteConsoleMsg(UserIndex, "He hecho " & Temp.RewardDSP & " DSP y " & Temp.RewardGLD & " Monedas de Oro", FontTypeNames.FONTTYPE_INFOGREEN)

End Sub

Public Sub Commerce_ReclamarGanancias(ByVal NpcIndex As Integer, _
                                      ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo Commerce_ReclamarGanancias_Err

    '</EhHeader>
    
    Dim CI   As Integer
    
    Dim Temp As tComerciantes

    CI = Npclist(NpcIndex).CommerceIndex
    Temp = Comerciantes(CI)
    
    Call WriteChatOverHead(UserIndex, "He hecho " & Temp.RewardDSP & " DSP y " & Temp.RewardGLD & " Monedas de Oro Y ¡Has retirado TODO!", Npclist(NpcIndex).Char.charindex, vbCyan)
    Call WriteConsoleMsg(UserIndex, "He hecho " & Temp.RewardDSP & " DSP y " & Temp.RewardGLD & " Monedas de Oro Y ¡Has retirado TODO!", FontTypeNames.FONTTYPE_INFOGREEN)
    
    With UserList(UserIndex)
        .Stats.Gld = .Stats.Gld + Temp.RewardGLD
        .Stats.Eldhir = .Stats.Eldhir + Temp.RewardDSP
        
        Call WriteUpdateUserStats(UserIndex)

    End With
    
    With Comerciantes(CI)
        .RewardDSP = 0
        .RewardGLD = 0
    
    End With

    '<EhFooter>
    Exit Sub

Commerce_ReclamarGanancias_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mComerciantes.Commerce_ReclamarGanancias " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

' El comerciante alquila un nuevo comerciante
Public Sub Commerce_SetNew(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    
    Dim Slot As Byte

    Dim Temp As tComerciantes
    
    Slot = Npclist(NpcIndex).CommerceIndex
    Temp = Comerciantes(Slot)
    
    If Not Commerce_CanUser_Rent(UserIndex, Temp) Then Exit Sub
    
    With Comerciantes(Slot)
        .Owner = UCase$(UserList(UserIndex).Name)
        .OwnerDate = DateAdd("d", .Days, Now)
    
        Npclist(NpcIndex).CommerceChar = .Owner

    End With
    
    With UserList(UserIndex)
        .Stats.Gld = .Stats.Gld - Temp.ValueGLD
        .Stats.Eldhir = .Stats.Eldhir - Temp.ValueDSP

    End With
    
    Call WriteConsoleMsg(UserIndex, "¡El cielo no tiene limites! Has alquilado el mercado hasta el día " & Comerciantes(Slot).OwnerDate & ". Esperamos que puedas vender todos tus objetos", FontTypeNames.FONTTYPE_USERPREMIUM)
    Call Logs_Security(eLog.eSecurity, eLogSecurity.eComerciantes, "Alquiler del comerciante " & Npclist(NpcIndex).Name & ".")
    Call WriteUpdateUserStats(UserIndex)
    
End Sub

' DESHABILITADO
Public Sub Commerce_Can_AddItem(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer)
    
    Dim CommerceIndex As Integer
    
    CommerceIndex = Npclist(NpcIndex).CommerceIndex
    
    With Comerciantes(CommerceIndex)

        If Npclist(NpcIndex).Invent.NroItems = .MaxItems Then
            'Call writeconsolemsg(Userindex,"¡Este comerciante admite solo " & A & " espacios para vender."
        
        End If
    
    End With

End Sub
