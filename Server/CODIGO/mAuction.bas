Attribute VB_Name = "mAuction"
Option Explicit

Public Const AUCTION_TIME As Long = 300

Public Type tAuction_Security

    IP As String
    Email As String

End Type

Public Type tAuction_Offer

    Name As String
    Gld As Long
    Eldhir As Long
    TimeLastOffer As Integer
    Security As tAuction_Security

End Type

Public Type tAuction

    Name As String
    
    ObjIndex As Integer
    Amount As Integer
    Durabilidad As Integer
    Gld As Long
    Eldhir As Long
    Time As Long
    
    Offer As tAuction_Offer
    
    Security As tAuction_Security

End Type

Public Auction As tAuction

Private Function Auction_Checking_SameUser(ByVal UserIndex As Integer, _
                                           ByVal Email As String, _
                                           ByVal IP As String) As Boolean

    '<EhHeader>
    On Error GoTo Auction_Checking_SameUser_Err

    '</EhHeader>
    
    Dim Account As String
    
    ' Mismo Email
    If StrComp(UserList(UserIndex).Account.Email, Email) = 0 Then
        Exit Function

    End If
    
    ' Misma IP
    If UserList(UserIndex).IpAddress <> vbNullString Then
        If StrComp(UserList(UserIndex).IpAddress, IP) = 0 Then
            Exit Function

        End If

    End If
    
    Auction_Checking_SameUser = True
    '<EhFooter>
    Exit Function

Auction_Checking_SameUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAuction.Auction_Checking_SameUser " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub Auction_CreateNew(ByVal UserIndex As Integer, _
                             ByVal ObjIndex As Integer, _
                             ByVal Amount As Integer, _
                             ByVal Gld As Long, _
                             ByVal Eldhir As Long)

    '<EhHeader>
    On Error GoTo Auction_CreateNew_Err

    '</EhHeader>
                             
    If UserList(UserIndex).Account.Premium < 2 Then
        Call WriteConsoleMsg(UserIndex, "Debes ser al menos Tier 2 para poder subastar objetos. Consulta las promociones en www.argentumgame.com/download", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
        
    If UserList(UserIndex).Pos.Map <> 1 Then
        Call WriteConsoleMsg(UserIndex, "No puedes subastar objetos si no estas en la ciudad principal.", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
    
    If Auction.ObjIndex > 0 Then
        Call WriteConsoleMsg(UserIndex, "Ya hay una subasta en trámite, espera a que termine.", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
    
    If ConfigServer.ModoSubastas = 0 Then
        Call WriteConsoleMsg(UserIndex, "Las subastas no estan permitidas momentaneamente.", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
    
    If Not Auction_ObjValid(ObjIndex) Then
        Call WriteConsoleMsg(UserIndex, "No puedes subastar este objeto.", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
    
    With Auction
        .Name = UCase$(UserList(UserIndex).Name)
        .ObjIndex = ObjIndex
        .Amount = Amount
        .Gld = Gld
        .Eldhir = Eldhir
        .Time = AUCTION_TIME
        
        .Security.Email = UserList(UserIndex).Account.Email
        .Security.IP = UserList(UserIndex).IpAddress
        .Offer.Name = UCase$(UserList(UserIndex).Name)
        .Offer.Gld = Gld
        .Offer.Eldhir = Eldhir

    End With
    
    Dim TempObj As String

    TempObj = ObjData(ObjIndex).Name
    
    If ObjData(ObjIndex).Bronce = 1 Then
        TempObj = TempObj & " [BRONCE]"

    End If
    
    If ObjData(ObjIndex).Plata = 1 Then
        TempObj = TempObj & " [PLATA]"

    End If
    
    If ObjData(ObjIndex).Oro = 1 Then
        TempObj = TempObj & " [ORO]"

    End If
    
    If ObjData(ObjIndex).Premium = 1 Then
        TempObj = TempObj & " [PREMIUM]"

    End If
    
    Call QuitarObjetos(ObjIndex, Amount, UserIndex)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Subasta» Un objeto se está vendiendo al mejor postor y es " & TempObj & " (x" & Amount & ")", FontTypeNames.FONTTYPE_CRITICO))
    Call Logs_Security(eSecurity, eSubastas, "Subasta nueva» El personaje " & Auction.Offer.Name & " puso el objeto " & TempObj & " (x" & Amount & ") a " & Format$(Gld, "#,###") & " Monedas de Oro Y " & Eldhir & " Monedas de Eldhir.")
    '<EhFooter>
    Exit Sub

Auction_CreateNew_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAuction.Auction_CreateNew " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Auction_Offer(ByVal UserIndex As Integer, _
                         ByVal Gld As Long, _
                         ByVal Eldhir As Long)

    '<EhHeader>
    On Error GoTo Auction_Offer_Err

    '</EhHeader>
    
    Dim tUser      As Integer

    Dim FilePath   As String

    Dim TempGld    As Long

    Dim TempEldhir As Long
    
    If ConfigServer.ModoSubastas = 0 Then
        Call WriteConsoleMsg(UserIndex, "Las subastas no estan permitidas momentaneamente.", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
    
    If Auction.ObjIndex = 0 Then
        Call WriteConsoleMsg(UserIndex, "¡No hay ninguna subasta en trámite!", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
    
    If Auction.Name = UCase$(UserList(UserIndex).Name) Then
        Call WriteConsoleMsg(UserIndex, "¡No te ofrezcas a ti mismo!", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
    
    If Not Auction_Checking_SameUser(UserIndex, Auction.Security.Email, Auction.Security.IP) Then
        Call Logs_Security(eSecurity, eAntiHack, "[DOBLE CLIENTE] El personaje " & UCase$(UserList(UserIndex).Name) & " con IP: " & Auction.Security.IP & " Y Email: " & Auction.Security.Email & " intentó ofrecerse a sí mismo y fue advertido.")
        Call WriteConsoleMsg(UserIndex, "¡No te ofrezcas a ti mismo!", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
    
    If Not Auction_Checking_SameUser(UserIndex, Auction.Offer.Security.Email, Auction.Offer.Security.IP) Then
        Call Logs_Security(eSecurity, eAntiHack, "[DOBLE CLIENTE] El personaje " & UCase$(UserList(UserIndex).Name) & " con IP: " & Auction.Offer.Security.IP & " Y Email: " & Auction.Offer.Security.Email & " intentó ofrecerse a sí mismo y fue advertido.")
        Call WriteConsoleMsg(UserIndex, "¡La última oferta la has realizado tu!", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
    
    If Auction.Offer.Name = UCase$(UserList(UserIndex).Name) Then
        Call WriteConsoleMsg(UserIndex, "¡Ya has ofrecido!", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
    
    If UserList(UserIndex).Stats.Gld < Gld Then
        Exit Sub

    End If
    
    If UserList(UserIndex).Stats.Eldhir < Eldhir Then
        Exit Sub

    End If
    
    With Auction.Offer

        If .TimeLastOffer > 0 Then
            Call WriteConsoleMsg(UserIndex, "Debes esperar unos momentos para realizar otra oferta.", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If
    
        If Gld < (.Gld * 1.1) Then
            Call WriteConsoleMsg(UserIndex, "¡Debes ofreer al menos un 10% más que la oferta anterior. En total sería (" & .Gld * 1.1 & "!", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If
    
        ' Le reintegramos el oro a la anterior oferta.
        If .Name <> vbNullString And .Name <> Auction.Name Then
            tUser = NameIndex(.Name)
            
            If tUser > 0 Then
                UserList(tUser).Stats.Gld = UserList(tUser).Stats.Gld + .Gld
                UserList(tUser).Stats.Eldhir = UserList(tUser).Stats.Eldhir + .Eldhir
                
                Call WriteUpdateUserStats(tUser)
                Call WriteConsoleMsg(tUser, "¡Han ofrecido más Monedas que tú! ¿Te darás por vencido?", FontTypeNames.FONTTYPE_INFORED)
            Else
                FilePath = CharPath & .Name & ".chr"
                TempGld = val(GetVar(FilePath, "STATS", "GLD"))
                TempEldhir = val(GetVar(FilePath, "STATS", "ELDHIR"))
                
                Call WriteVar(FilePath, "STATS", "GLD", CStr(TempGld + .Gld))
                Call WriteVar(FilePath, "STATS", "ELDHIR", CStr(TempEldhir + .Eldhir))

            End If

        End If
        
        .Name = UCase$(UserList(UserIndex).Name)
        .Gld = Gld
        .Eldhir = Eldhir
        .TimeLastOffer = 5
        .Security.Email = UserList(UserIndex).Account.Email
        .Security.IP = UserList(UserIndex).IpAddress
        
        UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld - Gld
        UserList(UserIndex).Stats.Eldhir = UserList(UserIndex).Stats.Eldhir - Eldhir
        Call WriteUpdateUserStats(UserIndex)
        
        Call SendData(SendTarget.toMap, 1, PrepareMessageConsoleMsg("Subasta» El personaje " & Auction.Offer.Name & " ha ofrecido: " & Format$(.Gld, "#,###") & " Monedas de Oro Y " & .Eldhir & " Monedas de Eldhir.", FontTypeNames.FONTTYPE_INFOGREEN))
        Call Logs_Security(eSecurity, eSubastas, "El personaje " & Auction.Offer.Name & " ha ofrecido: " & .Gld & " Monedas de Oro Y " & .Eldhir & " Monedas de Eldhir.")

    End With
    
    Auction.Time = Auction.Time + 20
    '<EhFooter>
    Exit Sub

Auction_Offer_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAuction.Auction_Offer " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Auction_Loop()

    '<EhHeader>
    On Error GoTo Auction_Loop_Err

    '</EhHeader>
    With Auction

        If .Offer.TimeLastOffer > 0 Then
            .Offer.TimeLastOffer = .Offer.TimeLastOffer - 1

        End If
        
        If .Time > 0 Then
            .Time = .Time - 1

            If (.Time Mod 60) = 0 Then
                
                If .Time = 0 Then
                    If .Offer.Name = .Name Then
                        Call Auction_RewardObj
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Subasta» La subasta del objeto " & ObjData(.ObjIndex).Name & " (x" & .Amount & ") ha concluído sin ofertas.", FontTypeNames.FONTTYPE_CRITICO))
                        
                    Else
                    
                        Call Auction_RewardGld
                        Call Auction_RewardObj

                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Subasta» La subasta del objeto " & ObjData(.ObjIndex).Name & " (x" & .Amount & ") ha concluído. ¡Se lo ha llevado el personaje " & .Offer.Name & "!", FontTypeNames.FONTTYPE_CRITICO))
                        
                    End If
                    
                    Call Auction_Reset
                ElseIf .Time <> 60 Then

                    If .Offer.Name <> .Name Then
                        Call SendData(SendTarget.toMap, 1, PrepareMessageConsoleMsg("Subasta» " & ObjData(.ObjIndex).Name & " (x" & .Amount & "). La última oferta es de " & Format$(.Offer.Gld, "#,###") & " Monedas de Oro y " & .Offer.Eldhir & " Monedas de Eldhir.", FontTypeNames.FONTTYPE_CRITICO))

                    End If

                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Subasta» Última posibilidad para ofertar por el objeto " & ObjData(.ObjIndex).Name & " (x" & .Amount & "). La última oferta es de " & Format$(.Offer.Gld, "#,###") & " Monedas de Oro y " & .Offer.Eldhir & " Monedas de Eldhir.", FontTypeNames.FONTTYPE_CRITICO))

                End If

            End If

        End If
    
    End With

    '<EhFooter>
    Exit Sub

Auction_Loop_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAuction.Auction_Loop " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub Auction_Object()

    '<EhHeader>
    On Error GoTo Auction_Object_Err

    '</EhHeader>
    Dim A        As Long

    Dim FilePath As String

    Dim Temp     As String
    
    FilePath = CharPath & Auction.Offer.Name & ".chr"
    
    For A = 1 To MAX_BANCOINVENTORY_SLOTS
        Temp = GetVar(FilePath, "BANCOINVENTORY", "OBJ" & A)
        
        If Temp = "0-0" Then
            Call WriteVar(FilePath, "BANCOINVENTORY", "OBJ" & A, Auction.ObjIndex & "-" & Auction.Amount)
            Call Logs_Security(eSecurity, eSubastas, "El personaje " & Auction.Offer.Name & " recibió en su boveda: " & ObjData(Auction.ObjIndex).Name & " (x" & Auction.Amount & ")")
            Exit Sub

        End If

    Next A
    
    Call Logs_Security(eSecurity, eSubastas, "El personaje " & Auction.Offer.Name & " no recibió el objeto en su boveda: " & ObjData(Auction.ObjIndex).Name & " (x" & Auction.Amount & ")")
    '<EhFooter>
    Exit Sub

Auction_Object_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAuction.Auction_Object " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Auction_RewardGld()

    '<EhHeader>
    On Error GoTo Auction_RewardGld_Err

    '</EhHeader>
    Dim tUser      As Integer

    Dim FilePath   As String

    Dim TempGld    As Long

    Dim TempEldhir As Long
    
    tUser = NameIndex(Auction.Name)
        
    ' Subastador
    If tUser > 0 Then
        Call WriteConsoleMsg(tUser, "Has recibido el dinero de tu subasta. ¡Has acumulado " & Auction.Offer.Gld & " Monedas de Oro Y " & Auction.Offer.Eldhir & " Monedas de Eldhir.", FontTypeNames.FONTTYPE_INFOGREEN)
        UserList(tUser).Stats.Gld = UserList(tUser).Stats.Gld + Auction.Offer.Gld
        UserList(tUser).Stats.Eldhir = UserList(tUser).Stats.Eldhir + Auction.Offer.Eldhir
        Call WriteUpdateUserStats(tUser)
    Else
        FilePath = CharPath & Auction.Name & ".chr"
        TempGld = val(GetVar(FilePath, "STATS", "GLD"))
        TempEldhir = val(GetVar(FilePath, "STATS", "ELDHIR"))
            
        Call WriteVar(FilePath, "STATS", "GLD", CStr(TempGld + Auction.Offer.Gld))
        Call WriteVar(FilePath, "STATS", "ELDHIR", CStr(TempEldhir + Auction.Offer.Eldhir))

    End If

    '<EhFooter>
    Exit Sub

Auction_RewardGld_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAuction.Auction_RewardGld " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub Auction_RewardObj()

    '<EhHeader>
    On Error GoTo Auction_RewardObj_Err

    '</EhHeader>

    Dim tUser      As Integer

    Dim FilePath   As String

    Dim TempGld    As Long

    Dim TempEldhir As Long
    
    Dim Obj        As Obj
    
    With Auction
        ' Personaje que recibe el objeto
        tUser = NameIndex(.Offer.Name)
            
        If tUser > 0 Then
            Obj.Amount = Auction.Amount
            Obj.ObjIndex = Auction.ObjIndex
            
            If Not MeterItemEnInventario(tUser, Obj) Then
                Call Logs_Security(eSecurity, eSubastas, "El personaje " & .Offer.Name & " no tenia espacio en inventario, por lo que no recibió el objeto de la subasta: " & ObjData(.ObjIndex).Name & " (x" & .Amount & ")")
            Else
                Call Logs_Security(eSecurity, eSubastas, "El personaje " & .Offer.Name & " recibió " & ObjData(.ObjIndex).Name & " (x" & .Amount & ") ofertando " & .Offer.Gld & " Monedas de Oro Y " & .Offer.Eldhir & " Monedas de Eldhir.")

            End If

        Else
            Call Auction_Object

        End If

    End With

    '<EhFooter>
    Exit Sub

Auction_RewardObj_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAuction.Auction_RewardObj " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Auction_Reset()

    '<EhHeader>
    On Error GoTo Auction_Reset_Err

    '</EhHeader>

    Dim NullAuction As tAuction
    
    Auction = NullAuction

    '<EhFooter>
    Exit Sub

Auction_Reset_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAuction.Auction_Reset " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Function Auction_ObjValid(ByVal ObjIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo Auction_ObjValid_Err

    '</EhHeader>
    If ObjData(ObjIndex).OBJType = otBebidas Or ObjData(ObjIndex).OBJType = otBotellaLlena Or ObjData(ObjIndex).OBJType = otBotellaVacia Or ObjData(ObjIndex).OBJType = otPociones Or ObjData(ObjIndex).OBJType = otUseOnce Then
        
        Exit Function

    End If
    
    Auction_ObjValid = True
    '<EhFooter>
    Exit Function

Auction_ObjValid_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAuction.Auction_ObjValid " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub Auction_Cancel()

    '<EhHeader>
    On Error GoTo Auction_Cancel_Err

    '</EhHeader>
    If Auction.ObjIndex = 0 Then Exit Sub
    Call Auction_RewardObj
    Call Auction_Reset
    '<EhFooter>
    Exit Sub

Auction_Cancel_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mAuction.Auction_Cancel " & "at line " & Erl
        
    '</EhFooter>
End Sub
