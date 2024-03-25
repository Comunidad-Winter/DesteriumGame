Attribute VB_Name = "modBanco"
'**************************************************************
' modBanco.bas - Handles the character's bank accounts.
'
' Implemented by Kevin Birmingham (NEB)
' kbneb@hotmail.com
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

Public Enum E_BANK

    e_User = 1
    e_Account = 2

End Enum

Sub IniciarDeposito(ByVal UserIndex As Integer, ByVal TypeBank As E_BANK)

    '<EhHeader>
    On Error GoTo IniciarDeposito_Err

    '</EhHeader>
                    
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    Select Case TypeBank

        Case E_BANK.e_User
            Call UpdateBanUserInv(True, UserIndex, 0)
            
        Case E_BANK.e_Account

            If UserList(UserIndex).Account.Premium < 2 Then
                Call WriteConsoleMsg(UserIndex, "Solo las cuentas TIER 2 o superior poseen un banco exclusivo. Consulta las promociones en www.argentumgame.com/download", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub

            End If

            Call UpdateBanUserInv_Account(True, UserIndex, 0)
        
    End Select
    
    Call WriteBankInit(UserIndex, TypeBank)
    'Call WriteUpdateUserStats(UserIndex)
    
    UserList(UserIndex).flags.Comerciando = True

    '<EhFooter>
    Exit Sub

IniciarDeposito_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.IniciarDeposito " & "at line " & Erl

    '</EhFooter>
End Sub

Sub SendBanObj(ByVal UserIndex As Integer, _
               ByVal Slot As Byte, _
               ByRef Object As UserOBJ, _
               ByVal TypeBank As E_BANK)

    '<EhHeader>
    On Error GoTo SendBanObj_Err

    '</EhHeader>
               
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    Select Case TypeBank

        Case E_BANK.e_User
            UserList(UserIndex).BancoInvent.Object(Slot) = Object
            Call WriteChangeBankSlot(UserIndex, Slot)
            
        Case E_BANK.e_Account
            UserList(UserIndex).Account.BancoInvent.Object(Slot) = Object
            Call WriteChangeBankSlot_Account(UserIndex, Slot)

    End Select

    '<EhFooter>
    Exit Sub

SendBanObj_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.SendBanObj " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, _
                     ByVal UserIndex As Integer, _
                     ByVal Slot As Byte)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo UpdateBanUserInv_Err

    '</EhHeader>

    Dim NullObj As UserOBJ

    Dim LoopC   As Byte

    With UserList(UserIndex)

        'Actualiza un solo slot
        If Not UpdateAll Then

            'Actualiza el inventario
            If .BancoInvent.Object(Slot).ObjIndex > 0 Then
                Call SendBanObj(UserIndex, Slot, .BancoInvent.Object(Slot), e_User)
            Else
                Call SendBanObj(UserIndex, Slot, NullObj, e_User)

            End If

        Else

            'Actualiza todos los slots
            For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS

                'Actualiza el inventario
                If .BancoInvent.Object(LoopC).ObjIndex > 0 Then
                    Call SendBanObj(UserIndex, LoopC, .BancoInvent.Object(LoopC), e_User)
                Else
                    Call SendBanObj(UserIndex, LoopC, NullObj, e_User)

                End If
            
            Next LoopC

        End If

    End With

    '<EhFooter>
    Exit Sub

UpdateBanUserInv_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.UpdateBanUserInv " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub UpdateBanUserInv_Account(ByVal UpdateAll As Boolean, _
                             ByVal UserIndex As Integer, _
                             ByVal Slot As Byte)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo UpdateBanUserInv_Account_Err

    '</EhHeader>

    Dim NullObj As UserOBJ

    Dim LoopC   As Byte

    With UserList(UserIndex).Account

        'Actualiza un solo slot
        If Not UpdateAll Then

            'Actualiza el inventario
            If .BancoInvent.Object(Slot).ObjIndex > 0 Then
                Call SendBanObj(UserIndex, Slot, .BancoInvent.Object(Slot), e_Account)
            Else
                Call SendBanObj(UserIndex, Slot, NullObj, e_Account)

            End If

        Else

            'Actualiza todos los slots
            For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS

                'Actualiza el inventario
                If .BancoInvent.Object(LoopC).ObjIndex > 0 Then
                    Call SendBanObj(UserIndex, LoopC, .BancoInvent.Object(LoopC), e_Account)
                Else
                    Call SendBanObj(UserIndex, LoopC, NullObj, e_Account)

                End If
            
            Next LoopC

        End If

    End With

    '<EhFooter>
    Exit Sub

UpdateBanUserInv_Account_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.UpdateBanUserInv_Account " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub UserRetiraItem(ByVal UserIndex As Integer, _
                   ByVal i As Integer, _
                   ByVal cantidad As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo UserRetiraItem_Err

    '</EhHeader>

    Dim ObjIndex  As Integer

    Dim SlotEvent As Byte

    If cantidad < 1 Then Exit Sub
    
    Call WriteUpdateUserStats(UserIndex)
    
    If UserList(UserIndex).BancoInvent.Object(i).Amount > 0 Then
        
        If cantidad > UserList(UserIndex).BancoInvent.Object(i).Amount Then cantidad = UserList(UserIndex).BancoInvent.Object(i).Amount
            
        ObjIndex = UserList(UserIndex).BancoInvent.Object(i).ObjIndex
        
        SlotEvent = UserList(UserIndex).flags.SlotEvent
        
        If SlotEvent > 0 Then
            If Events(SlotEvent).LimitRed > 0 Then
                If ObjIndex = POCION_ROJA Then
                    Call WriteConsoleMsg(UserIndex, "No puedes retirar pociones rojas en éste tipo de eventos.", FontTypeNames.FONTTYPE_INFORED)

                    Exit Sub

                End If

            End If
            
            If Events(SlotEvent).ChangeClass > 0 Or Events(SlotEvent).ChangeRaze > 0 Or Events(SlotEvent).ChangeLevel > 0 Then
                If Events(SlotEvent).TimeCancel > 0 Then
                    Call WriteConsoleMsg(UserIndex, "No puedes retirar objetos en este tipo de eventos.", FontTypeNames.FONTTYPE_INFORED)

                    Exit Sub

                End If

            End If

        End If
  
        'Agregamos el obj que compro al inventario
        Call UserReciveObj(UserIndex, CInt(i), cantidad)
        
        If ObjData(ObjIndex).Log = 1 Then
            Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eBov_Obj, UserList(UserIndex).Name & " retiró " & cantidad & " " & ObjData(ObjIndex).Name & "[" & ObjIndex & "]")

        End If

    End If
    
    'Actualizamos la ventana de comercio
    Call UpdateVentanaBanco(UserIndex)

    '<EhFooter>
    Exit Sub

UserRetiraItem_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.UserRetiraItem " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub UserRetiraItem_Account(ByVal UserIndex As Integer, _
                           ByVal i As Integer, _
                           ByVal cantidad As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo UserRetiraItem_Account_Err

    '</EhHeader>

    Dim ObjIndex  As Integer

    Dim SlotEvent As Byte
        
    If UserList(UserIndex).Account.Premium < 2 Then
        Call WriteConsoleMsg(UserIndex, "Solo las cuentas TIER 2 o superior poseen un banco exclusivo. Consulta las promociones en www.argentumgame.com/download", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
                
    If cantidad < 1 Then Exit Sub
    
    Call WriteUpdateUserStats(UserIndex)
    
    If UserList(UserIndex).Account.BancoInvent.Object(i).Amount > 0 Then
        
        If cantidad > UserList(UserIndex).Account.BancoInvent.Object(i).Amount Then cantidad = UserList(UserIndex).Account.BancoInvent.Object(i).Amount
            
        ObjIndex = UserList(UserIndex).Account.BancoInvent.Object(i).ObjIndex
        
        SlotEvent = UserList(UserIndex).flags.SlotEvent
        
        If SlotEvent > 0 Then
            If Events(SlotEvent).LimitRed > 0 Then
                If ObjIndex = POCION_ROJA Then
                    Call WriteConsoleMsg(UserIndex, "No puedes retirar pociones rojas en éste tipo de eventos.", FontTypeNames.FONTTYPE_INFORED)

                    Exit Sub

                End If

            End If
            
            If Events(SlotEvent).ChangeClass > 0 Or Events(SlotEvent).ChangeRaze > 0 Or Events(SlotEvent).ChangeLevel > 0 Then
                If Events(SlotEvent).TimeCancel > 0 Then
                    Call WriteConsoleMsg(UserIndex, "No puedes retirar objetos en este tipo de eventos.", FontTypeNames.FONTTYPE_INFORED)

                    Exit Sub

                End If

            End If

        End If
  
        'Agregamos el obj que compro al inventario
        Call UserReciveObj_Account(UserIndex, CInt(i), cantidad)
        
        If ObjData(ObjIndex).Log = 1 Then
            Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eBov_Obj, UserList(UserIndex).Name & " retiró " & cantidad & " " & ObjData(ObjIndex).Name & "[" & ObjIndex & "]")

        End If

    End If
    
    'Actualizamos la ventana de comercio
    Call UpdateVentanaBanco(UserIndex)

    '<EhFooter>
    Exit Sub

UserRetiraItem_Account_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.UserRetiraItem_Account " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub UserReciveObj(ByVal UserIndex As Integer, _
                  ByVal ObjIndex As Integer, _
                  ByVal cantidad As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo UserReciveObj_Err

    '</EhHeader>

    Dim Slot As Integer

    Dim obji As Integer

    With UserList(UserIndex)

        If .BancoInvent.Object(ObjIndex).Amount <= 0 Then Exit Sub
    
        obji = .BancoInvent.Object(ObjIndex).ObjIndex
    
        '¿Ya tiene un objeto de este tipo?
        Slot = 1

        Do Until .Invent.Object(Slot).ObjIndex = obji And .Invent.Object(Slot).Amount + cantidad <= MAX_INVENTORY_OBJS
        
            Slot = Slot + 1

            If Slot > .CurrentInventorySlots Then

                Exit Do

            End If

        Loop
    
        'Sino se fija por un slot vacio
        If Slot > .CurrentInventorySlots Then
            Slot = 1

            Do Until .Invent.Object(Slot).ObjIndex = 0
                Slot = Slot + 1

                If Slot > .CurrentInventorySlots Then
                    Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If

            Loop

            .Invent.NroItems = .Invent.NroItems + 1

        End If
    
        'Mete el obj en el slot
        If .Invent.Object(Slot).Amount + cantidad <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
            .Invent.Object(Slot).ObjIndex = obji
            .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + cantidad
        
            Call QuitarBancoInvItem(UserIndex, CByte(ObjIndex), cantidad)
        
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(False, UserIndex, Slot)
        
            'Actualizamos el banco
            Call UpdateBanUserInv(False, UserIndex, ObjIndex)
        Else
            Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    '<EhFooter>
    Exit Sub

UserReciveObj_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.UserReciveObj " & "at line " & Erl

    '</EhFooter>
End Sub

Sub UserReciveObj_Account(ByVal UserIndex As Integer, _
                          ByVal ObjIndex As Integer, _
                          ByVal cantidad As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo UserReciveObj_Account_Err

    '</EhHeader>

    Dim Slot As Integer

    Dim obji As Integer
        
    If UserList(UserIndex).Account.Premium < 2 Then
        Call WriteConsoleMsg(UserIndex, "Solo las cuentas TIER 2 o superior poseen un banco exclusivo. Consulta las promociones en www.argentumgame.com/download", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
                
    With UserList(UserIndex)

        If .Account.BancoInvent.Object(ObjIndex).Amount <= 0 Then Exit Sub
    
        obji = .Account.BancoInvent.Object(ObjIndex).ObjIndex
    
        '¿Ya tiene un objeto de este tipo?
        Slot = 1

        Do Until .Invent.Object(Slot).ObjIndex = obji And .Invent.Object(Slot).Amount + cantidad <= MAX_INVENTORY_OBJS
        
            Slot = Slot + 1

            If Slot > .CurrentInventorySlots Then

                Exit Do

            End If

        Loop
    
        'Sino se fija por un slot vacio
        If Slot > .CurrentInventorySlots Then
            Slot = 1

            Do Until .Invent.Object(Slot).ObjIndex = 0
                Slot = Slot + 1

                If Slot > .CurrentInventorySlots Then
                    Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If

            Loop

            .Invent.NroItems = .Invent.NroItems + 1

        End If
    
        'Mete el obj en el slot
        If .Invent.Object(Slot).Amount + cantidad <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
            .Invent.Object(Slot).ObjIndex = obji
            .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + cantidad
        
            Call QuitarBancoInvItem_Account(UserIndex, CByte(ObjIndex), cantidad)
        
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(False, UserIndex, Slot)
        
            'Actualizamos el banco
            Call UpdateBanUserInv_Account(False, UserIndex, ObjIndex)
        Else
            Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    '<EhFooter>
    Exit Sub

UserReciveObj_Account_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.UserReciveObj_Account " & "at line " & Erl

    '</EhFooter>
End Sub

Sub QuitarBancoInvItem(ByVal UserIndex As Integer, _
                       ByVal Slot As Byte, _
                       ByVal cantidad As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo QuitarBancoInvItem_Err

    '</EhHeader>

    Dim ObjIndex As Integer

    With UserList(UserIndex)
        ObjIndex = .BancoInvent.Object(Slot).ObjIndex

        'Quita un Obj

        .BancoInvent.Object(Slot).Amount = .BancoInvent.Object(Slot).Amount - cantidad
    
        If .BancoInvent.Object(Slot).Amount <= 0 Then
            .BancoInvent.NroItems = .BancoInvent.NroItems - 1
            .BancoInvent.Object(Slot).ObjIndex = 0
            .BancoInvent.Object(Slot).Amount = 0

        End If

    End With
    
    '<EhFooter>
    Exit Sub

QuitarBancoInvItem_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.QuitarBancoInvItem " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub QuitarBancoInvItem_Account(ByVal UserIndex As Integer, _
                               ByVal Slot As Byte, _
                               ByVal cantidad As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo QuitarBancoInvItem_Account_Err

    '</EhHeader>

    Dim ObjIndex As Integer

    With UserList(UserIndex)
        ObjIndex = .Account.BancoInvent.Object(Slot).ObjIndex

        'Quita un Obj

        .Account.BancoInvent.Object(Slot).Amount = .Account.BancoInvent.Object(Slot).Amount - cantidad
    
        If .Account.BancoInvent.Object(Slot).Amount <= 0 Then
            .Account.BancoInvent.NroItems = .Account.BancoInvent.NroItems - 1
            .Account.BancoInvent.Object(Slot).ObjIndex = 0
            .Account.BancoInvent.Object(Slot).Amount = 0

        End If

    End With
    
    '<EhFooter>
    Exit Sub

QuitarBancoInvItem_Account_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.QuitarBancoInvItem_Account " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub UpdateVentanaBanco(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo UpdateVentanaBanco_Err

    '</EhHeader>

    Call WriteBankOK(UserIndex)
    '<EhFooter>
    Exit Sub

UpdateVentanaBanco_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.UpdateVentanaBanco " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub UserDepositaItem(ByVal UserIndex As Integer, _
                     ByVal Item As Integer, _
                     ByVal cantidad As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo UserDepositaItem_Err

    '</EhHeader>

    Dim ObjIndex As Integer

    If UserList(UserIndex).Invent.Object(Item).Amount > 0 And cantidad > 0 Then
    
        If cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then cantidad = UserList(UserIndex).Invent.Object(Item).Amount
        
        ObjIndex = UserList(UserIndex).Invent.Object(Item).ObjIndex
            
        If ObjData(ObjIndex).NoNada = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes guardar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
                
        End If
            
        If ObjData(ObjIndex).OBJType = otTransformVIP Then
            If UserList(UserIndex).flags.TransformVIP = 1 Then
                Call TransformVIP_User(UserIndex, 0)

            End If

        End If
        
        If ObjData(ObjIndex).OBJType = otGemaTelep Then
            Call WriteConsoleMsg(UserIndex, "No puedes guardar este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        'Agregamos el obj que deposita al banco
        Call UserDejaObj(UserIndex, CInt(Item), cantidad)
        
        If ObjData(ObjIndex).Log = 1 Then
            Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eBov_Obj, UserList(UserIndex).Name & " depositó " & cantidad & " " & ObjData(ObjIndex).Name & "[" & ObjIndex & "]")

        End If

    End If
    
    'Actualizamos la ventana del banco
    Call UpdateVentanaBanco(UserIndex)

    '<EhFooter>
    Exit Sub

UserDepositaItem_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.UserDepositaItem " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub UserDepositaItem_Account(ByVal UserIndex As Integer, _
                             ByVal Item As Integer, _
                             ByVal cantidad As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo UserDepositaItem_Account_Err

    '</EhHeader>

    Dim ObjIndex As Integer

    If UserList(UserIndex).Invent.Object(Item).Amount > 0 And cantidad > 0 Then
    
        If cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then cantidad = UserList(UserIndex).Invent.Object(Item).Amount
        
        ObjIndex = UserList(UserIndex).Invent.Object(Item).ObjIndex
                
        If ObjData(ObjIndex).OBJType = otTransformVIP Then
            If UserList(UserIndex).flags.TransformVIP = 1 Then
                Call TransformVIP_User(UserIndex, 0)

            End If

        End If
        
        If ObjData(ObjIndex).OBJType = otGemaTelep Then
            Call WriteConsoleMsg(UserIndex, "No puedes guardar este objeto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If ObjData(ObjIndex).NoNada = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes guardar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
                
        End If
            
        'Agregamos el obj que deposita al banco
        Call UserDejaObj_Account(UserIndex, CInt(Item), cantidad)
        
        If ObjData(ObjIndex).Log = 1 Then
            Call Logs_User(UserList(UserIndex).Name, eLog.eUser, eBov_Obj, UserList(UserIndex).Name & " depositó " & cantidad & " " & ObjData(ObjIndex).Name & "[" & ObjIndex & "]")

        End If

    End If
    
    'Actualizamos la ventana del banco
    Call UpdateVentanaBanco(UserIndex)

    '<EhFooter>
    Exit Sub

UserDepositaItem_Account_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.UserDepositaItem_Account " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub UserDejaObj(ByVal UserIndex As Integer, _
                ByVal ObjIndex As Integer, _
                ByVal cantidad As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo UserDejaObj_Err

    '</EhHeader>

    Dim Slot As Integer

    Dim obji As Integer
    
    If cantidad < 1 Then Exit Sub
    
    With UserList(UserIndex)
        obji = .Invent.Object(ObjIndex).ObjIndex
        
        '¿Ya tiene un objeto de este tipo?
        Slot = 1

        Do Until .BancoInvent.Object(Slot).ObjIndex = obji And .BancoInvent.Object(Slot).Amount + cantidad <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
            
            If Slot > MAX_BANCOINVENTORY_SLOTS Then

                Exit Do

            End If

        Loop
        
        'Sino se fija por un slot vacio antes del slot devuelto
        If Slot > MAX_BANCOINVENTORY_SLOTS Then
            Slot = 1

            Do Until .BancoInvent.Object(Slot).ObjIndex = 0
                Slot = Slot + 1
                
                If Slot > MAX_BANCOINVENTORY_SLOTS Then
                    Call WriteConsoleMsg(UserIndex, "No tienes mas espacio en el banco!!", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If

            Loop
            
            .BancoInvent.NroItems = .BancoInvent.NroItems + 1

        End If
        
        If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido

            'Mete el obj en el slot
            If .BancoInvent.Object(Slot).Amount + cantidad <= MAX_INVENTORY_OBJS Then
                
                'Menor que MAX_INV_OBJS
                .BancoInvent.Object(Slot).ObjIndex = obji
                .BancoInvent.Object(Slot).Amount = .BancoInvent.Object(Slot).Amount + cantidad
                
                Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), cantidad)
                        
                'Actualizamos el inventario del usuario
                Call UpdateUserInv(False, UserIndex, ObjIndex)
                
                'Actualizamos el inventario del banco
                Call UpdateBanUserInv(False, UserIndex, Slot)
            Else
                Call WriteConsoleMsg(UserIndex, "El banco no puede cargar tantos objetos.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

    '<EhFooter>
    Exit Sub

UserDejaObj_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.UserDejaObj " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub UserDejaObj_Account(ByVal UserIndex As Integer, _
                        ByVal ObjIndex As Integer, _
                        ByVal cantidad As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo UserDejaObj_Account_Err

    '</EhHeader>

    Dim Slot As Integer

    Dim obji As Integer
    
    If cantidad < 1 Then Exit Sub
    
    With UserList(UserIndex)
        obji = .Invent.Object(ObjIndex).ObjIndex
        
        '¿Ya tiene un objeto de este tipo?
        Slot = 1

        Do Until .Account.BancoInvent.Object(Slot).ObjIndex = obji And .Account.BancoInvent.Object(Slot).Amount + cantidad <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
            
            If Slot > MAX_BANCOINVENTORY_SLOTS Then

                Exit Do

            End If

        Loop
        
        'Sino se fija por un slot vacio antes del slot devuelto
        If Slot > MAX_BANCOINVENTORY_SLOTS Then
            Slot = 1

            Do Until .Account.BancoInvent.Object(Slot).ObjIndex = 0
                Slot = Slot + 1
                
                If Slot > MAX_BANCOINVENTORY_SLOTS Then
                    Call WriteConsoleMsg(UserIndex, "No tienes mas espacio en el banco!!", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If

            Loop
            
            .Account.BancoInvent.NroItems = .Account.BancoInvent.NroItems + 1

        End If
        
        If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido

            'Mete el obj en el slot
            If .Account.BancoInvent.Object(Slot).Amount + cantidad <= MAX_INVENTORY_OBJS Then
                
                'Menor que MAX_INV_OBJS
                .Account.BancoInvent.Object(Slot).ObjIndex = obji
                .Account.BancoInvent.Object(Slot).Amount = .Account.BancoInvent.Object(Slot).Amount + cantidad
                
                Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), cantidad)
                        
                'Actualizamos el inventario del usuario
                Call UpdateUserInv(False, UserIndex, ObjIndex)
                
                'Actualizamos el inventario del banco
                Call UpdateBanUserInv_Account(False, UserIndex, Slot)
            Else
                Call WriteConsoleMsg(UserIndex, "El banco no puede cargar tantos objetos.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

    '<EhFooter>
    Exit Sub

UserDejaObj_Account_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.UserDejaObj_Account " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub SendUserBovedaTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo SendUserBovedaTxt_Err

    '</EhHeader>

    Dim j As Integer

    Call WriteConsoleMsg(sendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Tiene " & UserList(UserIndex).BancoInvent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)

    For j = 1 To MAX_BANCOINVENTORY_SLOTS

        If UserList(UserIndex).BancoInvent.Object(j).ObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(UserList(UserIndex).BancoInvent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).BancoInvent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)

        End If

    Next

    '<EhFooter>
    Exit Sub

SendUserBovedaTxt_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.SendUserBovedaTxt " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub SendUserBovedaTxt_Account(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo SendUserBovedaTxt_Account_Err

    '</EhHeader>

    Dim j As Integer

    Call WriteConsoleMsg(sendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Tiene " & UserList(UserIndex).Account.BancoInvent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)

    For j = 1 To MAX_BANCOINVENTORY_SLOTS

        If UserList(UserIndex).Account.BancoInvent.Object(j).ObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(UserList(UserIndex).Account.BancoInvent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).Account.BancoInvent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)

        End If

    Next

    '<EhFooter>
    Exit Sub

SendUserBovedaTxt_Account_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.SendUserBovedaTxt_Account " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub SendUserBovedaTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo SendUserBovedaTxtFromChar_Err

    '</EhHeader>

    Dim j        As Integer

    Dim Charfile As String, Tmp As String

    Dim ObjInd   As Long, ObjCant As Long

    Charfile = CharPath & charName & ".chr"

    If FileExist(Charfile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(Charfile, "BancoInventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)

        For j = 1 To MAX_BANCOINVENTORY_SLOTS
            Tmp = GetVar(Charfile, "BancoInventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))

            If ObjInd > 0 Then
                Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)

            End If

        Next

    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)

    End If

    '<EhFooter>
    Exit Sub

SendUserBovedaTxtFromChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.SendUserBovedaTxtFromChar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub SendUserBovedaTxtFromChar_Account(ByVal sendIndex As Integer, _
                                      ByVal charName As String)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo SendUserBovedaTxtFromChar_Account_Err

    '</EhHeader>

    Dim j        As Integer

    Dim Charfile As String, Tmp As String

    Dim ObjInd   As Long, ObjCant As Long
    
    Dim Account  As String
    
    Charfile = CharPath & charName & ".chr"
    
    Account = AccountPath & GetVar(Charfile, "INIT", "ACCOUNTNAME") & ACCOUNT_FORMAT
    
    If FileExist(Account, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(Account, "BancoInventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)

        For j = 1 To MAX_BANCOINVENTORY_SLOTS
            Tmp = GetVar(Account, "BancoInventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))

            If ObjInd > 0 Then
                Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)

            End If

        Next

    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)

    End If

    '<EhFooter>
    Exit Sub

SendUserBovedaTxtFromChar_Account_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modBanco.SendUserBovedaTxtFromChar_Account " & "at line " & Erl
        
    '</EhFooter>
End Sub
