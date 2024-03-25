Attribute VB_Name = "EventosDS_Reward"
' Los usuarios podrán compartir sus objetos y donarlos para los eventos predeterminados del juego

Option Explicit

' Buscamos un Slot Libre para agregar el objeto
Private Function Events_Reward_Slot(ByVal SlotEvent As Byte) As Byte

    '<EhHeader>
    On Error GoTo Events_Reward_Slot_Err

    '</EhHeader>

    Dim A As Long
    
    For A = 1 To MAX_REWARD_OBJ

        If Events(SlotEvent).RewardObj(A).ObjIndex = 0 Then
            Events_Reward_Slot = A
            Exit For

        End If

    Next A

    '<EhFooter>
    Exit Function

Events_Reward_Slot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.EventosDS_Reward.Events_Reward_Slot " & "at line " & Erl
        
    '</EhFooter>
End Function

' Agregamos un nuevo objeto a la lista de PREMIOS DONADOS
Public Sub Events_Reward_Add(ByVal UserIndex As Integer, _
                             ByVal SlotEvent As Byte, _
                             ByVal Slot As Byte, _
                             ByVal Amount As Integer)

    '<EhHeader>
    On Error GoTo Events_Reward_Add_Err

    '</EhHeader>
                             
    Dim ObjIndex   As Integer

    Dim A          As Long

    Dim SlotReward As Byte
    
    ' Chequeamos que haya lugar para guardar el nuevo objeto
    If Events(SlotEvent).LastReward = MAX_REWARD_OBJ Then
        Call WriteConsoleMsg(UserIndex, "No hay más espacio para agregar premios al evento.", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
    
    ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    SlotReward = Events_Reward_Slot(SlotEvent)
    
    ' Quitamos el Objeto al Usuario
    Call QuitarUserInvItem(UserIndex, Slot, Amount)
    Call UpdateUserInv(False, UserIndex, Slot)
        
    ' Agregamos el Objeto a la lista de premios
    Events(SlotEvent).RewardObj(SlotReward).ObjIndex = ObjIndex
    Events(SlotEvent).RewardObj(SlotReward).Amount = Amount
    Events(SlotEvent).LastReward = Events(SlotEvent).LastReward + 1
    Call WriteConsoleMsg(UserIndex, "Has donado para el evento " & Events(SlotEvent).Name & " el objeto " & ObjData(ObjIndex).Name & " (x" & Amount & ")", FontTypeNames.FONTTYPE_INFOGREEN)
    
    LogEventos "El personaje " & UserList(UserIndex).Name & " Ha donado para el evento " & Events(SlotEvent).Name & " el objeto " & ObjData(ObjIndex).Name & " (x" & Amount & ")"
    '<EhFooter>
    Exit Sub

Events_Reward_Add_Err:
    LogError Err.description & vbCrLf & "in Events_Reward_Add " & "at line " & Erl

    '</EhFooter>
End Sub

' Recorre la lista de premios donados y en caso de caer sobre uno válido, se lo "regala" al personaje campeón.
Public Sub Events_Reward_User(ByVal UserIndex As Integer, ByVal SlotEvent As Byte)

    '<EhHeader>
    On Error GoTo Events_Reward_User_Err

    '</EhHeader>
    
    Dim A    As Long

    Dim Slot As Byte
    
    With Events(SlotEvent)
        Slot = RandomNumber(1, MAX_REWARD_OBJ)
        
        If .RewardObj(Slot).ObjIndex > 0 Then
            If Not MeterItemEnInventario(UserIndex, .RewardObj(Slot)) Then
                WriteConsoleMsg UserIndex, "Tu premio Donado no ha sido entregado, envia esta foto a un Game Master.", FontTypeNames.FONTTYPE_INFO
                LogEventos ("Personaje " & UserList(UserIndex).Name & " no recibió: " & .RewardObj(Slot).ObjIndex & " (x" & .RewardObj(Slot).Amount & ")")
                Exit Sub

            End If
            
            .RewardObj(Slot).ObjIndex = 0
            .RewardObj(Slot).Amount = 0
            
            Call WriteConsoleMsg(UserIndex, "Un premio donado ha caído sobre tí por haber ganado el evento ¡Esto no siempre sucede, Felicitaciones!", FontTypeNames.FONTTYPE_INFOGREEN)
        
        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

Events_Reward_User_Err:
    LogError Err.description & vbCrLf & "in Events_Reward_User " & "at line " & Erl

    '</EhFooter>
End Sub

