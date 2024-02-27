Attribute VB_Name = "m_Auras"
Option Explicit
Public Enum UpdateAuras
    Arma
    Armadura
    Escudo
    casco
    Anillo
End Enum
 
Public Sub ActualizarAuras(ByVal UserIndex As Integer)
On Error GoTo errh
    With UserList(UserIndex)
        If .Invent.ArmourEqpObjIndex <> 0 Then
            With ObjData(.Invent.ArmourEqpObjIndex)
                If .Aura <> 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSendAura(UserIndex, .Aura, UpdateAuras.Armadura))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSendAura(UserIndex, 0, UpdateAuras.Armadura))
                End If
            End With
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSendAura(UserIndex, 0, UpdateAuras.Armadura))
        End If
 
        If .Invent.WeaponEqpObjIndex <> 0 Then
            With ObjData(.Invent.WeaponEqpObjIndex)
                If .Aura <> 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSendAura(UserIndex, .Aura, UpdateAuras.Arma))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSendAura(UserIndex, 0, UpdateAuras.Arma))
                End If
            End With
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSendAura(UserIndex, 0, UpdateAuras.Arma))
        End If
 
        If .Invent.EscudoEqpObjIndex <> 0 Then
            With ObjData(.Invent.EscudoEqpObjIndex)
                If .Aura <> 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSendAura(UserIndex, .Aura, UpdateAuras.Escudo))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSendAura(UserIndex, 0, UpdateAuras.Escudo))
                End If
            End With
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSendAura(UserIndex, 0, UpdateAuras.Escudo))
        End If
 
        If .Invent.CascoEqpObjIndex <> 0 Then
            With ObjData(.Invent.CascoEqpObjIndex)
                If .Aura <> 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSendAura(UserIndex, .Aura, UpdateAuras.casco))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSendAura(UserIndex, 0, UpdateAuras.casco))
                End If
            End With
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSendAura(UserIndex, 0, UpdateAuras.casco))
        End If
 
        If .Invent.AnilloEqpObjIndex <> 0 Then
            With ObjData(.Invent.AnilloEqpObjIndex)
                If .Aura <> 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSendAura(UserIndex, .Aura, UpdateAuras.Anillo))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSendAura(UserIndex, 0, UpdateAuras.Anillo))
                End If
            End With
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSendAura(UserIndex, 0, UpdateAuras.Anillo))
        End If
 
    End With
    Exit Sub
errh:
MsgBox "Error: " & Err.Number & " " & Err.description
End Sub
