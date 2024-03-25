Attribute VB_Name = "modNuevoTimer"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

'
' Las siguientes funciones devuelven TRUE o FALSE si el intervalo
' permite hacerlo. Si devuelve TRUE, setean automaticamente el
' timer para que no se pueda hacer la accion hasta el nuevo ciclo.
'

' CASTING DE HECHIZOS
Public Function IntervaloPermiteLanzarSpell(ByVal UserIndex As Integer, _
                                            Optional ByVal Actualizar As Boolean = True) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim TActual As Long

    TActual = GetTime

    If TActual - UserList(UserIndex).Counters.TimerLanzarSpell >= IntervaloUserPuedeCastear Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerLanzarSpell = TActual
            ' Actualizo spell-attack
            UserList(UserIndex).Counters.TimerMagiaGolpe = TActual

        End If

        IntervaloPermiteLanzarSpell = True
    Else
        IntervaloPermiteLanzarSpell = False

    End If

End Function

Public Function IntervaloPermiteShiftear(ByVal UserIndex As Integer, _
                                         Optional ByVal Actualizar As Boolean = True) As Boolean

    Dim TActual As Long

    TActual = GetTime

    If TActual - UserList(UserIndex).Counters.TimerShiftear >= IntervaloUserPuedeShiftear Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerShiftear = TActual
            ' Actualizo spell-attack
            UserList(UserIndex).Counters.TimerShiftear = TActual

        End If

        IntervaloPermiteShiftear = True
    Else
        IntervaloPermiteShiftear = False

    End If

End Function

Public Function IntervaloPermiteCaspear(ByVal UserIndex As Integer, _
                                        Optional ByVal Actualizar As Boolean = True) As Boolean

    Dim TActual As Long

    TActual = GetTime

    If TActual - UserList(UserIndex).Counters.CaspeoTime >= 2000 Then
        If Actualizar Then
            UserList(UserIndex).Counters.CaspeoTime = TActual

        End If

        IntervaloPermiteCaspear = True
    Else
        IntervaloPermiteCaspear = False

    End If

End Function

Public Function IntervaloPermiteMoverUsuario(ByVal UserIndex As Integer, _
                                             Optional ByVal Actualizar As Boolean = True) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo IntervaloPermiteLanzarSpell_Err

    '</EhHeader>

    Dim TActual As Long

    TActual = GetTime

    If TActual - UserList(UserIndex).Counters.TimerLanzarSpell >= IntervaloUserPuedeCastear Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerLanzarSpell = TActual
            ' Actualizo spell-attack
            UserList(UserIndex).Counters.TimerMagiaGolpe = TActual

        End If

        IntervaloPermiteMoverUsuario = True
    Else
        IntervaloPermiteMoverUsuario = False

    End If

    '<EhFooter>
    Exit Function

IntervaloPermiteLanzarSpell_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.IntervaloPermiteLanzarSpell " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, _
                                       Optional ByVal Actualizar As Boolean = True) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo IntervaloPermiteAtacar_Err

    '</EhHeader>

    Dim TActual As Long

    TActual = GetTime

    If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= IntervaloUserPuedeAtacar Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
            ' Actualizo attack-spell
            UserList(UserIndex).Counters.TimerGolpeMagia = TActual
            ' Actualizo attack-use
            UserList(UserIndex).Counters.TimerGolpeUsar = TActual

        End If

        IntervaloPermiteAtacar = True
    Else
        IntervaloPermiteAtacar = False

    End If

    '<EhFooter>
    Exit Function

IntervaloPermiteAtacar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.IntervaloPermiteAtacar " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function IntervaloPermiteGolpeUsar(ByVal UserIndex As Integer, _
                                          Optional ByVal Actualizar As Boolean = True) As Boolean

    '***************************************************
    'Author: ZaMa
    'Checks if the time that passed from the last hit is enough for the user to use a potion.
    'Last Modification: 06/04/2009
    '***************************************************
    '<EhHeader>
    On Error GoTo IntervaloPermiteGolpeUsar_Err

    '</EhHeader>

    Dim TActual As Long

    TActual = GetTime

    If TActual - UserList(UserIndex).Counters.TimerGolpeUsar >= IntervaloGolpeUsar Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerGolpeUsar = TActual

        End If

        IntervaloPermiteGolpeUsar = True
    Else
        IntervaloPermiteGolpeUsar = False

    End If

    '<EhFooter>
    Exit Function

IntervaloPermiteGolpeUsar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.IntervaloPermiteGolpeUsar " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function IntervaloPermiteMagiaGolpe(ByVal UserIndex As Integer, _
                                           Optional ByVal Actualizar As Boolean = True) As Boolean

    '<EhHeader>
    On Error GoTo IntervaloPermiteMagiaGolpe_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Dim TActual As Long
    
    With UserList(UserIndex)
        
        TActual = GetTime
        
        If TActual - .Counters.TimerLanzarSpell >= IntervaloMagiaGolpe Then
            If Actualizar Then
                .Counters.TimerMagiaGolpe = TActual

            End If

            IntervaloPermiteMagiaGolpe = True
        Else
            IntervaloPermiteMagiaGolpe = False

        End If

    End With

    '<EhFooter>
    Exit Function

IntervaloPermiteMagiaGolpe_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.IntervaloPermiteMagiaGolpe " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function IntervaloPermiteGolpeMagia(ByVal UserIndex As Integer, _
                                           Optional ByVal Actualizar As Boolean = True) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo IntervaloPermiteGolpeMagia_Err

    '</EhHeader>

    Dim TActual As Long
    
    TActual = GetTime
    
    If TActual - UserList(UserIndex).Counters.TimerGolpeMagia >= IntervaloGolpeMagia Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerGolpeMagia = TActual

        End If

        IntervaloPermiteGolpeMagia = True
    Else
        IntervaloPermiteGolpeMagia = False

    End If

    '<EhFooter>
    Exit Function

IntervaloPermiteGolpeMagia_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.IntervaloPermiteGolpeMagia " & "at line " & Erl
        
    '</EhFooter>
End Function

' TRABAJO
Public Function IntervaloPermiteTrabajar(ByVal UserIndex As Integer, _
                                         Optional ByVal Actualizar As Boolean = True) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo IntervaloPermiteTrabajar_Err

    '</EhHeader>

    Dim TActual As Long
    
    TActual = GetTime
    
    If TActual - UserList(UserIndex).Counters.TimerPuedeTrabajar >= IntervaloUserPuedeTrabajar Then
        If Actualizar Then UserList(UserIndex).Counters.TimerPuedeTrabajar = TActual
        IntervaloPermiteTrabajar = True
    Else
        IntervaloPermiteTrabajar = False

    End If

    '<EhFooter>
    Exit Function

IntervaloPermiteTrabajar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.IntervaloPermiteTrabajar " & "at line " & Erl
        
    '</EhFooter>
End Function

' USAR OBJETOS
Public Function IntervaloPermiteUsar(ByVal UserIndex As Integer, _
                                     Optional ByVal Actualizar As Boolean = True) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: 25/01/2010 (ZaMa)
    '25/01/2010: ZaMa - General adjustments.
    '***************************************************
    '<EhHeader>
    On Error GoTo IntervaloPermiteUsar_Err

    '</EhHeader>

    Dim TActual As Long
    
    TActual = GetTime
    
    If TActual - UserList(UserIndex).Counters.TimerUsar >= IntervaloUserPuedeUsar Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerUsar = TActual
            UserList(UserIndex).Counters.failedUsageAttempts = 0

        End If

        IntervaloPermiteUsar = True
    Else
        IntervaloPermiteUsar = False
        
        UserList(UserIndex).Counters.failedUsageAttempts = UserList(UserIndex).Counters.failedUsageAttempts + 1
        
        'Tolerancia arbitraria - 20 es MUY alta, la está chiteando zarpado
        If UserList(UserIndex).Counters.failedUsageAttempts = 10 Then
            'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[ANTI-CHEAT]: " & UserList(UserIndex).Name & " con IP: " & UserList(UserIndex).Ip & " estuvo alterando el intervalo 'IntervaloPermiteUsar'", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
            UserList(UserIndex).Counters.failedUsageAttempts = 0

        End If

    End If

    '<EhFooter>
    Exit Function

IntervaloPermiteUsar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.IntervaloPermiteUsar " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function IntervaloPermiteUsarArcos(ByVal UserIndex As Integer, _
                                          Optional ByVal Actualizar As Boolean = True) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo IntervaloPermiteUsarArcos_Err

    '</EhHeader>

    Dim TActual As Long
    
    TActual = GetTime
    
    If TActual - UserList(UserIndex).Counters.TimerPuedeUsarArco >= IntervaloFlechasCazadores Then
        If Actualizar Then UserList(UserIndex).Counters.TimerPuedeUsarArco = TActual
        IntervaloPermiteUsarArcos = True
    Else
        IntervaloPermiteUsarArcos = False

    End If

    '<EhFooter>
    Exit Function

IntervaloPermiteUsarArcos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.IntervaloPermiteUsarArcos " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function IntervaloPermiteSerAtacado(ByVal UserIndex As Integer, _
                                           Optional ByVal Actualizar As Boolean = False) As Boolean

    '<EhHeader>
    On Error GoTo IntervaloPermiteSerAtacado_Err

    '</EhHeader>

    '**************************************************************
    'Author: ZaMa
    'Last Modify by: ZaMa
    'Last Modify Date: 13/11/2009
    '13/11/2009: ZaMa - Add the Timer which determines wether the user can be atacked by a NPc or not
    '**************************************************************
    Dim TActual As Long
    
    TActual = GetTime
    
    With UserList(UserIndex)

        ' Inicializa el timer
        If Actualizar Then
            .Counters.TimerPuedeSerAtacado = TActual
            .flags.NoPuedeSerAtacado = True
            IntervaloPermiteSerAtacado = False
        Else

            If TActual - .Counters.TimerPuedeSerAtacado >= IntervaloPuedeSerAtacado Then
                .flags.NoPuedeSerAtacado = False
                IntervaloPermiteSerAtacado = True
            Else
                IntervaloPermiteSerAtacado = False

            End If

        End If

    End With

    '<EhFooter>
    Exit Function

IntervaloPermiteSerAtacado_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.IntervaloPermiteSerAtacado " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function IntervaloPerdioNpc(ByVal UserIndex As Integer, _
                                   Optional ByVal Actualizar As Boolean = False) As Boolean

    '<EhHeader>
    On Error GoTo IntervaloPerdioNpc_Err

    '</EhHeader>

    '**************************************************************
    'Author: ZaMa
    'Last Modify by: ZaMa
    'Last Modify Date: 13/11/2009
    '13/11/2009: ZaMa - Add the Timer which determines wether the user still owns a Npc or not
    '**************************************************************
    Dim TActual As Long
    
    TActual = GetTime
    
    With UserList(UserIndex)

        ' Inicializa el timer
        If Actualizar Then
            .Counters.TimerPerteneceNpc = TActual
            IntervaloPerdioNpc = False
        Else

            If TActual - .Counters.TimerPerteneceNpc >= IntervaloOwnedNpc Then
                IntervaloPerdioNpc = True
            Else
                IntervaloPerdioNpc = False

            End If

        End If

    End With

    '<EhFooter>
    Exit Function

IntervaloPerdioNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.IntervaloPerdioNpc " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function IntervaloGoHome(ByVal UserIndex As Integer, _
                                Optional ByVal TimeInterval As Long, _
                                Optional ByVal Actualizar As Boolean = False) As Boolean

    '<EhHeader>
    On Error GoTo IntervaloGoHome_Err

    '</EhHeader>

    '**************************************************************
    'Author: ZaMa
    'Last Modify by: ZaMa
    'Last Modify Date: 01/06/2010
    '01/06/2010: ZaMa - Add the Timer which determines wether the user can be teleported to its home or not
    '**************************************************************
    Dim TActual As Long
    
    TActual = GetTime
    
    With UserList(UserIndex)

        ' Inicializa el timer
        If Actualizar Then
            .flags.Traveling = 1
            .Counters.goHome = TActual + TimeInterval
        Else

            If TActual >= .Counters.goHome Then
                IntervaloGoHome = True
                Call WriteUpdateGlobalCounter(UserIndex, 4, 0)

            End If

        End If

    End With

    '<EhFooter>
    Exit Function

IntervaloGoHome_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.IntervaloGoHome " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function IntervaloPermiteUsarClick(ByVal UserIndex As Integer, _
                                          Optional ByVal Actualizar As Boolean = True) As Boolean

    '<EhHeader>
    On Error GoTo IntervaloPermiteUsarClick_Err

    '</EhHeader>

    Dim TActual As Long

    With UserList(UserIndex).Counters
        TActual = GetTime()

        If (TActual - UserList(UserIndex).Counters.TimerUsarClick) >= IntervaloUserPuedeUsarClick Then
            If Actualizar Then
                '.TimerUsar = TActual
                .TimerUsarClick = TActual

            End If

            IntervaloPermiteUsarClick = True
        Else
            IntervaloPermiteUsarClick = False
            
            UserList(UserIndex).Counters.failedUsageAttempts_Clic = UserList(UserIndex).Counters.failedUsageAttempts_Clic + 1
        
            'Tolerancia arbitraria - 20 es MUY alta, la está chiteando zarpado
            If UserList(UserIndex).Counters.failedUsageAttempts_Clic = 10 Then
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[ANTI-CHEAT]: " & UserList(UserIndex).Name & " con IP: " & UserList(UserIndex).Account.Sec.IP_Address & " estuvo alterando el intervalo 'IntervaloPermiteUsar'", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
                UserList(UserIndex).Counters.failedUsageAttempts_Clic = 0

            End If
            
        End If

    End With

    '<EhFooter>
    Exit Function

IntervaloPermiteUsarClick_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.IntervaloPermiteUsarClick " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function Interval_Drop(ByVal UserIndex As Integer, _
                              Optional ByVal Actualizar As Boolean = True) As Boolean

    '<EhHeader>
    On Error GoTo Interval_Drop_Err

    '</EhHeader>

    Dim TActual As Long

    TActual = GetTime()

    If TActual - UserList(UserIndex).Counters.TimeDrop >= IntervalDrop Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimeDrop = TActual

        End If

        Interval_Drop = True
    Else
        Interval_Drop = False

    End If

    '<EhFooter>
    Exit Function

Interval_Drop_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.Interval_Drop " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function Interval_InfoChar(ByVal UserIndex As Integer, _
                                  Optional ByVal Actualizar As Boolean = True) As Boolean

    '<EhHeader>
    On Error GoTo Interval_InfoChar_Err

    '</EhHeader>

    Dim TActual As Long

    TActual = GetTime()

    If TActual - UserList(UserIndex).Counters.TimeInfoChar >= 10000 Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimeInfoChar = TActual

        End If

        Interval_InfoChar = True
    Else
        Interval_InfoChar = False

    End If

    '<EhFooter>
    Exit Function

Interval_InfoChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.Interval_InfoChar " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function Interval_Commerce(ByVal UserIndex As Integer, _
                                  Optional ByVal Actualizar As Boolean = True) As Boolean
        
    '<EhHeader>
    On Error GoTo Interval_Commerce_Err

    '</EhHeader>
        
    Dim TActual As Long

    TActual = GetTime()

    If TActual - UserList(UserIndex).Counters.TimeCommerce >= IntervalCommerce Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimeCommerce = TActual

        End If

        Interval_Commerce = True
    Else
        Interval_Commerce = False

    End If

    '<EhFooter>
    Exit Function

Interval_Commerce_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.Interval_Commerce " & "at line " & Erl
        
End Function

Public Function Interval_Message(ByVal UserIndex As Integer, _
                                 Optional ByVal Actualizar As Boolean = True) As Boolean
    
    '<EhHeader>
    On Error GoTo Interval_Message_Err

    '</EhHeader>
        
    Dim TActual As Long

    TActual = GetTime()

    If TActual - UserList(UserIndex).Counters.TimeMessage >= IntervalMessage Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimeMessage = TActual

        End If

        Interval_Message = True
    Else
        Interval_Message = False

    End If

    '<EhFooter>
    Exit Function

Interval_Message_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.Interval_Message " & "at line " & Erl
        
End Function

Public Function Interval_Packet250(ByVal UserIndex As Integer, _
                                   Optional ByVal Actualizar As Boolean = True) As Boolean

    '<EhHeader>
    On Error GoTo Interval_Packet250_Err

    '</EhHeader>

    Dim TActual As Long

    TActual = GetTime()

    If TActual - UserList(UserIndex).Counters.Packet250 >= 250 Then
        If Actualizar Then
            UserList(UserIndex).Counters.Packet250 = TActual

        End If

        Interval_Packet250 = True
    Else
        Interval_Packet250 = False

    End If

    '<EhFooter>
    Exit Function

Interval_Packet250_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.Interval_Packet250 " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function Interval_Packet500(ByVal UserIndex As Integer, _
                                   Optional ByVal Actualizar As Boolean = True) As Boolean

    '<EhHeader>
    On Error GoTo Interval_Packet500_Err

    '</EhHeader>

    Dim TActual As Long

    TActual = GetTime()

    If TActual - UserList(UserIndex).Counters.Packet500 >= 500 Then
        If Actualizar Then
            UserList(UserIndex).Counters.Packet500 = TActual

        End If

        Interval_Packet500 = True
    Else
        Interval_Packet500 = False

    End If

    '<EhFooter>
    Exit Function

Interval_Packet500_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.Interval_Packet500 " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function Interval_Mao(ByVal UserIndex As Integer, _
                             Optional ByVal Actualizar As Boolean = True) As Boolean

    '<EhHeader>
    On Error GoTo Interval_Mao_Err

    '</EhHeader>

    Dim TActual As Long

    TActual = GetTime

    If TActual - UserList(UserIndex).Counters.TimeInfoMao >= IntervalInfoMao Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimeInfoMao = TActual

        End If

        Interval_Mao = True
    Else
        Interval_Mao = False

    End If

    '<EhFooter>
    Exit Function

Interval_Mao_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.Interval_Mao " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function Interval_Equipped(ByVal UserIndex As Integer, _
                                  Optional ByVal Actualizar As Boolean = True) As Boolean

    '<EhHeader>
    On Error GoTo Interval_Equipped_Err

    '</EhHeader>

    Dim TActual As Long

    TActual = GetTime

    If TActual - UserList(UserIndex).Counters.TimeEquipped >= IntervaloEquipped Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimeEquipped = TActual

        End If

        Interval_Equipped = True
    Else
        Interval_Equipped = False

    End If

    '<EhFooter>
    Exit Function

Interval_Equipped_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Interval_Equipped.Interval_Mao " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function checkInterval(ByRef startTime As Long, _
                              ByVal timeNow As Long, _
                              ByVal interval As Long) As Boolean

    '<EhHeader>
    On Error GoTo checkInterval_Err

    '</EhHeader>

    Dim lInterval As Long

    If timeNow < startTime Then
        lInterval = &H7FFFFFFF - startTime + timeNow + 1
    Else
        lInterval = timeNow - startTime

    End If

    If lInterval >= interval Then
        startTime = timeNow
        checkInterval = True
    Else
        checkInterval = False

    End If

    '<EhFooter>
    Exit Function

checkInterval_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.checkInterval " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function IntervaloPuedeRecibirAtaqueCriature(ByVal UserIndex As Integer, _
                                                    Optional ByVal Actualizar As Boolean = True) As Boolean

    '<EhHeader>
    On Error GoTo IntervaloPuedeRecibirAtaqueCriature_Err

    '</EhHeader>

    Dim TActual As Long

    If haciendoBK Then Exit Function

    TActual = GetTime

    With UserList(UserIndex).Counters

        If TActual - .TimerPuedeRecibirAtaqueCriature >= 800 Then
            If Actualizar Then
                .TimerPuedeRecibirAtaqueCriature = TActual

            End If

            IntervaloPuedeRecibirAtaqueCriature = True
        Else
            IntervaloPuedeRecibirAtaqueCriature = False

        End If

    End With

    '<EhFooter>
    Exit Function

IntervaloPuedeRecibirAtaqueCriature_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.IntervaloPuedeRecibirAtaqueCriature " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function IntervaloPermiteCastear(ByVal UserIndex As Integer, _
                                        Optional ByVal Actualizar As Boolean = True) As Boolean

    '<EhHeader>
    On Error GoTo IntervaloPermiteCastear_Err

    '</EhHeader>

    Dim TActual As Long

    If haciendoBK Then Exit Function

    TActual = GetTime

    With UserList(UserIndex).Counters

        If TActual - .TimerPuedeCastear >= IntervaloPuedeCastear Then
            If Actualizar Then
                .TimerPuedeCastear = TActual

            End If

            IntervaloPermiteCastear = True
        Else
            IntervaloPermiteCastear = False
            
            .failedUsageCastSpell = .failedUsageCastSpell + 1
        
            'Tolerancia arbitraria - 20 es MUY alta, la está chiteando zarpado
            If .failedUsageCastSpell = 10 Then
                'Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, UserList(UserIndex).Name & " con IP: " & UserList(UserIndex).Ip & " estuvo alterando el intervalo 'IntervaloPuedeCastear'")
                'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[ANTI-CHEAT]: " & UserList(UserIndex).Name & " con IP: " & UserList(UserIndex).Ip & " estuvo alterando el intervalo 'IntervaloPuedeCastear'", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
                .failedUsageCastSpell = 0

            End If

        End If

    End With

    '<EhFooter>
    Exit Function

IntervaloPermiteCastear_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.IntervaloPermiteCastear " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function Intervalo_BotUseItem(ByVal NpcIndex As Integer, _
                                     Optional ByVal Update As Boolean = True) As Boolean

    '<EhHeader>
    On Error GoTo Intervalo_BotUseItem_Err

    '</EhHeader>
    
    Dim TActual As Long

    TActual = GetTime
    
    With Npclist(NpcIndex)
    
        If TActual - .Contadores.UseItem >= BotIntelligence_Balance_UseItem(.Stats.Elv) Then
            If Update Then
                .Contadores.UseItem = TActual

            End If
            
            Intervalo_BotUseItem = True
        Else
            Intervalo_BotUseItem = False

        End If
        
    End With

    '<EhFooter>
    Exit Function

Intervalo_BotUseItem_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.Intervalo_BotUseItem " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function Intervalo_CriatureVelocity(ByVal NpcIndex As Integer, _
                                           Optional ByVal Update As Boolean = True) As Boolean

    '<EhHeader>
    On Error GoTo Intervalo_CriatureVelocity_Err

    '</EhHeader>
    
    Dim TActual As Long

    TActual = GetTime
    
    With Npclist(NpcIndex)
    
        If TActual - .Contadores.Velocity >= .Velocity Then
            If Update Then
                .Contadores.Velocity = TActual

            End If
            
            Intervalo_CriatureVelocity = True
        Else
            Intervalo_CriatureVelocity = False

        End If
        
    End With

    '<EhFooter>
    Exit Function

Intervalo_CriatureVelocity_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.Intervalo_CriatureVelocity " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function Intervalo_CriatureAttack(ByVal NpcIndex As Integer, _
                                         Optional ByVal Update As Boolean = True) As Boolean

    '<EhHeader>
    On Error GoTo Intervalo_CriatureAttack_Err

    '</EhHeader>
    
    Dim TActual As Long

    TActual = GetTime
    
    With Npclist(NpcIndex)
              
        If TActual - .Contadores.Attack >= .IntervalAttack Then
            If Update Then
                .Contadores.Attack = TActual

            End If
            
            Intervalo_CriatureAttack = True

        End If
        
    End With

    '<EhFooter>
    Exit Function

Intervalo_CriatureAttack_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.Intervalo_CriatureAttack " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function Intervalo_CriatureDescanso(ByVal NpcIndex As Integer, _
                                           Optional ByVal Update As Boolean = True) As Boolean

    '<EhHeader>
    On Error GoTo Intervalo_CriatureDescanso_Err

    '</EhHeader>
    
    Dim TActual As Long

    TActual = GetTime
    
    With Npclist(NpcIndex)
    
        If TActual - .Contadores.Descanso >= 30000 Then
            If Update Then
                .Contadores.Descanso = TActual

            End If
            
            Intervalo_CriatureDescanso = True

        End If
        
    End With

    '<EhFooter>
    Exit Function

Intervalo_CriatureDescanso_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.Intervalo_CriatureDescanso " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function Intervalo_CriatureMovimientoConstante(ByVal NpcIndex As Integer, _
                                                      Optional ByVal Update As Boolean = True) As Boolean

    '<EhHeader>
    On Error GoTo Intervalo_CriatureMovimientoConstante_Err

    '</EhHeader>
    
    Dim TActual As Long

    TActual = GetTime
    
    With Npclist(NpcIndex)
    
        If TActual - .Contadores.MovimientoConstante >= 10000 Then
            If Update Then
                .Contadores.MovimientoConstante = TActual

            End If
            
            Intervalo_CriatureMovimientoConstante = True

        End If
        
    End With

    '<EhFooter>
    Exit Function

Intervalo_CriatureMovimientoConstante_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modNuevoTimer.Intervalo_CriatureMovimientoConstante " & "at line " & Erl
        
    '</EhFooter>
End Function

