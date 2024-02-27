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
Public Function IntervaloPermiteLanzarSpell(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim TActual As Long

10    TActual = GetTickCount() And &H7FFFFFFF

20    If TActual - UserList(UserIndex).Counters.TimerLanzarSpell >= IntervaloUserPuedeCastear Then
30        If Actualizar Then
40            UserList(UserIndex).Counters.TimerLanzarSpell = TActual
50        End If
60        IntervaloPermiteLanzarSpell = True
70    Else
80        IntervaloPermiteLanzarSpell = False
90    End If

End Function

Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim TActual As Long

10    TActual = GetTickCount() And &H7FFFFFFF

20    If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= IntervaloUserPuedeAtacar Then
30        If Actualizar Then
40            UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
50            UserList(UserIndex).Counters.TimerGolpeUsar = TActual
60        End If
70        IntervaloPermiteAtacar = True
80    Else
90        IntervaloPermiteAtacar = False
100   End If
End Function

Public Function IntervaloPermiteGolpeUsar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
      '***************************************************
      'Author: ZaMa
      'Checks if the time that passed from the last hit is enough for the user to use a potion.
      'Last Modification: 06/04/2009
      '***************************************************

      Dim TActual As Long

10    TActual = GetTickCount() And &H7FFFFFFF

20    If TActual - UserList(UserIndex).Counters.TimerGolpeUsar >= IntervaloGolpeUsar Then
30        If Actualizar Then
40            UserList(UserIndex).Counters.TimerGolpeUsar = TActual
50        End If
60        IntervaloPermiteGolpeUsar = True
70    Else
80        IntervaloPermiteGolpeUsar = False
90    End If
End Function

Public Function IntervaloPermiteMagiaGolpe(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************
          Dim TActual As Long
          
10        With UserList(UserIndex)
20            If .Counters.TimerMagiaGolpe > .Counters.TimerLanzarSpell Then
30                Exit Function
40            End If
              
50            TActual = GetTickCount() And &H7FFFFFFF
              
60            If TActual - .Counters.TimerLanzarSpell >= IntervaloMagiaGolpe Then
70                If Actualizar Then
80                    .Counters.TimerMagiaGolpe = TActual
90                    .Counters.TimerPuedeAtacar = TActual
100                   .Counters.TimerGolpeUsar = TActual
110               End If
120               IntervaloPermiteMagiaGolpe = True
130           Else
140               IntervaloPermiteMagiaGolpe = False
150           End If
160       End With
End Function

Public Function IntervaloPermiteGolpeMagia(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim TActual As Long
          
10        If UserList(UserIndex).Counters.TimerGolpeMagia > UserList(UserIndex).Counters.TimerPuedeAtacar Then
20            Exit Function
30        End If
          
40        TActual = GetTickCount() And &H7FFFFFFF
          
50        If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= IntervaloGolpeMagia Then
60            If Actualizar Then
70                UserList(UserIndex).Counters.TimerGolpeMagia = TActual
80                UserList(UserIndex).Counters.TimerLanzarSpell = TActual
90            End If
100           IntervaloPermiteGolpeMagia = True
110       Else
120           IntervaloPermiteGolpeMagia = False
130       End If
End Function

' ATAQUE CUERPO A CUERPO
'Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'Dim TActual As Long
'
'TActual = GetTickCount() And &H7FFFFFFF''
'
'If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= IntervaloUserPuedeAtacar Then
'    If Actualizar Then UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
'    IntervaloPermiteAtacar = True
'Else
'    IntervaloPermiteAtacar = False
'End If
'End Function

' TRABAJO
Public Function IntervaloPermiteTrabajar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim TActual As Long
          
10        TActual = GetTickCount() And &H7FFFFFFF
          
20        If TActual - UserList(UserIndex).Counters.TimerPuedeTrabajar >= IntervaloUserPuedeTrabajar Then
30            If Actualizar Then UserList(UserIndex).Counters.TimerPuedeTrabajar = TActual
40            IntervaloPermiteTrabajar = True
50        Else
60            IntervaloPermiteTrabajar = False
70        End If
End Function

Private Function getInterval(ByVal TimeNow As Long, _
                             ByVal StartTime As Long) As Long ' 0.13.5

10        If TimeNow < StartTime Then
20            getInterval = &H7FFFFFFF - StartTime + TimeNow + 1
30        Else
40            getInterval = TimeNow - StartTime

50        End If

End Function

Public Function CheckInterval(ByVal StartTime As Long, _
                              ByVal TimeNow As Long, _
                              ByVal Interval As Long) As Boolean
          Dim lInterval As Long

10        lInterval = getInterval(TimeNow, StartTime)

20        If lInterval >= Interval Then
30            CheckInterval = True
40        Else
50            CheckInterval = False

60        End If

End Function

Public Function IntervaloPermiteUsarClick(ByVal UserIndex As Integer, _
                                     Optional ByVal Actualizar As Boolean = True) As Boolean

          Dim TActual As Long

10        With UserList(UserIndex).Counters
20            TActual = GetTickCount() And &H7FFFFFFF
          
30            If CheckInterval(.TimerUsarClick, TActual, (IntervaloUserPuedeUsar / 2)) Then
          
40                If Actualizar Then
50                    .TimerUsar = TActual
60                    .TimerUsarClick = TActual
70                    .failedUsageAttempts = 0

80                End If

90                IntervaloPermiteUsarClick = True
100           Else
110               IntervaloPermiteUsarClick = False
              
120               .failedUsageAttempts = .failedUsageAttempts + 1
              
130               If .failedUsageAttempts = 8 Then
                      Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("ANTICHEAT > posible modificación de intervalos por parte de " & UserList(UserIndex).Name & " Hora: " & time$, FontTypeNames.FONTTYPE_EJECUCION))
140                   .failedUsageAttempts = 0

150               End If

160           End If

170       End With

End Function

' USAR OBJETOS
Public Function IntervaloPermiteUsar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: 25/01/2010 (ZaMa)
      '25/01/2010: ZaMa - General adjustments.
      '***************************************************

          Dim TActual As Long
          
10        TActual = GetTickCount() And &H7FFFFFFF
          
20        If TActual - UserList(UserIndex).Counters.TimerUsar >= IntervaloUserPuedeUsar Then
30            If Actualizar Then
40                UserList(UserIndex).Counters.TimerUsar = TActual
50                UserList(UserIndex).Counters.TimerUsarClick = TActual
60                UserList(UserIndex).Counters.failedUsageAttempts = 0
70            End If
80            IntervaloPermiteUsar = True
90        Else
100           IntervaloPermiteUsar = False
              
110        UserList(UserIndex).Counters.failedUsageAttempts = UserList(UserIndex).Counters.failedUsageAttempts + 1
              'Tolerancia arbitraria - 20 es MUY alta, la está chiteando zarpado
120           If UserList(UserIndex).Counters.failedUsageAttempts >= 8 Then
                  Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("ANTICHEAT > posible modificación de intervalos por parte de " & UserList(UserIndex).Name & " Hora: " & time$, FontTypeNames.FONTTYPE_EJECUCION))
130               UserList(UserIndex).Counters.failedUsageAttempts = 0

140           End If
150       End If

End Function

Public Function IntervaloPermiteUsarArcos(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim TActual As Long
          
10        TActual = GetTickCount() And &H7FFFFFFF
          
20        If TActual - UserList(UserIndex).Counters.TimerPuedeUsarArco >= IntervaloFlechasCazadores Then
30            If Actualizar Then UserList(UserIndex).Counters.TimerPuedeUsarArco = TActual
40            IntervaloPermiteUsarArcos = True
50        Else
60            IntervaloPermiteUsarArcos = False
70        End If

End Function

Public Function IntervaloPermiteSerAtacado(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = False) As Boolean
      '**************************************************************
      'Author: ZaMa
      'Last Modify by: ZaMa
      'Last Modify Date: 13/11/2009
      '13/11/2009: ZaMa - Add the Timer which determines wether the user can be atacked by a NPc or not
      '**************************************************************
          Dim TActual As Long
          
10        TActual = GetTickCount() And &H7FFFFFFF
          
20        With UserList(UserIndex)
              ' Inicializa el timer
30            If Actualizar Then
40                .Counters.TimerPuedeSerAtacado = TActual
50                .flags.NoPuedeSerAtacado = True
60                IntervaloPermiteSerAtacado = False
70            Else
80                If TActual - .Counters.TimerPuedeSerAtacado >= IntervaloPuedeSerAtacado Then
90                    .flags.NoPuedeSerAtacado = False
100                   IntervaloPermiteSerAtacado = True
110               Else
120                   IntervaloPermiteSerAtacado = False
130               End If
140           End If
150       End With

End Function

Public Function IntervaloPerdioNpc(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = False) As Boolean
      '**************************************************************
      'Author: ZaMa
      'Last Modify by: ZaMa
      'Last Modify Date: 13/11/2009
      '13/11/2009: ZaMa - Add the Timer which determines wether the user still owns a Npc or not
      '**************************************************************
          Dim TActual As Long
          
10        TActual = GetTickCount() And &H7FFFFFFF
          
20        With UserList(UserIndex)
              ' Inicializa el timer
30            If Actualizar Then
40                .Counters.TimerPerteneceNpc = TActual
50                IntervaloPerdioNpc = False
60            Else
70                If TActual - .Counters.TimerPerteneceNpc >= IntervaloOwnedNpc Then
80                    IntervaloPerdioNpc = True
90                Else
100                   IntervaloPerdioNpc = False
110               End If
120           End If
130       End With

End Function

Public Function IntervaloEstadoAtacable(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = False) As Boolean
      '**************************************************************
      'Author: ZaMa
      'Last Modify by: ZaMa
      'Last Modify Date: 13/01/2010
      '13/01/2010: ZaMa - Add the Timer which determines wether the user can be atacked by an user or not
      '**************************************************************
          Dim TActual As Long
          
10        TActual = GetTickCount() And &H7FFFFFFF
          
20        With UserList(UserIndex)
              ' Inicializa el timer
30            If Actualizar Then
40                .Counters.TimerEstadoAtacable = TActual
50                IntervaloEstadoAtacable = True
60            Else
70                If TActual - .Counters.TimerEstadoAtacable >= IntervaloAtacable Then
80                    IntervaloEstadoAtacable = False
90                Else
100                   IntervaloEstadoAtacable = True
110               End If
120           End If
130       End With

End Function
