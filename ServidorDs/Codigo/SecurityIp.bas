Attribute VB_Name = "SecurityIp"
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


'**************************************************************
' General_IpSecurity.Bas - Maneja la seguridad de las IPs
'
' Escrito y diseñado por DuNga (ltourrilhes@gmail.com)
'**************************************************************
Option Explicit

'*************************************************  *************
' General_IpSecurity.Bas - Maneja la seguridad de las IPs
'
' Escrito y diseñado por DuNga (ltourrilhes@gmail.com)
'*************************************************  *************

Private IpTables()      As Long 'USAMOS 2 LONGS: UNO DE LA IP, SEGUIDO DE UNO DE LA INFO
Private EntrysCounter   As Long
Private MaxValue        As Long
Private Multiplicado    As Long 'Cuantas veces multiplike el EntrysCounter para que me entren?
Private Const IntervaloEntreConexiones As Long = 1000

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Declaraciones para maximas conexiones por usuario
'Agregado por EL OSO
Private MaxConTables()      As Long
Private MaxConTablesEntry   As Long     'puntero a la ultima insertada

Private Const LIMITECONEXIONESxIP As Long = 10

Private Enum e_SecurityIpTabla
    IP_INTERVALOS = 1
    IP_LIMITECONEXIONES = 2
End Enum

Public Sub InitIpTables(ByVal OptCountersValue As Long)
      '*************************************************  *************
      'Author: Lucio N. Tourrilhes (DuNga)
      'Last Modify Date: EL OSO 21/01/06. Soporte para MaxConTables
      '
      '*************************************************  *************
10        EntrysCounter = OptCountersValue
20        Multiplicado = 1

30        ReDim IpTables(EntrysCounter * 2) As Long
40        MaxValue = 0

50        ReDim MaxConTables(Declaraciones.MaxUsers * 2 - 1) As Long
60        MaxConTablesEntry = 0

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''FUNCIONES PARA INTERVALOS'''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub IpSecurityMantenimientoLista()
      '*************************************************  *************
      'Author: Lucio N. Tourrilhes (DuNga)
      'Last Modify Date: Unknow
      '
      '*************************************************  *************
          'Las borro todas cada 1 hora, asi se "renuevan"
10        EntrysCounter = EntrysCounter \ Multiplicado
20        Multiplicado = 1
30        ReDim IpTables(EntrysCounter * 2) As Long
40        MaxValue = 0
End Sub

Public Function IpSecurityAceptarNuevaConexion(ByVal ip As Long) As Boolean
      '*************************************************  *************
      'Author: Lucio N. Tourrilhes (DuNga)
      'Last Modify Date: Unknow
      '
      '*************************************************  *************
      Dim IpTableIndex As Long
          

10        IpTableIndex = FindTableIp(ip, IP_INTERVALOS)
          
20        If IpTableIndex >= 0 Then
30            If IpTables(IpTableIndex + 1) + IntervaloEntreConexiones <= GetTickCount Then   'No está saturando de connects?
40                IpTables(IpTableIndex + 1) = GetTickCount
50                IpSecurityAceptarNuevaConexion = True
60                Debug.Print "CONEXION ACEPTADA"
70                Exit Function
80            Else
90                IpSecurityAceptarNuevaConexion = False

100               Debug.Print "CONEXION NO ACEPTADA"
110               Exit Function
120           End If
130       Else
140           IpTableIndex = Not IpTableIndex
150           AddNewIpIntervalo ip, IpTableIndex
160           IpTables(IpTableIndex + 1) = GetTickCount
170           IpSecurityAceptarNuevaConexion = True
180           Exit Function
190       End If

End Function


Private Sub AddNewIpIntervalo(ByVal ip As Long, ByVal index As Long)
      '*************************************************  *************
      'Author: Lucio N. Tourrilhes (DuNga)
      'Last Modify Date: Unknow
      '
      '*************************************************  *************
          '2) Pruebo si hay espacio, sino agrando la lista
10        If MaxValue + 1 > EntrysCounter Then
20            EntrysCounter = EntrysCounter \ Multiplicado
30            Multiplicado = Multiplicado + 1
40            EntrysCounter = EntrysCounter * Multiplicado
              
50            ReDim Preserve IpTables(EntrysCounter * 2) As Long
60        End If
          
          '4) Corro todo el array para arriba
70        Call CopyMemory(IpTables(index + 2), IpTables(index), (MaxValue - index \ 2) * 8)   '*4 (peso del long) * 2(cantidad de elementos por c/u)
80        IpTables(index) = ip
          
          '3) Subo el indicador de el maximo valor almacenado y listo :)
90        MaxValue = MaxValue + 1
End Sub

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ''''''''''''''''''''FUNCIONES PARA LIMITES X IP''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function IPSecuritySuperaLimiteConexiones(ByVal ip As Long) As Boolean
      Dim IpTableIndex As Long

10        IpTableIndex = FindTableIp(ip, IP_LIMITECONEXIONES)
          
20        If IpTableIndex >= 0 Then
              
30            If MaxConTables(IpTableIndex + 1) < LIMITECONEXIONESxIP Then
40                LogIP ("Agregamos conexion a " & ip & " iptableindex=" & IpTableIndex & ". Conexiones: " & MaxConTables(IpTableIndex + 1))
50                Debug.Print "suma conexion a " & ip & " total " & MaxConTables(IpTableIndex + 1) + 1
60                MaxConTables(IpTableIndex + 1) = MaxConTables(IpTableIndex + 1) + 1
70                IPSecuritySuperaLimiteConexiones = False
80            Else
90                LogIP ("rechazamos conexion de " & ip & " iptableindex=" & IpTableIndex & ". Conexiones: " & MaxConTables(IpTableIndex + 1))
100               Debug.Print "rechaza conexion a " & ip
110               IPSecuritySuperaLimiteConexiones = True
120           End If
130       Else
140           IPSecuritySuperaLimiteConexiones = False
150           If MaxConTablesEntry < Declaraciones.MaxUsers Then  'si hay espacio..
160               IpTableIndex = Not IpTableIndex
170               AddNewIpLimiteConexiones ip, IpTableIndex    'iptableindex es donde lo agrego
180               MaxConTables(IpTableIndex + 1) = 1
190           Else
200               Call LogCriticEvent("SecurityIP.IPSecuritySuperaLimiteConexiones: Se supero la disponibilidad de slots.")
210           End If
220       End If

End Function

Private Sub AddNewIpLimiteConexiones(ByVal ip As Long, ByVal index As Long)
      '*************************************************  *************
      'Author: (EL OSO)
      'Last Modify Date: Unknow
      '
      '*************************************************  *************
          'Debug.Print "agrega conexion a " & ip
          'Debug.Print "(Declaraciones.MaxUsers - index) = " & (Declaraciones.MaxUsers - Index)
          '4) Corro todo el array para arriba
          'Call CopyMemory(MaxConTables(Index + 2), MaxConTables(Index), (MaxConTablesEntry - Index \ 2) * 8)    '*4 (peso del long) * 2(cantidad de elementos por c/u)
          'MaxConTables(Index) = ip

          '3) Subo el indicador de el maximo valor almacenado y listo :)
          'MaxConTablesEntry = MaxConTablesEntry + 1


      '*************************************************    *************
      'Author: (EL OSO)
      'Last Modify Date: 16/2/2006
      'Modified by Juan Martín Sotuyo Dodero (Maraxus)
      '*************************************************    *************
10        Debug.Print "agrega conexion a " & ip
20        Debug.Print "(Declaraciones.MaxUsers - index) = " & (Declaraciones.MaxUsers - index)
30        Debug.Print "Agrega conexion a nueva IP " & ip
          '4) Corro todo el array para arriba
          Dim temp() As Long
40        ReDim temp((MaxConTablesEntry - index \ 2) * 2) As Long  'VB no deja inicializar con rangos variables...
50        Call CopyMemory(temp(0), MaxConTables(index), (MaxConTablesEntry - index \ 2) * 8)    '*4 (peso del long) * 2(cantidad de elementos por c/u)
60        Call CopyMemory(MaxConTables(index + 2), temp(0), (MaxConTablesEntry - index \ 2) * 8)    '*4 (peso del long) * 2(cantidad de elementos por c/u)
70        MaxConTables(index) = ip

          '3) Subo el indicador de el maximo valor almacenado y listo :)
80        MaxConTablesEntry = MaxConTablesEntry + 1

End Sub

Public Sub IpRestarConexion(ByVal ip As Long)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim key As Long
10        Debug.Print "resta conexion a " & ip
          
20        key = FindTableIp(ip, IP_LIMITECONEXIONES)
          
30        If key >= 0 Then
40            If MaxConTables(key + 1) > 0 Then
50                MaxConTables(key + 1) = MaxConTables(key + 1) - 1
60            End If
70            Call LogIP("restamos conexion a " & ip & " key=" & key & ". Conexiones: " & MaxConTables(key + 1))
80            If MaxConTables(key + 1) <= 0 Then
                  'la limpiamos
90                Call CopyMemory(MaxConTables(key), MaxConTables(key + 2), (MaxConTablesEntry - (key \ 2) + 1) * 8)
100               MaxConTablesEntry = MaxConTablesEntry - 1
110           End If
120       Else 'Key <= 0
130           Call LogIP("restamos conexion a " & ip & " key=" & key & ". NEGATIVO!!")
              'LogCriticEvent "SecurityIp.IpRestarconexion obtuvo un valor negativo en key"
140       End If
End Sub



' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ''''''''''''''''''''''''FUNCIONES GENERALES''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Function FindTableIp(ByVal ip As Long, ByVal Tabla As e_SecurityIpTabla) As Long
      '*************************************************  *************
      'Author: Lucio N. Tourrilhes (DuNga)
      'Last Modify Date: Unknow
      'Modified by Juan Martín Sotuyo Dodero (Maraxus) to use Binary Insertion
      '*************************************************  *************
      Dim First As Long
      Dim Last As Long
      Dim Middle As Long
          
10        Select Case Tabla
              Case e_SecurityIpTabla.IP_INTERVALOS
20                First = 0
30                Last = MaxValue
40                Do While First <= Last
50                    Middle = (First + Last) \ 2
                      
60                    If (IpTables(Middle * 2) < ip) Then
70                        First = Middle + 1
80                    ElseIf (IpTables(Middle * 2) > ip) Then
90                        Last = Middle - 1
100                   Else
110                       FindTableIp = Middle * 2
120                       Exit Function
130                   End If
140               Loop
150               FindTableIp = Not (Middle * 2)
              
160           Case e_SecurityIpTabla.IP_LIMITECONEXIONES
                  
170               First = 0
180               Last = MaxConTablesEntry

190               Do While First <= Last
200                   Middle = (First + Last) \ 2

210                   If MaxConTables(Middle * 2) < ip Then
220                       First = Middle + 1
230                   ElseIf MaxConTables(Middle * 2) > ip Then
240                       Last = Middle - 1
250                   Else
260                       FindTableIp = Middle * 2
270                       Exit Function
280                   End If
290               Loop
300               FindTableIp = Not (Middle * 2)
310       End Select
End Function

Public Function DumpTables()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim i As Integer

10        For i = 0 To MaxConTablesEntry * 2 - 1 Step 2
20            Call LogCriticEvent(GetAscIP(MaxConTables(i)) & " > " & MaxConTables(i + 1))
30        Next i

End Function

