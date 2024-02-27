VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAntiMassClon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const MaximoPersonajesPorIP = 5
Private m_coleccion As New Collection

Public Function MaxPersonajes(sIp As String) As Boolean
          Dim i As Long

10        For i = 1 To m_coleccion.Count
20            If m_coleccion.Item(i).ip = sIp Then
30                m_coleccion.Item(i).PersonajesCreados = m_coleccion.Item(i).PersonajesCreados + 1
40                MaxPersonajes = (m_coleccion.Item(i).PersonajesCreados > MaximoPersonajesPorIP)
50                If MaxPersonajes Then m_coleccion.Item(i).PersonajesCreados = 6
60                Exit Function
70            End If
80        Next i

#If SeguridadAlkon Then
          Dim oIp As New UserIpAdress
90        oIp.ip = sIp
100       oIp.PersonajesCreados = 1
110       m_coleccion.Add oIp
#End If

120       MaxPersonajes = False
130       Exit Function
End Function

Public Function VaciarColeccion()

10    On Error GoTo Errhandler

      Dim i As Integer

20    For i = 1 To m_coleccion.Count
30       Call m_coleccion.Remove(1)
40    Next


50    Exit Function
Errhandler:
60        Call LogError("Error en RestarConexion " & Err.Description)
End Function


