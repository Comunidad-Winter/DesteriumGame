Attribute VB_Name = "Queue"
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

Public Type tVertice
    X As Integer
    Y As Integer
End Type

Private Const MAXELEM As Integer = 1000

Private m_array() As tVertice
Private m_lastelem As Integer
Private m_firstelem As Integer
Private m_size As Integer

Public Function IsEmpty() As Boolean
10    IsEmpty = m_size = 0
End Function

Public Function IsFull() As Boolean
10    IsFull = m_lastelem = MAXELEM
End Function

Public Function Push(ByRef Vertice As tVertice) As Boolean

10    If Not IsFull Then
          
20        If IsEmpty Then m_firstelem = 1
          
30        m_lastelem = m_lastelem + 1
40        m_size = m_size + 1
50        m_array(m_lastelem) = Vertice
          
60        Push = True
70    Else
80        Push = False
90    End If

End Function

Public Function Pop() As tVertice

10    If Not IsEmpty Then
          
20        Pop = m_array(m_firstelem)
30        m_firstelem = m_firstelem + 1
40        m_size = m_size - 1
          
50        If m_firstelem > m_lastelem And m_size = 0 Then
60                m_lastelem = 0
70                m_firstelem = 0
80                m_size = 0
90        End If
         
100   End If

End Function

Public Sub InitQueue()
10    ReDim m_array(MAXELEM) As tVertice
20    m_lastelem = 0
30    m_firstelem = 0
40    m_size = 0
End Sub

