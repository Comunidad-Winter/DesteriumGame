VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCola"
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
'                    Metodos publicos
'
' Public sub Push(byval i as variant) mete el elemento i
' al final de la cola.
'
' Public Function Pop As Variant: quita de la cola el primer elem
' y lo devuelve
'
' Public Function VerElemento(ByVal Index As Integer) As Variant
' muestra el elemento numero Index de la cola sin quitarlo
'
' Public Function PopByVal() As Variant: muestra el primer
' elemento de la cola sin quitarlo
'
' Public Property Get Longitud() As Integer: devuelve la
' cantidad de elementos que tiene la cola.

Private Const FRENTE = 1

Private Cola As Collection

Public Sub Reset()
10    On Error Resume Next

      Dim i As Integer
20    For i = 1 To Me.Longitud
30        Cola.Remove FRENTE
40    Next i

End Sub

Public Property Get Longitud() As Integer
10    Longitud = Cola.Count
End Property

Private Function IndexValido(ByVal i As Integer) As Boolean
10    IndexValido = i >= 1 And i <= Me.Longitud
End Function

Private Sub Class_Initialize()
10    Set Cola = New Collection
End Sub

Public Function VerElemento(ByVal index As Integer) As String
10    On Error Resume Next
20        If IndexValido(index) Then
              'Pablo
30            VerElemento = UCase$(Cola.Item(index))
              '/Pablo
              'VerElemento = Cola(Index)
40        Else
50            VerElemento = 0
60        End If
End Function


Public Sub Push(ByVal Nombre As String)
10    On Error Resume Next
      'Mete elemento en la cola
      'Pablo
      Dim aux As String
20    aux = time$ + " " + UCase$(Nombre)
30    Call Cola.Add(aux)
      '/Pablo

      'Call Cola.Add(UCase$(Nombre))
End Sub

Public Function Pop() As String
10    On Error Resume Next
      'Quita elemento de la cola
20    If Cola.Count > 0 Then
30        Pop = Cola(FRENTE)
40        Call Cola.Remove(FRENTE)
50    Else
60        Pop = 0
70    End If
End Function

Public Function PopByVal() As String
10    On Error Resume Next
      'Call LogTarea("PopByVal SOS")

      'Quita elemento de la cola
20    If Cola.Count > 0 Then
30        PopByVal = Cola.Item(1)
40    Else
50        PopByVal = 0
60    End If

End Function

Public Function Existe(ByVal Nombre As String) As Boolean
10    On Error Resume Next
        
       Dim V As String
       Dim i As Integer
       Dim NombreEnMayusculas As String
20     NombreEnMayusculas = UCase$(Nombre)
       
30     For i = 1 To Me.Longitud
              'Pablo
40            V = mid$(Me.VerElemento(i), 10, Len(Me.VerElemento(i)))
              '/Pablo
              'V = Me.VerElemento(i)
50            If V = NombreEnMayusculas Then
60                Existe = True
70                Exit Function
80            End If
90     Next
100    Existe = False

End Function

Public Sub Quitar(ByVal Nombre As String)
10    On Error Resume Next
      Dim V As String
      Dim i As Integer
      Dim NombreEnMayusculas As String
       
20    NombreEnMayusculas = UCase$(Nombre)
       
30    For i = 1 To Me.Longitud
          'Pablo
40        V = mid$(Me.VerElemento(i), 10, Len(Me.VerElemento(i)))
          '/Pablo
          'V = Me.VerElemento(i)
50        If V = NombreEnMayusculas Then
60            Call Cola.Remove(i)
70            Exit Sub
80        End If
90    Next i
       
End Sub

Public Sub QuitarIndex(ByVal index As Integer)
10    On Error Resume Next
20    If IndexValido(index) Then Call Cola.Remove(index)
End Sub


Private Sub Class_Terminate()
      'Destruimos el objeto Cola
10    Set Cola = Nothing
End Sub
