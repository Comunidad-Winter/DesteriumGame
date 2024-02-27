VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "diccionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' diccionario.cls
'
' Designed and implemented by Mariono Barrou (El Oso)
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

'clase diccionario
'basico, plain sin queso ni papa fritas


'mi idea cuando hice esto, lo encontre en el rigido :p. Hecha por el oso

Option Explicit

Private Const MAX_ELEM = 100

Private Type diccElem
    clave As String
    def As Variant
End Type

Private p_elementos(1 To MAX_ELEM) As diccElem 'visual basic es una mierda para usar memoria dinamica, asi que uso esto
Private p_cant As Integer

Public Sub Class_Initialize()
      'constructor
10        p_cant = 0
End Sub

Public Sub Class_Terminate()
'destructor
    'destruir los variants?????
End Sub

Public Property Get CantElem() As Integer
10        CantElem = p_cant
End Property

Public Function AtPut(ByVal clave As String, ByRef elem As Variant) As Boolean
      Dim i As Integer

10    AtPut = False

20    If LenB(clave) = 0 Then Exit Function

30    clave = UCase$(clave)

40    If p_cant = MAX_ELEM Then
50        AtPut = False
60    Else
70        For i = 1 To p_cant
80            If clave = p_elementos(i).clave Then
90                p_elementos(i).def = elem
100               AtPut = True
110               Exit For    ' epa ;)
120           End If
130       Next i
140       If Not AtPut Then
150           p_cant = p_cant + 1
160           p_elementos(p_cant).def = elem
170           p_elementos(p_cant).clave = clave
180           AtPut = True
190       End If

200   End If
End Function

Public Function At(ByVal clave As String) As Variant
      Dim i As Integer

10    clave = UCase$(clave)

20        At = Null
30        For i = 1 To p_cant
40            If clave = p_elementos(i).clave Then
50                At = p_elementos(i).def
60                Exit Function
70            End If
80        Next i

End Function

Public Function AtIndex(ByVal i As Integer) As String
10        AtIndex = p_elementos(i).clave
End Function


Public Function MayorValor(ByRef cant As Integer) As String
      'parchecito para el AO, me da la clave con mayor valor en valor
      'y la cantidad de claves con ese valor (por si hay empate)
      Dim i As Integer
      Dim max As Integer
      Dim clave As String
10        max = -1
20        cant = 0
30        clave = vbNullString
40        For i = 1 To p_cant
50            If max <= CInt(p_elementos(i).def) Then
60                cant = IIf(max = CInt(p_elementos(i).def), cant + 1, 1)
70                clave = IIf(max = CInt(p_elementos(i).def), clave & "," & p_elementos(i).clave, p_elementos(i).clave)
80                max = CInt(p_elementos(i).def)
90            End If
100       Next i

110       MayorValor = clave

End Function

Public Sub DumpAll()
      Dim i As Integer

10    For i = 1 To MAX_ELEM
20        p_elementos(i).clave = vbNullString
30        p_elementos(i).def = Null
40    Next i
50    p_cant = 0

End Sub
