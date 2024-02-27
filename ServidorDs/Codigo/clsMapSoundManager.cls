VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoundMapInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' SoundMapInfo.cls
'
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

Private Type p_tSoundMapInfo
    Cantidad As Integer
    SoundIndex() As Integer
    flags() As Long
    Probabilidad() As Single
End Type

Private Enum p_eSoundFlags
    ninguna = 0
    Lluvia = 1
End Enum

Private p_Mapas() As p_tSoundMapInfo


'sonidos conocidos, pasados a enum para intelisense
Public Enum e_SoundIndex
    MUERTE_HOMBRE = 11
    MUERTE_MUJER = 74
    FLECHA_IMPACTO = 65
    CONVERSION_BARCO = 55
    MORFAR_MANZANA = 82
    SOUND_COMIDA = 7
    MUERTE_MUJER_AGUA = 211
    MUERTE_HOMBRE_AGUA = 212
End Enum

Private Sub Class_Initialize()
      'armar el array
10        ReDim p_Mapas(1 To NumMaps) As p_tSoundMapInfo
20        Call LoadSoundMapInfo
End Sub

Public Sub LoadSoundMapInfo()
          Dim i As Integer
          Dim j As Integer
          Dim Temps As String
          Dim MAPFILE As String
          
10        MAPFILE = App.Path & MapPath & "MAPA"
          
          'Usage of Val() prevents errors when dats are corrputed or incomplete. All invalid values are assumed to be zero.
          
      'TODO : Log the error in the dat for correction.
20        For i = 1 To UBound(p_Mapas)
30            Temps = GetVar(MAPFILE & i & ".dat", "SONIDOS", "Cantidad")
              
40            If IsNumeric(Temps) Then
50                p_Mapas(i).Cantidad = val(Temps)
                  
60                ReDim p_Mapas(i).flags(1 To p_Mapas(i).Cantidad) As Long
70                ReDim p_Mapas(i).Probabilidad(1 To p_Mapas(i).Cantidad) As Single
80                ReDim p_Mapas(i).SoundIndex(1 To p_Mapas(i).Cantidad) As Integer
                  
90                For j = 1 To p_Mapas(i).Cantidad
100                   p_Mapas(i).flags(j) = val(GetVar(MAPFILE & i & ".dat", "SONIDO" & j, "Flags"))
110                   p_Mapas(i).Probabilidad(j) = val(GetVar(MAPFILE & i & ".dat", "SONIDO" & j, "Probabilidad"))
120                   p_Mapas(i).SoundIndex(j) = val(GetVar(MAPFILE & i & ".dat", "SONIDO" & j, "Sonido"))
130               Next j
140           Else
150               p_Mapas(i).Cantidad = 0
160           End If
170       Next i
End Sub

Public Sub ReproducirSonidosDeMapas()
          Dim i As Long
          Dim SonidoMapa As Byte
          Dim posX As Byte
          Dim posY As Byte
          
          'Sounds are played at a random position
10        posX = RandomNumber(XMinMapSize, XMaxMapSize)
20        posY = RandomNumber(YMinMapSize, YMaxMapSize)
          
30        For i = 1 To UBound(p_Mapas)
40            If p_Mapas(i).Cantidad > 0 Then
50                SonidoMapa = RandomNumber(1, p_Mapas(i).Cantidad)
60                If RandomNumber(1, 100) <= p_Mapas(i).Probabilidad(SonidoMapa) Then
                      'tocarlo
70                    If Lloviendo Then
80                        If p_Mapas(i).flags(SonidoMapa) Xor p_eSoundFlags.Lluvia Then
90                            Call SendData(SendTarget.toMap, i, PrepareMessagePlayWave(p_Mapas(i).SoundIndex(SonidoMapa), posX, posY))
100                       End If
110                   Else
120                       If p_Mapas(i).flags(SonidoMapa) Xor p_eSoundFlags.ninguna Then
130                           Call SendData(SendTarget.toMap, i, PrepareMessagePlayWave(p_Mapas(i).SoundIndex(SonidoMapa), posX, posY))
140                       End If
150                   End If
160               End If
170           End If
180       Next i
End Sub

Public Sub ReproducirSonido(ByVal Destino As SendTarget, ByVal index As Integer, ByVal SoundIndex As Integer)
10        Call SendData(Destino, index, PrepareMessagePlayWave(SoundIndex, UserList(index).Pos.X, UserList(index).Pos.Y))
End Sub

