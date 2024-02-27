Attribute VB_Name = "ModAreas"
'Desterium AO 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'Desterium AO is based on Baronsoft's VB6 Online RPG
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

'LAS GUARDAMOS PARA PROCESAR LOS MPs y sabes si borrar personajes
Public MinLimiteX As Integer
Public MaxLimiteX As Integer
Public MinLimiteY As Integer
Public MaxLimiteY As Integer

Public Sub CambioDeArea(ByVal X As Byte, ByVal Y As Byte)
          Dim loopX As Long, LoopY As Long
          
10        MinLimiteX = (X \ 9 - 1) * 9
20        MaxLimiteX = MinLimiteX + 26
          
30        MinLimiteY = (Y \ 9 - 1) * 9
40        MaxLimiteY = MinLimiteY + 26
          
50        For loopX = 1 To 100
60            For LoopY = 1 To 100
                  
70                If (LoopY < MinLimiteY) Or (LoopY > MaxLimiteY) Or (loopX < _
                      MinLimiteX) Or (loopX > MaxLimiteX) Then
                      'Erase NPCs
                      
80                    If MapData(loopX, LoopY).CharIndex > 0 Then
90                        If MapData(loopX, LoopY).CharIndex <> UserCharIndex Then
100                           Call EraseChar(MapData(loopX, LoopY).CharIndex)
110                       End If
120                   End If
                      
                      
                      If MapData(loopX, LoopY).ObjGrh.GrhIndex > 0 Then
                      
                        #If Wgl = 1 Then
                            With GrhData(MapData(loopX, LoopY).ObjGrh.GrhIndex)
                                Call g_Swarm.Remove(4, loopX, LoopY, .TileWidth, .TileHeight)
                            End With
                        #End If
                        
                        'Erase OBJs
130                     MapData(loopX, LoopY).ObjGrh.GrhIndex = 0
                      End If
140               End If
150           Next LoopY
160       Next loopX
          
170       Call RefreshAllChars
End Sub
