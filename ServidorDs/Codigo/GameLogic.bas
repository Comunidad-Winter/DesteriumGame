Attribute VB_Name = "Extra"
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

Public Function EsNewbie(ByVal Userindex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        EsNewbie = UserList(Userindex).Stats.ELV <= LimiteNewbie
End Function
Public Function esArmada(ByVal Userindex As Integer) As Boolean
      '***************************************************
      'Autor: Pablo (ToxicWaste)
      'Last Modification: 23/01/2007
      '***************************************************

10        esArmada = (UserList(Userindex).Faccion.ArmadaReal = 1)
End Function

Public Function esCaos(ByVal Userindex As Integer) As Boolean
      '***************************************************
      'Autor: Pablo (ToxicWaste)
      'Last Modification: 23/01/2007
      '***************************************************

10        esCaos = (UserList(Userindex).Faccion.FuerzasCaos = 1)
End Function

Public Function EsGm(ByVal Userindex As Integer) As Boolean
      '***************************************************
      'Autor: Pablo (ToxicWaste)
      'Last Modification: 23/01/2007
      '***************************************************

10        EsGm = (UserList(Userindex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero))
End Function
Public Function EsVip(ByVal Userindex As Integer) As Boolean 'es un usuario vip?
10    EsVip = UserList(Userindex).flags.Oro
End Function
Public Function EsVipp(ByVal Userindex As Integer) As Boolean 'es un usuario vip?
10    EsVipp = UserList(Userindex).flags.Plata
End Function
Public Function EsVipb(ByVal Userindex As Integer) As Boolean 'es un usuario vip?
10    EsVipb = UserList(Userindex).flags.Bronce
End Function
Public Function EsCuarenta(ByVal Userindex As Integer) As Boolean
10        EsCuarenta = UserList(Userindex).Stats.ELV >= 40
End Function
Public Function EsSiete(ByVal Userindex As Integer) As Boolean
10        EsSiete = UserList(Userindex).Stats.ELV >= 47
End Function
Public Function EsOcho(ByVal Userindex As Integer) As Boolean
10        EsOcho = UserList(Userindex).Stats.ELV >= 48
End Function
Public Function EsNueve(ByVal Userindex As Integer) As Boolean
10        EsNueve = UserList(Userindex).Stats.ELV >= 49
End Function
Public Function EsQuince(ByVal Userindex As Integer) As Boolean
10        EsQuince = UserList(Userindex).Stats.ELV >= 15
End Function
Public Function EsVeinte(ByVal Userindex As Integer) As Boolean
10        EsVeinte = UserList(Userindex).Stats.ELV >= 20
End Function
Public Function EsVeinticinco(ByVal Userindex As Integer) As Boolean
10        EsVeinticinco = UserList(Userindex).Stats.ELV >= 25
End Function
Public Function EsQuinceM(ByVal Userindex As Integer) As Boolean
10        EsQuinceM = UserList(Userindex).Stats.ELV <= 15
End Function
Public Function EsTreintaM(ByVal Userindex As Integer) As Boolean
10        EsTreintaM = UserList(Userindex).Stats.ELV >= 13
End Function
Public Function EsHM(ByVal Userindex As Integer) As Boolean
10        EsHM = UserList(Userindex).Stats.ELV >= 30
End Function
Public Function EsUM(ByVal Userindex As Integer) As Boolean
10        EsUM = UserList(Userindex).Stats.ELV >= 35
End Function
Public Function EsMM(ByVal Userindex As Integer) As Boolean
10        EsMM = UserList(Userindex).Stats.ELV >= 45
End Function
Public Function NoEsUM(ByVal Userindex As Integer) As Boolean 'es un usuario vip?
10    NoEsUM = UserList(Userindex).flags.Oro <= 0
End Function

Public Sub DoTileEvents(ByVal Userindex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
      '***************************************************
      'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
      'Last Modification: 06/03/2010
      'Handles the Map passage of Users. Allows the existance
      'of exclusive maps for Newbies, Royal Army and Caos Legion members
      'and enables GMs to enter every map without restriction.
      'Uses: Mapinfo(map).Restringir = "NEWBIE" (newbies), "ARMADA", "CAOS", "FACCION" or "NO".
      ' 06/03/2010 : Now we have 5 attemps to not fall into a map change or another teleport while going into a teleport. (Marco)
      '***************************************************

          Dim nPos As WorldPos
          Dim FxFlag As Boolean
          Dim TelepRadio As Integer
          Dim DestPos As WorldPos
          
10    On Error GoTo Errhandler
          'Controla las salidas
20        If InMapBounds(map, X, Y) Then
30            With MapData(map, X, Y)
40                If .ObjInfo.ObjIndex > 0 Then
50                    FxFlag = ObjData(.ObjInfo.ObjIndex).ObjType = eOBJType.otTeleport
60                    TelepRadio = ObjData(.ObjInfo.ObjIndex).Radio
70                End If
                  
80                If .TileExit.map > 0 And .TileExit.map <= NumMaps Then
                      
                      ' Es un teleport, entra en una posicion random, acorde al radio (si es 0, es pos fija)
                      ' We have 5 attempts to not falling into another teleport or a map exit.. If we get to the fifth attemp,
                      ' the teleport will act as if its radius = 0.
90                    If FxFlag And TelepRadio > 0 Then
                          Dim attemps As Long
                          Dim exitMap As Boolean
100                       Do
110                           DestPos.X = .TileExit.X + RandomNumber(TelepRadio * (-1), TelepRadio)
120                           DestPos.Y = .TileExit.Y + RandomNumber(TelepRadio * (-1), TelepRadio)
                              
130                           attemps = attemps + 1
                              
140                           exitMap = MapData(.TileExit.map, DestPos.X, DestPos.Y).TileExit.map > 0 And _
                                      MapData(.TileExit.map, DestPos.X, DestPos.Y).TileExit.map <= NumMaps
150                       Loop Until (attemps >= 5 Or exitMap = False)
                          
160                       If attemps >= 5 Then
170                           DestPos.X = .TileExit.X
180                           DestPos.Y = .TileExit.Y
190                       End If
                      ' Posicion fija
200                   Else
210                       DestPos.X = .TileExit.X
220                       DestPos.Y = .TileExit.Y
230                   End If
                      
240                   DestPos.map = .TileExit.map
                      
250                   If EsGm(Userindex) Then
260                       Call LogGM(UserList(Userindex).Name, "Utilizó un teleport hacia el mapa " & _
                              DestPos.map & " (" & DestPos.X & "," & DestPos.Y & ")")
270                   End If
                      
                      '¿Es mapa de newbies?
280                   If UCase$(MapInfo(DestPos.map).Restringir) = "NEWBIE" Then
                          '¿El usuario es un newbie?
290                       If EsNewbie(Userindex) Or EsGm(Userindex) Then
300                           If LegalPos(DestPos.map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
310                               Call WarpUserChar(Userindex, DestPos.map, DestPos.X, DestPos.Y, FxFlag)
320                           Else
330                               Call ClosestLegalPos(DestPos, nPos)
340                               If nPos.X <> 0 And nPos.Y <> 0 Then
350                                   Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
360                               End If
370                           End If
380                       Else 'No es newbie
390                           Call WriteConsoleMsg(Userindex, "Mapa exclusivo para newbies.", FontTypeNames.FONTTYPE_INFO)
400                           Call ClosestStablePos(UserList(Userindex).Pos, nPos)
              
410                           If nPos.X <> 0 And nPos.Y <> 0 Then
420                               Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
430                           End If
440                       End If
450                   ElseIf UCase$(MapInfo(DestPos.map).Restringir) = "ARMADA" Then '¿Es mapa de Armadas?
                          '¿El usuario es Armada?
460                       If esArmada(Userindex) Or EsGm(Userindex) Then
470                           If LegalPos(DestPos.map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
480                               Call WarpUserChar(Userindex, DestPos.map, DestPos.X, DestPos.Y, FxFlag)
490                           Else
500                               Call ClosestLegalPos(DestPos, nPos)
510                               If nPos.X <> 0 And nPos.Y <> 0 Then
520                                   Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
530                               End If
540                           End If
550                       Else 'No es armada
560                           Call WriteConsoleMsg(Userindex, "Mapa exclusivo para miembros del ejército real.", FontTypeNames.FONTTYPE_INFO)
570                           Call ClosestStablePos(UserList(Userindex).Pos, nPos)
                              
580                           If nPos.X <> 0 And nPos.Y <> 0 Then
590                               Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
600                           End If
610                       End If
620                   ElseIf UCase$(MapInfo(DestPos.map).Restringir) = "CAOS" Then '¿Es mapa de Caos?
                          '¿El usuario es Caos?
630                       If esCaos(Userindex) Or EsGm(Userindex) Then
640                           If LegalPos(DestPos.map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
650                               Call WarpUserChar(Userindex, DestPos.map, DestPos.X, DestPos.Y, FxFlag)
660                           Else
670                               Call ClosestLegalPos(DestPos, nPos)
680                               If nPos.X <> 0 And nPos.Y <> 0 Then
690                                   Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
700                               End If
710                           End If
720                       Else 'No es caos
730                           Call WriteConsoleMsg(Userindex, "Mapa exclusivo para miembros de la legión oscura.", FontTypeNames.FONTTYPE_INFO)
740                           Call ClosestStablePos(UserList(Userindex).Pos, nPos)
                              
750                           If nPos.X <> 0 And nPos.Y <> 0 Then
760                               Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
770                           End If
780                       End If
790                   ElseIf UCase$(MapInfo(DestPos.map).Restringir) = "FACCION" Then '¿Es mapa de faccionarios?
                          '¿El usuario es Armada o Caos?
800                       If esArmada(Userindex) Or esCaos(Userindex) Or EsGm(Userindex) Then
810                           If LegalPos(DestPos.map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
820                               Call WarpUserChar(Userindex, DestPos.map, DestPos.X, DestPos.Y, FxFlag)
830                           Else
840                               Call ClosestLegalPos(DestPos, nPos)
850                               If nPos.X <> 0 And nPos.Y <> 0 Then
860                                   Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
870                               End If
880                           End If
890                       Else 'No es Faccionario
900                           Call WriteConsoleMsg(Userindex, "Solo se permite entrar al mapa si eres miembro de alguna facción.", FontTypeNames.FONTTYPE_INFO)
910                           Call ClosestStablePos(UserList(Userindex).Pos, nPos)
                              
920                           If nPos.X <> 0 And nPos.Y <> 0 Then
930                               Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
940                           End If
950                       End If
       
                          'QUince
                         
960       ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "CUARENTA" Then
970                              If UserList(Userindex).Stats.ELV >= 40 Then
980                               If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
990                                   Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
1000                              Else
1010                                  Call ClosestLegalPos(.TileExit, nPos)
1020                                 If nPos.X <> 0 And nPos.Y <> 0 Then
1030                                      Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
1040                                  End If
1050                             End If
1060                         Else
1070                             Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
1080                             Call ClosestStablePos(UserList(Userindex).Pos, nPos)
       
1090                             If nPos.X <> 0 And nPos.Y <> 0 Then
1100                                 Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
1110                             End If
1120                         End If
                             
                             
1130                 ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "MENOSCINCO" Then
1140                             If UserList(Userindex).Stats.ELV <= 45 Then
1150                              If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
1160                                  Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
1170                              Else
1180                                  Call ClosestLegalPos(.TileExit, nPos)
1190                                 If nPos.X <> 0 And nPos.Y <> 0 Then
1200                                      Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
1210                                  End If
1220                             End If
1230                         Else
1240                             Call WriteConsoleMsg(Userindex, "Tu nivel es muy elevado para ingresar en este mapa.", FontTypeNames.FONTTYPE_INFO)
1250                             Call ClosestStablePos(UserList(Userindex).Pos, nPos)
       
1260                             If nPos.X <> 0 And nPos.Y <> 0 Then
1270                                 Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
1280                             End If
1290                         End If
                             
1300                  ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "MENOSCUATRO" Then
1310                             If UserList(Userindex).Stats.ELV <= 40 Then
1320                              If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
1330                                  Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
1340                              Else
1350                                  Call ClosestLegalPos(.TileExit, nPos)
1360                                 If nPos.X <> 0 And nPos.Y <> 0 Then
1370                                      Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
1380                                  End If
1390                             End If
1400                         Else
1410                             Call WriteConsoleMsg(Userindex, "Tu nivel es muy elevado para ingresar en este mapa.", FontTypeNames.FONTTYPE_INFO)
1420                             Call ClosestStablePos(UserList(Userindex).Pos, nPos)
       
1430                             If nPos.X <> 0 And nPos.Y <> 0 Then
1440                                 Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
1450                             End If
1460                         End If
                             
1470                          ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "CINCO" Then
1480                             If UserList(Userindex).Stats.ELV >= 45 Then
1490                              If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
1500                                  Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
1510                              Else
1520                                  Call ClosestLegalPos(.TileExit, nPos)
1530                                 If nPos.X <> 0 And nPos.Y <> 0 Then
1540                                      Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
1550                                  End If
1560                             End If
1570                         Else
1580                             Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
1590                             Call ClosestStablePos(UserList(Userindex).Pos, nPos)
       
1600                             If nPos.X <> 0 And nPos.Y <> 0 Then
1610                                 Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
1620                             End If
1630                         End If
                             
1640                          ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "SEIS" Then
1650                             If UserList(Userindex).Stats.ELV >= 46 Then
1660                              If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
1670                                  Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
1680                              Else
1690                                  Call ClosestLegalPos(.TileExit, nPos)
1700                                 If nPos.X <> 0 And nPos.Y <> 0 Then
1710                                      Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
1720                                  End If
1730                             End If
1740                         Else
1750                             Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
1760                             Call ClosestStablePos(UserList(Userindex).Pos, nPos)
       
1770                             If nPos.X <> 0 And nPos.Y <> 0 Then
1780                                 Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
1790                             End If
1800                         End If
                          
1810          ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "SIETE" Then
1820                             If UserList(Userindex).Stats.ELV >= 47 Then
1830                              If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
1840                                  Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
1850                              Else
1860                                  Call ClosestLegalPos(.TileExit, nPos)
1870                                 If nPos.X <> 0 And nPos.Y <> 0 Then
1880                                      Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
1890                                  End If
1900                             End If
1910                         Else
1920                             Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
1930                             Call ClosestStablePos(UserList(Userindex).Pos, nPos)
       
1940                             If nPos.X <> 0 And nPos.Y <> 0 Then
1950                                 Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
1960                             End If
1970                         End If
                             
1980              ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "OCHO" Then
1990                             If UserList(Userindex).Stats.ELV >= 48 Then
2000                              If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
2010                                  Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
2020                              Else
2030                                  Call ClosestLegalPos(.TileExit, nPos)
2040                                 If nPos.X <> 0 And nPos.Y <> 0 Then
2050                                      Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
2060                                  End If
2070                             End If
2080                         Else
2090                             Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
2100                             Call ClosestStablePos(UserList(Userindex).Pos, nPos)
       
2110                             If nPos.X <> 0 And nPos.Y <> 0 Then
2120                                 Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
2130                             End If
2140                         End If
                             
                             
2150                  ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "NUEVE" Then
2160                             If UserList(Userindex).Stats.ELV >= 49 Then
2170                              If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
2180                                  Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
2190                              Else
2200                                  Call ClosestLegalPos(.TileExit, nPos)
2210                                 If nPos.X <> 0 And nPos.Y <> 0 Then
2220                                      Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
2230                                  End If
2240                             End If
2250                         Else
2260                             Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
2270                             Call ClosestStablePos(UserList(Userindex).Pos, nPos)
       
2280                             If nPos.X <> 0 And nPos.Y <> 0 Then
2290                                 Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
2300                             End If
2310                         End If
       
        'QUInce FIN
        
        
2320        ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "QUINCE" Then
2330                             If UserList(Userindex).Stats.ELV >= 15 Then
2340                              If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
2350                                  Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
2360                              Else
2370                                  Call ClosestLegalPos(.TileExit, nPos)
2380                                 If nPos.X <> 0 And nPos.Y <> 0 Then
2390                                      Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
2400                                  End If
2410                             End If
2420                         Else
2430                             Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
2440                             Call ClosestStablePos(UserList(Userindex).Pos, nPos)
       
2450                             If nPos.X <> 0 And nPos.Y <> 0 Then
2460                                 Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
2470                             End If
2480                         End If
        
        
2490          ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "VEINTE" Then
2500                             If UserList(Userindex).Stats.ELV >= 20 Then
2510                              If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
2520                                  Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
2530                              Else
2540                                  Call ClosestLegalPos(.TileExit, nPos)
2550                                 If nPos.X <> 0 And nPos.Y <> 0 Then
2560                                      Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
2570                                  End If
2580                             End If
2590                         Else
2600                             Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
2610                             Call ClosestStablePos(UserList(Userindex).Pos, nPos)
       
2620                             If nPos.X <> 0 And nPos.Y <> 0 Then
2630                                 Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
2640                             End If
2650                         End If
                             
2660                                 ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "VEINTICINCO" Then
2670                             If UserList(Userindex).Stats.ELV >= 25 Then
2680                              If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
2690                                  Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
2700                              Else
2710                                  Call ClosestLegalPos(.TileExit, nPos)
2720                                 If nPos.X <> 0 And nPos.Y <> 0 Then
2730                                      Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
2740                                  End If
2750                             End If
2760                         Else
2770                             Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
2780                             Call ClosestStablePos(UserList(Userindex).Pos, nPos)
       
2790                             If nPos.X <> 0 And nPos.Y <> 0 Then
2800                                 Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
2810                             End If
2820                         End If
        
       
       
2830   ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "VIP" Then
2840  If UserList(Userindex).flags.Oro = 1 Then
2850  If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
2860  Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
2870  Else
2880  Call ClosestLegalPos(.TileExit, nPos)
2890  If nPos.X <> 0 And nPos.Y <> 0 Then
2900  Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
2910  End If
2920  End If
2930  Else
2940  Call WriteConsoleMsg(Userindex, "Solo los usuarios Oro pueden ingresar a este mapa.", FontTypeNames.FONTTYPE_CONSE)
2950  Call ClosestStablePos(UserList(Userindex).Pos, nPos)

2960  If nPos.X <> 0 And nPos.Y <> 0 Then
2970  Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
2980  End If
2990  End If

3000   ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "PREMIUM" Then
3010  If UserList(Userindex).flags.Premium = 1 Then
3020  If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
3030  Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
3040  Else
3050  Call ClosestLegalPos(.TileExit, nPos)
3060  If nPos.X <> 0 And nPos.Y <> 0 Then
3070  Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
3080  End If
3090  End If
3100  Else
3110  Call WriteConsoleMsg(Userindex, "Sólo los usuarios PREMIUM pueden ingresar a este mapa.", FontTypeNames.FONTTYPE_CONSE)
3120  Call ClosestStablePos(UserList(Userindex).Pos, nPos)

3130  If nPos.X <> 0 And nPos.Y <> 0 Then
3140  Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
3150  End If
3160  End If

3170   ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "VIPP" Then
3180  If UserList(Userindex).flags.Plata = 1 Or UserList(Userindex).flags.Bronce = 1 Or UserList(Userindex).flags.Oro = 1 Then
3190  If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
3200  Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
3210  Else
3220  Call ClosestLegalPos(.TileExit, nPos)
3230  If nPos.X <> 0 And nPos.Y <> 0 Then
3240  Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
3250  End If
3260  End If
3270  Else
3280  Call WriteConsoleMsg(Userindex, "Sólo los usuarios Bronce, Plata y Oro pueden ingresar a este mapa.", FontTypeNames.FONTTYPE_CONSE)
3290  Call ClosestStablePos(UserList(Userindex).Pos, nPos)

3300  If nPos.X <> 0 And nPos.Y <> 0 Then
3310  Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
3320  End If
3330  End If

3340   ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "NOESUM" Then
3350  If UserList(Userindex).flags.Oro <= 0 Then
3360  If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
3370  Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
3380  Else
3390  Call ClosestLegalPos(.TileExit, nPos)
3400  If nPos.X <> 0 And nPos.Y <> 0 Then
3410  Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
3420  End If
3430  End If
3440  Else
3450  Call WriteConsoleMsg(Userindex, "Los usuarios Oro no pueden ingresar a este mapa.", FontTypeNames.FONTTYPE_CONSE)
3460  Call ClosestStablePos(UserList(Userindex).Pos, nPos)

3470  If nPos.X <> 0 And nPos.Y <> 0 Then
3480  Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
3490  End If
3500  End If

       
        'QUInce FIN
3510                      ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "PREMIUM" Then
3520  If UserList(Userindex).flags.Premium = 1 Then
3530  If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
3540  Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
3550  Else
3560  Call ClosestLegalPos(.TileExit, nPos)
3570  If nPos.X <> 0 And nPos.Y <> 0 Then
3580  Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
3590  End If
3600  End If
3610  Else
3620  Call WriteConsoleMsg(Userindex, "Para ingresar a este mapa tienes que ser un Usuario PREMIUM.", FontTypeNames.FONTTYPE_INFO)
3630  Call ClosestStablePos(UserList(Userindex).Pos, nPos)

3640  If nPos.X <> 0 And nPos.Y <> 0 Then
3650  Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
3660  End If
3670  End If
3680                  Else 'No es un mapa de newbies, ni Armadas, ni Caos, ni faccionario.
3690                      If LegalPos(DestPos.map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
3700                          Call WarpUserChar(Userindex, DestPos.map, DestPos.X, DestPos.Y, FxFlag)
3710                      Else
3720                          Call ClosestLegalPos(DestPos, nPos)
3730                          If nPos.X <> 0 And nPos.Y <> 0 Then
3740                              Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
3750                          End If
3760                      End If
3770                  End If
                      
                      'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
                      Dim aN As Integer
                      
3780                  aN = UserList(Userindex).flags.AtacadoPorNpc
3790                  If aN > 0 Then
3800                     Npclist(aN).Movement = Npclist(aN).flags.OldMovement
3810                     Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
3820                     Npclist(aN).flags.AttackedBy = vbNullString
3830                  End If
                  
3840                  aN = UserList(Userindex).flags.NPCAtacado
3850                  If aN > 0 Then
3860                      If Npclist(aN).flags.AttackedFirstBy = UserList(Userindex).Name Then
3870                          Npclist(aN).flags.AttackedFirstBy = vbNullString
3880                      End If
3890                  End If
3900                  UserList(Userindex).flags.AtacadoPorNpc = 0
3910                  UserList(Userindex).flags.NPCAtacado = 0
3920              End If
3930          End With
3940      End If
3950  Exit Sub

Errhandler:
3960      Call LogError("Error en DotileEvents. Error: " & Err.Number & " - Desc: " & Err.Description)
End Sub

Function InRangoVision(ByVal Userindex As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If X > UserList(Userindex).Pos.X - MinXBorder And X < UserList(Userindex).Pos.X + MinXBorder Then
20            If Y > UserList(Userindex).Pos.Y - MinYBorder And Y < UserList(Userindex).Pos.Y + MinYBorder Then
30                InRangoVision = True
40                Exit Function
50            End If
60        End If
70        InRangoVision = False

End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, X As Integer, Y As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If X > Npclist(NpcIndex).Pos.X - MinXBorder And X < Npclist(NpcIndex).Pos.X + MinXBorder Then
20            If Y > Npclist(NpcIndex).Pos.Y - MinYBorder And Y < Npclist(NpcIndex).Pos.Y + MinYBorder Then
30                InRangoVisionNPC = True
40                Exit Function
50            End If
60        End If
70        InRangoVisionNPC = False

End Function


Function InMapBounds(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If (map <= 0 Or map > NumMaps) Or X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
20            InMapBounds = False
30        Else
40            InMapBounds = True
50        End If
          
    End Function

Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos, Optional PuedeAgua As Boolean = False, Optional PuedeTierra As Boolean = True)
      '*****************************************************************
      'Author: Unknown (original version)
      'Last Modification: 24/01/2007 (ToxicWaste)
      'Encuentra la posicion legal mas cercana y la guarda en nPos
      '*****************************************************************

      Dim Notfound As Boolean
      Dim LoopC As Integer
      Dim tX As Long
      Dim tY As Long

10    nPos.map = Pos.map

20    Do While Not LegalPos(Pos.map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra)
30        If LoopC > 12 Then
40            Notfound = True
50            Exit Do
60        End If
          
70        For tY = Pos.Y - LoopC To Pos.Y + LoopC
80            For tX = Pos.X - LoopC To Pos.X + LoopC
                  
90                If LegalPos(nPos.map, tX, tY, PuedeAgua, PuedeTierra) Then
100                   nPos.X = tX
110                   nPos.Y = tY
                      '¿Hay objeto?
                      
120                   tX = Pos.X + LoopC
130                   tY = Pos.Y + LoopC
140               End If
150           Next tX
160       Next tY
          
170       LoopC = LoopC + 1
180   Loop

190   If Notfound = True Then
200       nPos.X = 0
210       nPos.Y = 0
220   End If

End Sub

Public Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
      '*****************************************************************

          Dim Notfound As Boolean
          Dim LoopC As Integer
          Dim tX As Long
          Dim tY As Long
          
10        nPos.map = Pos.map
          
20        Do While Not LegalPos(Pos.map, nPos.X, nPos.Y)
30            If LoopC > 12 Then
40                Notfound = True
50                Exit Do
60            End If
              
70            For tY = Pos.Y - LoopC To Pos.Y + LoopC
80                For tX = Pos.X - LoopC To Pos.X + LoopC
                      
90                    If LegalPos(nPos.map, tX, tY) And MapData(nPos.map, tX, tY).TileExit.map = 0 Then
100                       nPos.X = tX
110                       nPos.Y = tY
                          '¿Hay objeto?
                          
120                       tX = Pos.X + LoopC
130                       tY = Pos.Y + LoopC
140                   End If
150               Next tX
160           Next tY
              
170           LoopC = LoopC + 1
180       Loop
          
190       If Notfound = True Then
200           nPos.X = 0
210           nPos.Y = 0
220       End If

End Sub

Function NameIndex(ByVal Name As String) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim Userindex As Long
          
          '¿Nombre valido?
10        If LenB(Name) = 0 Then
20            NameIndex = 0
30            Exit Function
40        End If
          
50        If InStrB(Name, "+") <> 0 Then
60            Name = UCase$(Replace(Name, "+", " "))
70        End If
          
80        Userindex = 1
90        Do Until UCase$(UserList(Userindex).Name) = UCase$(Name)
              
100           Userindex = Userindex + 1
              
110           If Userindex > MaxUsers Then
120               NameIndex = 0
130               Exit Function
140           End If
150       Loop
           
160       NameIndex = Userindex
End Function
Function CheckForSameHD(ByVal Userindex As Integer, ByVal UserHD As String) As Boolean '//Disco.
       
          Dim LoopC As Long
         
   On Error GoTo CheckForSameHD_Error

10        For LoopC = 1 To MaxUsers
20            If UserList(LoopC).flags.UserLogged = True Then
30                If UserList(LoopC).HD = UserHD And Userindex <> LoopC Then
40                    CheckForSameHD = True
50                    Exit Function
60                End If
70            End If
80        Next LoopC
         
90        CheckForSameHD = False

   On Error GoTo 0
   Exit Function

CheckForSameHD_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckForSameHD of Módulo Extra in line " & Erl
End Function

Function CheckForSameIP(ByVal Userindex As Integer, ByVal UserIP As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim LoopC As Long
          Dim NumPjs As Byte
          
          
10        For LoopC = 1 To MaxUsers
20            If UserList(LoopC).flags.UserLogged = True Then
30                If UserList(LoopC).ip = UserIP And Userindex <> LoopC Then
40                    CheckForSameIP = True
50                    Exit Function
60                End If
70            End If
80        Next LoopC
          
90        CheckForSameIP = False
End Function

Function CheckForSameName(ByVal Name As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      'Controlo que no existan usuarios con el mismo nombre
          Dim LoopC As Long
          
10        For LoopC = 1 To LastUser
20            If UserList(LoopC).flags.UserLogged Then
                  
                  'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
                  'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
                  'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
                  'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
                  'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
                  
30                If UCase$(UserList(LoopC).Name) = UCase$(Name) Then
40                    CheckForSameName = True
50                    Exit Function
60                End If
70            End If
80        Next LoopC
          
90        CheckForSameName = False
End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      'Toma una posicion y se mueve hacia donde esta perfilado
      '*****************************************************************

10        Select Case Head
              Case eHeading.NORTH
20                Pos.Y = Pos.Y - 1
              
30            Case eHeading.SOUTH
40                Pos.Y = Pos.Y + 1
              
50            Case eHeading.EAST
60                Pos.X = Pos.X + 1
              
70            Case eHeading.WEST
80                Pos.X = Pos.X - 1
90        End Select
End Sub

Function LegalPos(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal CheckExitTile As Boolean = False) As Boolean
      '***************************************************
      'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
      'Last Modification: 23/01/2007
      'Checks if the position is Legal.
      '***************************************************

          '¿Es un mapa valido?
10        If (map <= 0 Or map > NumMaps) Or _
             (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
20                    LegalPos = False
30        Else
40            With MapData(map, X, Y)
50                If PuedeAgua And PuedeTierra Then
60                    LegalPos = (.Blocked <> 1) And _
                                 (.Userindex = 0) And _
                                 (.NpcIndex = 0)
70                ElseIf PuedeTierra And Not PuedeAgua Then
80                    LegalPos = (.Blocked <> 1) And _
                                 (.Userindex = 0) And _
                                 (.NpcIndex = 0) And _
                                 (Not HayAgua(map, X, Y))
90                ElseIf PuedeAgua And Not PuedeTierra Then
100                   LegalPos = (.Blocked <> 1) And _
                                 (.Userindex = 0) And _
                                 (.NpcIndex = 0) And _
                                 (HayAgua(map, X, Y))
110               Else
120                   LegalPos = False
130               End If
140           End With
              
150           If CheckExitTile Then
160               LegalPos = LegalPos And (MapData(map, X, Y).TileExit.map = 0)
170           End If
              
180       End If

End Function

Function MoveToLegalPos(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True) As Boolean
      '***************************************************
      'Autor: ZaMa
      'Last Modification: 13/07/2009
      'Checks if the position is Legal, but considers that if there's a casper, it's a legal movement.
      '13/07/2009: ZaMa - Now it's also legal move where an invisible admin is.
      '***************************************************

      Dim Userindex As Integer
      Dim IsDeadChar As Boolean
      Dim IsAdminInvisible As Boolean


      '¿Es un mapa valido?
10    If (map <= 0 Or map > NumMaps) Or _
         (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
20                MoveToLegalPos = False
30        Else
40            Userindex = MapData(map, X, Y).Userindex
50            If Userindex > 0 Then
60                IsDeadChar = UserList(Userindex).flags.Muerto = 1
70                IsAdminInvisible = UserList(Userindex).flags.AdminInvisible = 1
80            Else
90                IsDeadChar = False
100               IsAdminInvisible = False
110           End If
          
120       If PuedeAgua And PuedeTierra Then
130           MoveToLegalPos = (MapData(map, X, Y).Blocked <> 1) And _
                         (Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                         (MapData(map, X, Y).NpcIndex = 0)
140       ElseIf PuedeTierra And Not PuedeAgua Then
150           MoveToLegalPos = (MapData(map, X, Y).Blocked <> 1) And _
                         (Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                         (MapData(map, X, Y).NpcIndex = 0) And _
                         (Not HayAgua(map, X, Y))
160       ElseIf PuedeAgua And Not PuedeTierra Then
170           MoveToLegalPos = (MapData(map, X, Y).Blocked <> 1) And _
                         (Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                         (MapData(map, X, Y).NpcIndex = 0) And _
                         (HayAgua(map, X, Y))
180       Else
190           MoveToLegalPos = False
200       End If
        
210   End If

End Function
Public Sub FindLegalPos(ByVal Userindex As Integer, ByVal map As Integer, ByRef X As Integer, ByRef Y As Integer)
      '***************************************************
      'Autor: ZaMa
      'Last Modification: 26/03/2009
      'Search for a Legal pos for the user who is being teleported.
      '***************************************************


10        If MapData(map, X, Y).Userindex <> 0 Or _
              MapData(map, X, Y).NpcIndex <> 0 Then
                          
              ' Se teletransporta a la misma pos a la que estaba
20            If MapData(map, X, Y).Userindex = Userindex Then Exit Sub
                                  
              Dim FoundPlace As Boolean
              Dim tX As Long
              Dim tY As Long
              Dim Rango As Long
              Dim OtherUserIndex As Integer
          
30            For Rango = 1 To 5
40                For tY = Y - Rango To Y + Rango
50                    For tX = X - Rango To X + Rango
                          'Reviso que no haya User ni NPC
60                        If MapData(map, tX, tY).Userindex = 0 And _
                              MapData(map, tX, tY).NpcIndex = 0 Then
                              
70                            If InMapBounds(map, tX, tY) Then FoundPlace = True
                              
80                            Exit For
90                        End If

100                   Next tX
              
110                   If FoundPlace Then _
                          Exit For
120               Next tY
                  
130               If FoundPlace Then _
                          Exit For
140           Next Rango

          
150           If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
160               X = tX
170               Y = tY
180           Else
                  'Muy poco probable, pero..
                  'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
190               OtherUserIndex = MapData(map, X, Y).Userindex
200               If OtherUserIndex <> 0 Then
                      'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
210                   If UserList(OtherUserIndex).ComUsu.DestUsu > 0 Then
                          'Le avisamos al que estaba comerciando que se tuvo que ir.
220                       If UserList(UserList(OtherUserIndex).ComUsu.DestUsu).flags.UserLogged Then
230                           Call FinComerciarUsu(UserList(OtherUserIndex).ComUsu.DestUsu)
240                           Call WriteConsoleMsg(UserList(OtherUserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
250                           Call FlushBuffer(UserList(OtherUserIndex).ComUsu.DestUsu)
260                       End If
                          'Lo sacamos.
270                       If UserList(OtherUserIndex).flags.UserLogged Then
280                           Call FinComerciarUsu(OtherUserIndex)
290                           Call WriteErrorMsg(OtherUserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
300                           Call FlushBuffer(OtherUserIndex)
310                       End If
320                   End If
                  
330                   Call CloseSocket(OtherUserIndex)
340               End If
350           End If
360       End If

End Sub

Function LegalPosNPC(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte, Optional ByVal IsPet As Boolean = False) As Boolean
      '***************************************************
      'Autor: Unkwnown
      'Last Modification: 09/23/2009
      'Checks if it's a Legal pos for the npc to move to.
      '09/23/2009: Pato - If UserIndex is a AdminInvisible, then is a legal pos.
      '***************************************************
      Dim IsDeadChar As Boolean
      Dim Userindex As Integer
      Dim IsAdminInvisible As Boolean
          
          
10        If (map <= 0 Or map > NumMaps) Or _
              (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
20            LegalPosNPC = False
30            Exit Function
40        End If

50        With MapData(map, X, Y)
60            Userindex = .Userindex
70            If Userindex > 0 Then
80                IsDeadChar = UserList(Userindex).flags.Muerto = 1
90                IsAdminInvisible = (UserList(Userindex).flags.AdminInvisible = 1)
100           Else
110               IsDeadChar = False
120               IsAdminInvisible = False
130           End If
          
140           If AguaValida = 0 Then
150               LegalPosNPC = (.Blocked <> 1) And _
                  (.Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                  (.NpcIndex = 0) And _
                  (.trigger <> eTrigger.POSINVALIDA Or IsPet) _
                  And Not HayAgua(map, X, Y)
160           Else
170               LegalPosNPC = (.Blocked <> 1) And _
                  (.Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                  (.NpcIndex = 0) And _
                  (.trigger <> eTrigger.POSINVALIDA Or IsPet)
180           End If
190       End With
End Function


Sub SendHelp(ByVal Index As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim NumHelpLines As Integer
      Dim LoopC As Integer

10    NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

20    For LoopC = 1 To NumHelpLines
30        Call WriteConsoleMsg(Index, GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC), FontTypeNames.FONTTYPE_INFO)
40    Next LoopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal Userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If Npclist(NpcIndex).NroExpresiones > 0 Then
              Dim randomi
20            randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
30            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(Npclist(NpcIndex).Expresiones(randomi), Npclist(NpcIndex).Char.CharIndex, vbWhite))
40        End If
End Sub

Sub LookatTile(ByVal Userindex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 26/03/2009
      '13/02/2009: ZaMa - EL nombre del gm que aparece por consola al clickearlo, tiene el color correspondiente a su rango
      '***************************************************

10    On Error GoTo Errhandler

      'Responde al click del usuario sobre el mapa
      Dim FoundChar As Byte
      Dim FoundSomething As Byte
      Dim TempCharIndex As Integer
      Dim Stat As String
      Dim Ft As FontTypeNames
      Dim UserName As String

20    With UserList(Userindex)
          '¿Rango Visión? (ToxicWaste)
30        If (Abs(.Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(.Pos.X - X) > RANGO_VISION_X) Then
40            Exit Sub
50        End If
          
          '¿Posicion valida?
60        If InMapBounds(map, X, Y) Then
70            With .flags
80                .TargetMap = map
90                .TargetX = X
100               .TargetY = Y
                  '¿Es un obj?
110               If MapData(map, X, Y).ObjInfo.ObjIndex > 0 Then
                      'Informa el nombre
120                   .TargetObjMap = map
130                   .TargetObjX = X
140                   .TargetObjY = Y
150                   FoundSomething = 1
160               ElseIf MapData(map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
                      'Informa el nombre
170                   If ObjData(MapData(map, X + 1, Y).ObjInfo.ObjIndex).ObjType = eOBJType.otPuertas Then
180                       .TargetObjMap = map
190                       .TargetObjX = X + 1
200                       .TargetObjY = Y
210                       FoundSomething = 1
220                   End If
230               ElseIf MapData(map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
240                   If ObjData(MapData(map, X + 1, Y + 1).ObjInfo.ObjIndex).ObjType = eOBJType.otPuertas Then
                          'Informa el nombre
250                       .TargetObjMap = map
260                       .TargetObjX = X + 1
270                       .TargetObjY = Y + 1
280                       FoundSomething = 1
290                   End If
300               ElseIf MapData(map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
310                   If ObjData(MapData(map, X, Y + 1).ObjInfo.ObjIndex).ObjType = eOBJType.otPuertas Then
                          'Informa el nombre
320                       .TargetObjMap = map
330                       .TargetObjX = X
340                       .TargetObjY = Y + 1
350                       FoundSomething = 1
360                   End If
370               End If
                  
380               If FoundSomething = 1 Then
390                   .TargetObj = MapData(map, .TargetObjX, .TargetObjY).ObjInfo.ObjIndex
400                   If MostrarCantidad(.TargetObj) Then
410                       Call WriteConsoleMsg(Userindex, ObjData(.TargetObj).Name & " - " & MapData(.TargetObjMap, .TargetObjX, .TargetObjY).ObjInfo.Amount & "", FontTypeNames.FONTTYPE_INFO)
420                   Else
430                       Call WriteConsoleMsg(Userindex, ObjData(.TargetObj).Name, FontTypeNames.FONTTYPE_INFO)
440                   End If
                  
450               End If
                   '¿Es un personaje?
460               If Y + 1 <= YMaxMapSize Then
470                   If MapData(map, X, Y + 1).Userindex > 0 Then
480                       TempCharIndex = MapData(map, X, Y + 1).Userindex
490                       FoundChar = 1
500                   End If
510                   If MapData(map, X, Y + 1).NpcIndex > 0 Then
520                       TempCharIndex = MapData(map, X, Y + 1).NpcIndex
530                       FoundChar = 2
540                   End If
550               End If
                  '¿Es un personaje?
560               If FoundChar = 0 Then
570                   If MapData(map, X, Y).Userindex > 0 Then
580                       TempCharIndex = MapData(map, X, Y).Userindex
590                       FoundChar = 1
600                   End If
610                   If MapData(map, X, Y).NpcIndex > 0 Then
620                       TempCharIndex = MapData(map, X, Y).NpcIndex
630                       FoundChar = 2
640                   End If
650               End If
660           End With
          
          
              'Reaccion al personaje
670           If FoundChar = 1 Then '  ¿Encontro un Usuario?
680              If UserList(TempCharIndex).flags.AdminInvisible = 0 Or .flags.Privilegios And PlayerType.Dios Then
690                   With UserList(TempCharIndex)
700                       If LenB(.DescRM) = 0 And .showName Then 'No tiene descRM y quiere que se vea su nombre.
710
                              
790                           If .GuildIndex > 0 Then
800                               Stat = Stat & " <" & modGuilds.GuildName(.GuildIndex) & ">"
810                           End If
                              
820                           If Len(.desc) > 0 Then
830                               Stat = .Name & Stat & " - " & .desc
840                           Else
850                               Stat = .Name & Stat
860                           End If
                            Ft = FontTypeNames.FONTTYPE_NICK
                              Call WriteConsoleMsg(Userindex, Stat, Ft, True)
                                        
                            Stat = ""
                                        
                              If EsNewbie(TempCharIndex) Then
720                               Stat = " <NEWBIE>"
730                           End If
                              
                              If criminal(TempCharIndex) Then
1120                                 Stat = " <CRIMINAL>"
1140                              Else
1150                                  Stat = " <CIUDADANO>"
1170                             End If
                                Ft = FontTypeNames.FONTTYPE_EJECUCION
                              If Len(Stat) > 0 Then
                                Call WriteConsoleMsg(Userindex, Stat, Ft, False)
                              End If
                              Stat = ""
740                           If .Faccion.ArmadaReal = 1 Then
750                               Stat = " <Ejército Real> " & "<" & TituloReal(TempCharIndex) & ">"
760                               Ft = FontTypeNames.FONTTYPE_CITIZEN
                                ElseIf .Faccion.FuerzasCaos = 1 Then
770                               Stat = " <Legión Oscura> " & "<" & TituloCaos(TempCharIndex) & ">"
780                                Ft = FontTypeNames.FONTTYPE_FIGHT
                                
                                End If
                                
                                If Len(Stat) > 0 Then
                                Call WriteConsoleMsg(Userindex, Stat, Ft, False)
                              End If
                                Stat = ""
                              If EsGmChar(.Name) Then
                                    Stat = Stat & GetRangeData(.Name)
                              End If
                              Ft = FontTypeNames.FONTTYPE_GMMSG
                              
870                           If .flags.Privilegios And PlayerType.RoyalCouncil Then
880                               Stat = Stat & " [CONSEJO DE BANDERBILL]"
890                               Ft = FontTypeNames.FONTTYPE_CONSEJOVesA
900                           ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
910                               Stat = Stat & " [CONCILIO DE LAS SOMBRAS]"
920                               Ft = FontTypeNames.FONTTYPE_EJECUCION
930                           End If

                                If Len(Stat) > 0 Then
                                Call WriteConsoleMsg(Userindex, Stat, Ft, False)
                              End If
                                Stat = ""


1310                           If .flags.Premium > 0 Then
1320                          Stat = " [PREMIUM]"
1330                          Ft = FontTypeNames.FONTTYPE_PREMIUM
1340                          End If

                              If Len(Stat) > 0 Then
                                Call WriteConsoleMsg(Userindex, Stat, Ft, False)
                              End If
                              Stat = ""
1350                          If .flags.IsDios Then
1360                              Stat = Stat & " [DIOS]"
1370                          End If

                              
                              
                              'esto envia al Ser vip el nombre <VIP> al clickear
                              
1410                          If .flags.DiosTerrenal > 0 Then
1420                              Stat = Stat & " [Dios Terrenal]"
1430                              Ft = FontTypeNames.FONTTYPE_TALK
1440                          End If
                                If Len(Stat) > 0 Then
                                Call WriteConsoleMsg(Userindex, Stat, Ft, False)
                              End If
                              Stat = ""
1450                          If .flags.Oro > 0 Then
1460                              Stat = Stat & " [ORO]"
1470                              Ft = FontTypeNames.FONTTYPE_ORO
1480                          End If
                              If Len(Stat) > 0 Then
                                Call WriteConsoleMsg(Userindex, Stat, Ft, False)
                              End If
                              
                              Stat = ""
1490                          If .flags.Plata > 0 Then
1500                              Stat = Stat & " [PLATA]"
1510                              Ft = FontTypeNames.FONTTYPE_PLATA
1520                          End If
                              If Len(Stat) > 0 Then
                                Call WriteConsoleMsg(Userindex, Stat, Ft, False)
                              End If
                              Stat = ""
1530                          If .flags.Bronce > 0 Then
1540                              Stat = Stat & " [BRONCE]"
1550                              Ft = FontTypeNames.FONTTYPE_BRONCE
1560                          End If
                              If Len(Stat) > 0 Then
                                Call WriteConsoleMsg(Userindex, Stat, Ft, False)
                              End If
                              Stat = ""
1570                           If .flags.Infectado > 0 Then
1580                              Stat = Stat & " [INFECTADO]"
1590                              Ft = FontTypeNames.FONTTYPE_CONSE
1600                          End If
                              If Len(Stat) > 0 Then
                                Call WriteConsoleMsg(Userindex, Stat, Ft, False)
                              End If
                              Stat = ""
1610                          If .flags.Angel > 0 Then
1620                              Stat = Stat & " [ÁNGEL]"
1630                              Ft = FontTypeNames.FONTTYPE_CONSEJOVesA
1640                          End If
                              If Len(Stat) > 0 Then
                                Call WriteConsoleMsg(Userindex, Stat, Ft, False)
                              End If
                              Stat = ""
1650                                      If .flags.Demonio > 0 Then
1660                              Stat = Stat & " [DEMONIO]"
1670                              Ft = FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
1680                          End If
                             If Len(Stat) > 0 Then
                                Call WriteConsoleMsg(Userindex, Stat, Ft, False)
                              End If
                              Stat = ""
1690                          If .flags.Muerto = 1 Then
1700                                 Stat = Stat & " <MUERTO>"
1710                                 Ft = FontTypeNames.FONTTYPE_EJECUCION
1720                              End If
If Len(Stat) > 0 Then
                                Call WriteConsoleMsg(Userindex, Stat, Ft, False)
                              End If
Stat = ""
1730                      Else  'Si tiene descRM la muestro siempre.
1740                          Stat = .DescRM
1750                          Ft = FontTypeNames.FONTTYPE_INFOBOLD
1760                      End If
1770                  End With


                      
1780                  If LenB(Stat) > 0 Then
1790                      Call WriteConsoleMsg(Userindex, Stat, Ft)
1800                  End If
                      
1810                  FoundSomething = 1
1820                  .flags.TargetUser = TempCharIndex
1830                  .flags.TargetNPC = 0
1840                  .flags.TargetNpcTipo = eNPCType.Comun
1850             End If
1860          End If
          
1870          With .flags
1880              If FoundChar = 2 Then '¿Encontro un NPC?
                      Dim estatus As String
                      Dim MinHp As Long
                      Dim MaxHp As Long
                      Dim SupervivenciaSkill As Byte
                      Dim sDesc As String
                      
1890                  MinHp = Npclist(TempCharIndex).Stats.MinHp
1900                  MaxHp = Npclist(TempCharIndex).Stats.MaxHp
1910                  SupervivenciaSkill = UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia)
                      
1920                  If .Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then
1930                      estatus = "(" & MinHp & "/" & MaxHp & ") "
1940                  Else
1950                       If UserList(Userindex).flags.Muerto = 0 Then
1960                      If UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 10 Then
1970                          estatus = "(Dudoso) "
1980                      ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) > 10 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 20 Then
1990                          If Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp / 2) Then
2000                              estatus = "(Herido) "
2010                          Else
2020                              estatus = "(Sano) "
2030                          End If
2040                      ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) > 20 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 30 Then
2050                          If Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.5) Then
2060                              estatus = "(Malherido) "
2070                          ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.75) Then
2080                              estatus = "(Herido) "
2090                          Else
2100                              estatus = "(Sano) "
2110                          End If
2120                      ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) > 30 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 40 Then
2130                          If Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.25) Then
2140                              estatus = "(Muy malherido) "
2150                          ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.5) Then
2160                              estatus = "(Herido) "
2170                          ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.75) Then
2180                              estatus = "(Levemente herido) "
2190                          Else
2200                              estatus = "(Sano) "
2210                          End If
2220                      ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) > 40 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) < 60 Then
2230                          If Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.05) Then
2240                              estatus = "(Agonizando) "
2250                          ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.1) Then
2260                              estatus = "(Casi muerto) "
2270                          ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.25) Then
2280                              estatus = "(Muy Malherido) "
2290                          ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.5) Then
2300                              estatus = "(Herido) "
2310                          ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.75) Then
2320                              estatus = "(Levemente herido) "
2330                          ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp) Then
2340                              estatus = "(Sano) "
2350                          Else
2360                              estatus = "(Intacto) "
2370                          End If
2380                      ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 60 Then
2390                          estatus = "(" & Npclist(TempCharIndex).Stats.MinHp & "/" & Npclist(TempCharIndex).Stats.MaxHp & ") "
2400                      Else
2410                          estatus = "!error!"
2420                          End If
2430                      End If
2440                  End If
                      
2450                  If Len(Npclist(TempCharIndex).desc) > 1 Then
                          'Call WriteChatOverHead(Userindex, Npclist(TempCharIndex).desc, Npclist(TempCharIndex).Char.CharIndex, vbWhite)
2460                      WriteDescNpcs Userindex, Npclist(TempCharIndex).Char.CharIndex, Npclist(TempCharIndex).Numero
2470                  ElseIf TempCharIndex = CentinelaNPCIndex Then
                          'Enviamos nuevamente el texto del centinela según quien pregunta
2480                      Call modCentinela.CentinelaSendClave(Userindex)
2490                  Else
2500                      If Npclist(TempCharIndex).MaestroUser > 0 Then
2510                          Call WriteConsoleMsg(Userindex, estatus & Npclist(TempCharIndex).Name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).Name & ".", FontTypeNames.FONTTYPE_INFO)
2520                      Else
2530                          sDesc = estatus & Npclist(TempCharIndex).Name
2540                          If Npclist(TempCharIndex).Owner > 0 Then sDesc = sDesc & " le pertenece a " & UserList(Npclist(TempCharIndex).Owner).Name
2550                          sDesc = sDesc & "."
                              
2560                          Call WriteConsoleMsg(Userindex, sDesc, FontTypeNames.FONTTYPE_INFO)
                              
2570                          If .Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
2580                              Call WriteConsoleMsg(Userindex, "Le pegó primero: " & Npclist(TempCharIndex).flags.AttackedFirstBy & ".", FontTypeNames.FONTTYPE_INFO)
2590                          End If
2600                      End If
2610                  End If
                      
2620                  FoundSomething = 1
2630                  .TargetNpcTipo = Npclist(TempCharIndex).NPCtype
2640                  .TargetNPC = TempCharIndex
2650                  .TargetUser = 0
2660                  .TargetObj = 0
2670              End If
                  
2680              If FoundChar = 0 Then
2690                  .TargetNPC = 0
2700                  .TargetNpcTipo = eNPCType.Comun
2710                  .TargetUser = 0
2720              End If
                  
                  '*** NO ENCOTRO NADA ***
2730              If FoundSomething = 0 Then
2740                  .TargetNPC = 0
2750                  .TargetNpcTipo = eNPCType.Comun
2760                  .TargetUser = 0
2770                  .TargetObj = 0
2780                  .TargetObjMap = 0
2790                  .TargetObjX = 0
2800                  .TargetObjY = 0
2810                  Call WriteMultiMessage(Userindex, eMessages.DontSeeAnything)
2820              End If
2830          End With
2840      Else
2850          If FoundSomething = 0 Then
2860              With .flags
2870                  .TargetNPC = 0
2880                  .TargetNpcTipo = eNPCType.Comun
2890                  .TargetUser = 0
2900                  .TargetObj = 0
2910                  .TargetObjMap = 0
2920                  .TargetObjX = 0
2930                  .TargetObjY = 0
2940              End With
                  
2950              Call WriteMultiMessage(Userindex, eMessages.DontSeeAnything)
2960          End If
2970      End If
2980  End With

2990  Exit Sub

Errhandler:
3000      Call LogError("Error en LookAtTile. Error " & Err.Number & " : " & Err.Description)

End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As eHeading
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      'Devuelve la direccion en la cual el target se encuentra
      'desde pos, 0 si la direc es igual
      '*****************************************************************

          Dim X As Integer
          Dim Y As Integer
          
10        X = Pos.X - Target.X
20        Y = Pos.Y - Target.Y
          
          'NE
30        If Sgn(X) = -1 And Sgn(Y) = 1 Then
40            FindDirection = IIf(RandomNumber(0, 1), eHeading.NORTH, eHeading.EAST)
50            Exit Function
60        End If
          
          'NW
70        If Sgn(X) = 1 And Sgn(Y) = 1 Then
80            FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.NORTH)
90            Exit Function
100       End If
          
          'SW
110       If Sgn(X) = 1 And Sgn(Y) = -1 Then
120           FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.SOUTH)
130           Exit Function
140       End If
          
          'SE
150       If Sgn(X) = -1 And Sgn(Y) = -1 Then
160           FindDirection = IIf(RandomNumber(0, 1), eHeading.SOUTH, eHeading.EAST)
170           Exit Function
180       End If
          
          'Sur
190       If Sgn(X) = 0 And Sgn(Y) = -1 Then
200           FindDirection = eHeading.SOUTH
210           Exit Function
220       End If
          
          'norte
230       If Sgn(X) = 0 And Sgn(Y) = 1 Then
240           FindDirection = eHeading.NORTH
250           Exit Function
260       End If
          
          'oeste
270       If Sgn(X) = 1 And Sgn(Y) = 0 Then
280           FindDirection = eHeading.WEST
290           Exit Function
300       End If
          
          'este
310       If Sgn(X) = -1 And Sgn(Y) = 0 Then
320           FindDirection = eHeading.EAST
330           Exit Function
340       End If
          
          'misma
350       If Sgn(X) = 0 And Sgn(Y) = 0 Then
360           FindDirection = 0
370           Exit Function
380       End If

End Function

Public Function ItemNoEsDeMapa(ByVal Index As Integer, ByVal bIsExit As Boolean) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With ObjData(Index)
20            ItemNoEsDeMapa = .ObjType <> eOBJType.otPuertas And _
                          .ObjType <> eOBJType.otForos And _
                          .ObjType <> eOBJType.otCarteles And _
                          .ObjType <> eOBJType.otarboles And _
                          .ObjType <> eOBJType.otYacimiento And _
                          .ObjType <> eOBJType.otTeleport And _
                          Not (.ObjType = eOBJType.otTeleport And bIsExit)
          
30        End With

End Function

Public Function MostrarCantidad(ByVal Index As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With ObjData(Index)
20            MostrarCantidad = .ObjType <> eOBJType.otPuertas And _
                          .ObjType <> eOBJType.otForos And _
                          .ObjType <> eOBJType.otCarteles And _
                          .ObjType <> eOBJType.otarboles And _
                          .ObjType <> eOBJType.otYacimiento And _
                          .ObjType <> eOBJType.otTeleport
30        End With

End Function

Public Function EsObjetoFijo(ByVal ObjType As eOBJType) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        EsObjetoFijo = ObjType = eOBJType.otForos Or _
                         ObjType = eOBJType.otCarteles Or _
                         ObjType = eOBJType.otarboles Or _
                         ObjType = eOBJType.otYacimiento
End Function
Private Function RhombLegalPos(ByRef Pos As WorldPos, ByRef vX As Long, ByRef vY As Long, _
                               ByVal Distance As Long, Optional PuedeAgua As Boolean = False, _
                               Optional PuedeTierra As Boolean = True, _
                               Optional ByVal CheckExitTile As Boolean = False) As Boolean
      '***************************************************
      'Author: Marco Vanotti (Marco)
      'Last Modification: -
      ' walks all the perimeter of a rhomb of side  "distance + 1",
      ' which starts at Pos.x - Distance and Pos.y
      '***************************************************

          Dim i As Long
          
10        vX = Pos.X - Distance
20        vY = Pos.Y
          
30        For i = 0 To Distance - 1
40            If (LegalPos(Pos.map, vX + i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
50                vX = vX + i
60                vY = vY - i
70                RhombLegalPos = True
80                Exit Function
90            End If
100       Next
          
110       vX = Pos.X
120       vY = Pos.Y - Distance
          
130       For i = 0 To Distance - 1
140           If (LegalPos(Pos.map, vX + i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
150               vX = vX + i
160               vY = vY + i
170               RhombLegalPos = True
180               Exit Function
190           End If
200       Next
          
210       vX = Pos.X + Distance
220       vY = Pos.Y
          
230       For i = 0 To Distance - 1
240           If (LegalPos(Pos.map, vX - i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
250               vX = vX - i
260               vY = vY + i
270               RhombLegalPos = True
280               Exit Function
290           End If
300       Next
          
310       vX = Pos.X
320       vY = Pos.Y + Distance
          
330       For i = 0 To Distance - 1
340           If (LegalPos(Pos.map, vX - i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
350               vX = vX - i
360               vY = vY - i
370               RhombLegalPos = True
380               Exit Function
390           End If
400       Next
          
410       RhombLegalPos = False
          
End Function

Public Function RhombLegalTilePos(ByRef Pos As WorldPos, ByRef vX As Long, ByRef vY As Long, _
                                  ByVal Distance As Long, ByVal ObjIndex As Integer, ByVal ObjAmount As Long, _
                                  ByVal PuedeAgua As Boolean, ByVal PuedeTierra As Boolean) As Boolean
      '***************************************************
      'Author: ZaMa
      'Last Modification: -
      ' walks all the perimeter of a rhomb of side  "distance + 1",
      ' which starts at Pos.x - Distance and Pos.y
      ' and searchs for a valid position to drop items
      '***************************************************
10    On Error GoTo Errhandler

          Dim i As Long
          Dim HayObj As Boolean
          
          Dim X As Integer
          Dim Y As Integer
          Dim MapObjIndex As Integer
          
20        vX = Pos.X - Distance
30        vY = Pos.Y
          
40        For i = 0 To Distance - 1
              
50            X = vX + i
60            Y = vY - i
              
70            If (LegalPos(Pos.map, X, Y, PuedeAgua, PuedeTierra, True)) Then
                  
                  ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
80                If Not HayObjeto(Pos.map, X, Y, ObjIndex, ObjAmount) Then
90                    vX = X
100                   vY = Y
                      
110                   RhombLegalTilePos = True
120                   Exit Function
130               End If
                  
140           End If
150       Next
          
160       vX = Pos.X
170       vY = Pos.Y - Distance
          
180       For i = 0 To Distance - 1
              
190           X = vX + i
200           Y = vY + i
              
210           If (LegalPos(Pos.map, X, Y, PuedeAgua, PuedeTierra, True)) Then
                  
                  ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
220               If Not HayObjeto(Pos.map, X, Y, ObjIndex, ObjAmount) Then
230                   vX = X
240                   vY = Y
                      
250                   RhombLegalTilePos = True
260                   Exit Function
270               End If
280           End If
290       Next
          
300       vX = Pos.X + Distance
310       vY = Pos.Y
          
320       For i = 0 To Distance - 1
              
330           X = vX - i
340           Y = vY + i
          
350           If (LegalPos(Pos.map, X, Y, PuedeAgua, PuedeTierra, True)) Then
              
                  ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
360               If Not HayObjeto(Pos.map, X, Y, ObjIndex, ObjAmount) Then
370                   vX = X
380                   vY = Y
                      
390                   RhombLegalTilePos = True
400                   Exit Function
410               End If
420           End If
430       Next
          
440       vX = Pos.X
450       vY = Pos.Y + Distance
          
460       For i = 0 To Distance - 1
              
470           X = vX - i
480           Y = vY - i
          
490           If (LegalPos(Pos.map, X, Y, PuedeAgua, PuedeTierra, True)) Then
                  ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
500               If Not HayObjeto(Pos.map, X, Y, ObjIndex, ObjAmount) Then
510                   vX = X
520                   vY = Y
                      
530                   RhombLegalTilePos = True
540                   Exit Function
550               End If
560           End If
570       Next
          
580       RhombLegalTilePos = False
          
590       Exit Function
          
Errhandler:
600       Call LogError("Error en RhombLegalTilePos. Error: " & Err.Number & " - " & Err.Description)
End Function

Public Function HayObjeto(ByVal Mapa As Integer, ByVal X As Long, ByVal Y As Long, _
                          ByVal ObjIndex As Integer, ByVal ObjAmount As Long) As Boolean
      '***************************************************
      'Author: ZaMa
      'Last Modification: -
      'Checks if there's space in a tile to add an itemAmount
      '***************************************************
          Dim MapObjIndex As Integer
10        MapObjIndex = MapData(Mapa, X, Y).ObjInfo.ObjIndex
                  
          ' Hay un objeto tirado?
20        If MapObjIndex <> 0 Then
              ' Es el mismo objeto?
30            If MapObjIndex = ObjIndex Then
                  ' La suma es menor a 10k?
40                HayObjeto = (MapData(Mapa, X, Y).ObjInfo.Amount + ObjAmount > MAX_INVENTORY_OBJS)
50            Else
60                HayObjeto = True
70            End If
80        Else
90            HayObjeto = False
100       End If

End Function
Public Function EsPremium(ByVal Userindex As Integer) As Boolean 'es un usuario premium?
10    EsPremium = UserList(Userindex).flags.Premium
End Function

Public Sub ShowMenu(ByVal Userindex As Integer, ByVal map As Integer, _
    ByVal X As Integer, ByVal Y As Integer)
      '***************************************************
      'Autor: ZaMa
      'Last Modification: 10/05/2010
      'Shows menu according to user, npc or object right clicked.
      '***************************************************

10    On Error GoTo Errhandler

20        With UserList(Userindex)
              
              ' In Vision Range
30            If (Abs(.Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(.Pos.X - X) > RANGO_VISION_X) Then Exit Sub
              
              ' Valid position?
40            If Not InMapBounds(map, X, Y) Then Exit Sub
              
50            With .flags
                  ' Alive?
60                If .Muerto = 1 Then Exit Sub
                  
                  ' Trading?
70                If .Comerciando Then Exit Sub
                  
                  ' Reset flags
80                .TargetNPC = 0
90                .TargetNpcTipo = eNPCType.Comun
100               .TargetUser = 0
110               .TargetObj = 0
120               .TargetObjMap = 0
130               .TargetObjX = 0
140               .TargetObjY = 0
                  
150               .TargetMap = map
160               .TargetX = X
170               .TargetY = Y
                  
                  Dim tmpIndex As Integer
                  Dim FoundChar As Byte
                  Dim MenuIndex As Integer
                  
                  ' Npc or user? (lower position)
180               If Y + 1 <= YMaxMapSize Then
                      
                      ' User?
190                   tmpIndex = MapData(map, X, Y + 1).Userindex
200                   If tmpIndex > 0 Then
                          ' Invalid: Admin invisible, user invisible or hidden (if not is userindex)
210                       If (UserList(tmpIndex).flags.AdminInvisible Or _
                              UserList(tmpIndex).flags.invisible Or _
                              UserList(tmpIndex).flags.Oculto) = 0 Or tmpIndex = Userindex Then
                              
220                           FoundChar = 1
230                       End If
240                   End If
                      
                      ' Npc?
250                   If MapData(map, X, Y + 1).NpcIndex > 0 Then
260                       tmpIndex = MapData(map, X, Y + 1).NpcIndex
270                       FoundChar = 2
280                   End If
290               End If
                   
                  ' Npc or user? (upper position)
300               If FoundChar = 0 Then
                      
                      ' User?
310                   tmpIndex = MapData(map, X, Y).Userindex
320                   If tmpIndex > 0 Then
                          ' Invalid: Admin invisible, user invisible or hidden (if not is userindex)
330                       If (UserList(tmpIndex).flags.AdminInvisible Or _
                              UserList(tmpIndex).flags.invisible Or _
                              UserList(tmpIndex).flags.Oculto) = 0 Or tmpIndex = Userindex Then
                              
340                           FoundChar = 1
350                       End If
360                   End If
                      
                      ' Npc?
370                   If MapData(map, X, Y).NpcIndex > 0 Then
380                       tmpIndex = MapData(map, X, Y).NpcIndex
390                       FoundChar = 2
400                   End If
410               End If
                  
                  ' User
420               If FoundChar = 1 Then
430                   MenuIndex = eMenues.ieUser
                          
440                   .TargetUser = tmpIndex
                      
                  ' Npc
450               ElseIf FoundChar = 2 Then
                      ' Has menu attached?
                      'If Npclist(tmpIndex).MenuIndex <> 0 Then
                          'MenuIndex = Npclist(tmpIndex).MenuIndex
                     ' End If
                      
                      '.TargetNpcTipo = Npclist(tmpIndex).NPCtype
                      '.TargetNPC = tmpIndex
460               End If
                  
                  ' No user or npc found
470               If FoundChar = 0 Then
                      
                      ' Is there any object?
480                   tmpIndex = MapData(map, X, Y).ObjInfo.ObjIndex
490                   If tmpIndex > 0 Then
                          ' Has menu attached?
                          'MenuIndex = ObjData(tmpIndex).MenuIndex
                          
                          'If MenuIndex = eMenues.ieFogata Then
                              'If .Descansar = 1 Then MenuIndex = eMenues.ieFogataDescansando
                          'End If
                          
                          '.TargetObjMap = Map
                          '.TargetObjX = X
                         ' .TargetObjY = Y
500                   End If
510               End If
520           End With
530       End With
          
          ' Show it
540       If MenuIndex <> 0 Then _
              Call WriteShowMenu(Userindex, MenuIndex)
          
550       Exit Sub

Errhandler:
560       Call LogError("Error en ShowMenu. Error " & Err.Number & " : " & Err.Description)
End Sub

