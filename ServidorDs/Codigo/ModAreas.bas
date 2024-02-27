Attribute VB_Name = "ModAreas"
'**************************************************************
' ModAreas.bas - Module to allow the usage of areas instead of maps.
' Saves a lot of bandwidth.
'
' Original Idea by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
' Implemented by Lucio N. Tourrilhes (DuNga)
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

' Modulo de envio por areas compatible con la versión 9.10.x ... By DuNga

Option Explicit

'>>>>>>AREAS>>>>>AREAS>>>>>>>>AREAS>>>>>>>AREAS>>>>>>>>>>
Public Type AreaInfo
    AreaPerteneceX As Integer
    AreaPerteneceY As Integer
    
    AreaReciveX As Integer
    AreaReciveY As Integer
    
    MinX As Integer '-!!!
    MinY As Integer '-!!!
    
    AreaID As Long
End Type

Public Type ConnGroup
    CountEntrys As Long
    OptValue As Long
    UserEntrys() As Long
End Type

Public Const USER_NUEVO As Byte = 255

'Cuidado:
' ¡¡¡LAS AREAS ESTÁN HARDCODEADAS!!!
Private CurDay As Byte
Private CurHour As Byte

Private AreasInfo(1 To 100, 1 To 100) As Byte
Private PosToArea(1 To 100) As Byte

Private AreasRecive(12) As Integer

Public ConnGroups() As ConnGroup

Public Sub InitAreas()
      '**************************************************************
      'Author: Lucio N. Tourrilhes (DuNga)
      'Last Modify Date: Unknow
      '
      '**************************************************************
          Dim LoopC As Long
          Dim LoopX As Long

      ' Setup areas...
10        For LoopC = 0 To 11
20            AreasRecive(LoopC) = (2 ^ LoopC) Or IIf(LoopC <> 0, 2 ^ (LoopC - 1), 0) Or IIf(LoopC <> 11, 2 ^ (LoopC + 1), 0)
30        Next LoopC
          
40        For LoopC = 1 To 100
50            PosToArea(LoopC) = LoopC \ 9
60        Next LoopC
          
70        For LoopC = 1 To 100
80            For LoopX = 1 To 100
                  'Usamos 121 IDs de area para saber si pasasamos de area "más rápido"
90                AreasInfo(LoopC, LoopX) = (LoopC \ 9 + 1) * (LoopX \ 9 + 1)
100           Next LoopX
110       Next LoopC

      'Setup AutoOptimizacion de areas
120       CurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
130       CurHour = Fix(Hour(time) \ 3) 'A ke parte de la hora pertenece
          
140       ReDim ConnGroups(1 To NumMaps) As ConnGroup
          
150       For LoopC = 1 To NumMaps
160           ConnGroups(LoopC).OptValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, CurDay & "-" & CurHour))
              
170           If ConnGroups(LoopC).OptValue = 0 Then ConnGroups(LoopC).OptValue = 1
180           ReDim ConnGroups(LoopC).UserEntrys(1 To ConnGroups(LoopC).OptValue) As Long
190       Next LoopC
End Sub

Public Sub AreasOptimizacion()
      '**************************************************************
      'Author: Lucio N. Tourrilhes (DuNga)
      'Last Modify Date: Unknow
      'Es la función de autooptimizacion.... la idea es no mandar redimensionando arrays grandes todo el tiempo
      '**************************************************************
          Dim LoopC As Long
          Dim tCurDay As Byte
          Dim tCurHour As Byte
          Dim EntryValue As Long
          
10        If (CurDay <> IIf(Weekday(Date) > 6, 1, 2)) Or (CurHour <> Fix(Hour(time) \ 3)) Then
              
20            tCurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
30            tCurHour = Fix(Hour(time) \ 3) 'A ke parte de la hora pertenece
              
40            For LoopC = 1 To NumMaps
50                EntryValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, CurDay & "-" & CurHour))
60                Call WriteVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, CurDay & "-" & CurHour, CInt((EntryValue + ConnGroups(LoopC).OptValue) \ 2))
                  
70                ConnGroups(LoopC).OptValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, tCurDay & "-" & tCurHour))
80                If ConnGroups(LoopC).OptValue = 0 Then ConnGroups(LoopC).OptValue = 1
90                If ConnGroups(LoopC).OptValue >= MapInfo(LoopC).NumUsers Then ReDim Preserve ConnGroups(LoopC).UserEntrys(1 To ConnGroups(LoopC).OptValue) As Long
100           Next LoopC
              
110           CurDay = tCurDay
120           CurHour = tCurHour
130       End If
End Sub

Public Sub CheckUpdateNeededUser(ByVal Userindex As Integer, ByVal Head As Byte, Optional ByVal ButIndex As Boolean = False)
      '**************************************************************
      'Author: Lucio N. Tourrilhes (DuNga)
      'Last Modify Date: 28/10/2010
      'Es la función clave del sistema de areas... Es llamada al mover un user
      '15/07/2009: ZaMa - Now it doesn't send an invisible admin char info
      '28/10/2010: ZaMa - Now it doesn't send a saling char invisible message.
      '**************************************************************
10        If UserList(Userindex).AreasInfo.AreaID = AreasInfo(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y) Then Exit Sub
          
          Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long
          Dim TempInt As Long, map As Long
          Dim isZonaOscura As Boolean
          
20        With UserList(Userindex)
30            MinX = .AreasInfo.MinX
40            MinY = .AreasInfo.MinY
              
50            If Head = eHeading.NORTH Then
60                MaxY = MinY - 1
70                MinY = MinY - 9
80                MaxX = MinX + 26
90                .AreasInfo.MinX = CInt(MinX)
100               .AreasInfo.MinY = CInt(MinY)
              
110           ElseIf Head = eHeading.SOUTH Then
120               MaxY = MinY + 35
130               MinY = MinY + 27
140               MaxX = MinX + 26
150               .AreasInfo.MinX = CInt(MinX)
160               .AreasInfo.MinY = CInt(MinY - 18)
              
170           ElseIf Head = eHeading.WEST Then
180               MaxX = MinX - 1
190               MinX = MinX - 9
200               MaxY = MinY + 26
210               .AreasInfo.MinX = CInt(MinX)
220               .AreasInfo.MinY = CInt(MinY)
              
              
230           ElseIf Head = eHeading.EAST Then
240               MaxX = MinX + 35
250               MinX = MinX + 27
260               MaxY = MinY + 26
270               .AreasInfo.MinX = CInt(MinX - 18)
280               .AreasInfo.MinY = CInt(MinY)
              
                 
290           ElseIf Head = USER_NUEVO Then
                  'Esto pasa por cuando cambiamos de mapa o logeamos...
300               MinY = ((.Pos.Y \ 9) - 1) * 9
310               MaxY = MinY + 26
                  
320               MinX = ((.Pos.X \ 9) - 1) * 9
330               MaxX = MinX + 26
                  
340               .AreasInfo.MinX = CInt(MinX)
350               .AreasInfo.MinY = CInt(MinY)
360           End If
              
370           If MinY < 1 Then MinY = 1
380           If MinX < 1 Then MinX = 1
390           If MaxY > 100 Then MaxY = 100
400           If MaxX > 100 Then MaxX = 100
              
410           map = .Pos.map
              
              'Esto es para ke el cliente elimine lo "fuera de area..."
420           Call WriteAreaChanged(Userindex)
              
430           isZonaOscura = (MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.zonaOscura)
              
              'Actualizamos!!!
440           For X = MinX To MaxX
450               For Y = MinY To MaxY
                      
                      '<<< User >>>
460                   If MapData(map, X, Y).Userindex Then
                          
470                       TempInt = MapData(map, X, Y).Userindex
                          
480                       If Userindex <> TempInt Then
                              
                              ' Solo avisa al otro cliente si no es un admin invisible
490                           If Not (UserList(TempInt).flags.AdminInvisible = 1) Then
500                               Call MakeUserChar(False, Userindex, TempInt, map, X, Y)
                                  
                                  ' Si esta navegando, siempre esta visible
510                               If UserList(TempInt).flags.Navegando = 0 Then
520                                   If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
530                                       If MapData(map, X, Y).trigger = eTrigger.zonaOscura Then
540                                           Call WriteSetInvisible(Userindex, UserList(TempInt).Char.CharIndex, True)
550                                       Else
                                              'Si el user estaba invisible le avisamos al nuevo cliente de eso
560                                           If UserList(TempInt).flags.invisible Or UserList(TempInt).flags.Oculto Then
570                                               Call WriteSetInvisible(Userindex, UserList(TempInt).Char.CharIndex, True)
580                                           End If
590                                       End If
600                                   End If
610                               End If
620                           End If
                              
                              
                              ' Solo avisa al otro cliente si no es un admin invisible
630                           If Not (.flags.AdminInvisible = 1) Then
640                               Call MakeUserChar(False, TempInt, Userindex, .Pos.map, .Pos.X, .Pos.Y)
                                  
                                  ' Si esta navegando, siempre esta visible
650                               If .flags.Navegando = 0 Then
660                                   If UserList(TempInt).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
670                                       If isZonaOscura Then
680                                           Call WriteSetInvisible(TempInt, .Char.CharIndex, True)
690                                       Else
700                                           If .flags.invisible Or .flags.Oculto Then
710                                               Call WriteSetInvisible(TempInt, .Char.CharIndex, True)
720                                           End If
730                                       End If
740                                   End If
750                               End If
760                           End If
                              
                              
770                           Call FlushBuffer(TempInt)
                          
780                       ElseIf Head = USER_NUEVO Then
790                           If Not ButIndex Then
800                               Call MakeUserChar(False, Userindex, Userindex, map, X, Y)
810                           End If
820                       End If
830                   End If
                      
                      '<<< Npc >>>
840                   If MapData(map, X, Y).NpcIndex Then
850                       Call MakeNPCChar(False, Userindex, MapData(map, X, Y).NpcIndex, map, X, Y)
860                   End If
                       
                      '<<< Item >>>
870                   If MapData(map, X, Y).ObjInfo.objindex Then
880                       If MapData(map, X, Y).trigger <> eTrigger.zonaOscura Then
890                           TempInt = MapData(map, X, Y).ObjInfo.objindex
900                           If Not EsObjetoFijo(ObjData(TempInt).OBJType) Then
910                               Call WriteObjectCreate(Userindex, ObjData(TempInt).GrhIndex, X, Y)
                                  
920                               If ObjData(TempInt).OBJType = eOBJType.otPuertas Then
930                                   Call Bloquear(False, Userindex, X, Y, MapData(map, X, Y).Blocked)
940                                   Call Bloquear(False, Userindex, X - 1, Y, MapData(map, X - 1, Y).Blocked)
950                               End If
960                           End If
970                       End If
980                   End If
                  
990               Next Y
1000          Next X
              
              'Precalculados :P
1010          TempInt = .Pos.X \ 9
1020          .AreasInfo.AreaReciveX = AreasRecive(TempInt)
1030          .AreasInfo.AreaPerteneceX = 2 ^ TempInt
              
1040          TempInt = .Pos.Y \ 9
1050          .AreasInfo.AreaReciveY = AreasRecive(TempInt)
1060          .AreasInfo.AreaPerteneceY = 2 ^ TempInt
              
1070          .AreasInfo.AreaID = AreasInfo(.Pos.X, .Pos.Y)
1080      End With
End Sub

Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal Head As Byte)
      '**************************************************************
      'Author: Lucio N. Tourrilhes (DuNga)
      'Last Modify Date: Unknow
      ' Se llama cuando se mueve un Npc
      '**************************************************************
10        If Npclist(NpcIndex).AreasInfo.AreaID = AreasInfo(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then Exit Sub
          
          Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long
          Dim TempInt As Long
          Dim Userindex As Integer
          Dim isZonaOscura As Boolean
          
20        With Npclist(NpcIndex)
30            MinX = .AreasInfo.MinX
40            MinY = .AreasInfo.MinY
              
50            If Head = eHeading.NORTH Then
60                MaxY = MinY - 1
70                MinY = MinY - 9
80                MaxX = MinX + 26
90                .AreasInfo.MinX = CInt(MinX)
100               .AreasInfo.MinY = CInt(MinY)
              
110           ElseIf Head = eHeading.SOUTH Then
120               MaxY = MinY + 35
130               MinY = MinY + 27
140               MaxX = MinX + 26
150               .AreasInfo.MinX = CInt(MinX)
160               .AreasInfo.MinY = CInt(MinY - 18)
              
170           ElseIf Head = eHeading.WEST Then
180               MaxX = MinX - 1
190               MinX = MinX - 9
200               MaxY = MinY + 26
210               .AreasInfo.MinX = CInt(MinX)
220               .AreasInfo.MinY = CInt(MinY)
              
              
230           ElseIf Head = eHeading.EAST Then
240               MaxX = MinX + 35
250               MinX = MinX + 27
260               MaxY = MinY + 26
270               .AreasInfo.MinX = CInt(MinX - 18)
280               .AreasInfo.MinY = CInt(MinY)
              
                 
290           ElseIf Head = USER_NUEVO Then
                  'Esto pasa por cuando cambiamos de mapa o logeamos...
300               MinY = ((.Pos.Y \ 9) - 1) * 9
310               MaxY = MinY + 26
                  
320               MinX = ((.Pos.X \ 9) - 1) * 9
330               MaxX = MinX + 26
                  
340               .AreasInfo.MinX = CInt(MinX)
350               .AreasInfo.MinY = CInt(MinY)
360           End If
              
370           If MinY < 1 Then MinY = 1
380           If MinX < 1 Then MinX = 1
390           If MaxY > 100 Then MaxY = 100
400           If MaxX > 100 Then MaxX = 100

              
              'Actualizamos!!!
410           If MapInfo(.Pos.map).NumUsers <> 0 Then
420               isZonaOscura = (MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.zonaOscura)
                  
430               For X = MinX To MaxX
440                   For Y = MinY To MaxY
450                       Userindex = MapData(.Pos.map, X, Y).Userindex
                          
460                       If Userindex Then
470                           Call MakeNPCChar(False, Userindex, NpcIndex, .Pos.map, .Pos.X, .Pos.Y)
                              
480                           If isZonaOscura Then
490                               Call WriteSetInvisible(Userindex, .Char.CharIndex, True)
500                           End If
510                       End If
520                   Next Y
530               Next X
540           End If
              
              'Precalculados :P
550           TempInt = .Pos.X \ 9
560           .AreasInfo.AreaReciveX = AreasRecive(TempInt)
570           .AreasInfo.AreaPerteneceX = 2 ^ TempInt
                  
580           TempInt = .Pos.Y \ 9
590           .AreasInfo.AreaReciveY = AreasRecive(TempInt)
600           .AreasInfo.AreaPerteneceY = 2 ^ TempInt
              
610           .AreasInfo.AreaID = AreasInfo(.Pos.X, .Pos.Y)
620       End With
End Sub

Public Sub QuitarUser(ByVal Userindex As Integer, ByVal map As Integer)
      '**************************************************************
      'Author: Lucio N. Tourrilhes (DuNga)
      'Last Modify Date: Unknow
      '
      '**************************************************************
10    On Error GoTo ErrorHandler

          Dim TempVal As Long
          Dim LoopC As Long
          
          'Search for the user
20        For LoopC = 1 To ConnGroups(map).CountEntrys
30            If ConnGroups(map).UserEntrys(LoopC) = Userindex Then Exit For
40        Next LoopC
          
          'Char not found
50        If LoopC > ConnGroups(map).CountEntrys Then Exit Sub
          
          'Remove from old map
60        ConnGroups(map).CountEntrys = ConnGroups(map).CountEntrys - 1
70        TempVal = ConnGroups(map).CountEntrys
          
          'Move list back
80        For LoopC = LoopC To TempVal
90            ConnGroups(map).UserEntrys(LoopC) = ConnGroups(map).UserEntrys(LoopC + 1)
100       Next LoopC
          
110       If TempVal > ConnGroups(map).OptValue Then 'Nescesito Redim?
120           ReDim Preserve ConnGroups(map).UserEntrys(1 To TempVal) As Long
130       End If
          
140       Exit Sub
          
ErrorHandler:
          
          Dim UserName As String
150       If Userindex > 0 Then UserName = UserList(Userindex).Name

160       Call LogError("Error en QuitarUser " & Err.Number & ": " & Err.Description & _
                        ". User: " & UserName & "(" & Userindex & ")")

End Sub

Public Sub AgregarUser(ByVal Userindex As Integer, ByVal map As Integer, Optional ByVal ButIndex As Boolean = False)
      '**************************************************************
      'Author: Lucio N. Tourrilhes (DuNga)
      'Last Modify Date: 04/01/2007
      'Modified by Juan Martín Sotuyo Dodero (Maraxus)
      '   - Now the method checks for repetead users instead of trusting parameters.
      '   - If the character is new to the map, update it
      '**************************************************************
          Dim TempVal As Long
          Dim EsNuevo As Boolean
          Dim i As Long
          
10        If Not MapaValido(map) Then Exit Sub
          
20        EsNuevo = True
          
          'Prevent adding repeated users
30        For i = 1 To ConnGroups(map).CountEntrys
40            If ConnGroups(map).UserEntrys(i) = Userindex Then
50                EsNuevo = False
60                Exit For
70            End If
80        Next i
          
90        If EsNuevo Then
              'Update map and connection groups data
100           ConnGroups(map).CountEntrys = ConnGroups(map).CountEntrys + 1
110           TempVal = ConnGroups(map).CountEntrys
              
120           If TempVal > ConnGroups(map).OptValue Then 'Nescesito Redim
130               ReDim Preserve ConnGroups(map).UserEntrys(1 To TempVal) As Long
140           End If
              
150           ConnGroups(map).UserEntrys(TempVal) = Userindex
160       End If
          
170       With UserList(Userindex)
              'Update user
180           .AreasInfo.AreaID = 0
              
190           .AreasInfo.AreaPerteneceX = 0
200           .AreasInfo.AreaPerteneceY = 0
210           .AreasInfo.AreaReciveX = 0
220           .AreasInfo.AreaReciveY = 0
230       End With
          
240       Call CheckUpdateNeededUser(Userindex, USER_NUEVO, ButIndex)
End Sub

Public Sub AgregarNpc(ByVal NpcIndex As Integer)
      '**************************************************************
      'Author: Lucio N. Tourrilhes (DuNga)
      'Last Modify Date: Unknow
      '
      '**************************************************************
10        With Npclist(NpcIndex)
20            .AreasInfo.AreaID = 0
              
30            .AreasInfo.AreaPerteneceX = 0
40            .AreasInfo.AreaPerteneceY = 0
50            .AreasInfo.AreaReciveX = 0
60            .AreasInfo.AreaReciveY = 0
70        End With
          
80        Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
End Sub


