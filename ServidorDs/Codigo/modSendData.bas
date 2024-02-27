Attribute VB_Name = "modSendData"
'**************************************************************
' SendData.bas - Has all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' Implemented by Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@gmail.com)
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

''
' Contains all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20070107

Option Explicit

Public ElmasbuscadoFusion As String

Public Enum SendTarget
    ToAll = 1
    toMap
    ToPCArea
    ToAllButIndex
    ToMapButIndex
    ToGM
    ToNPCArea
    ToGuildMembers
    ToAdmins
    ToPCAreaButIndex
    ToAdminsAreaButConsejeros
    ToDiosesYclan
    ToConsejo
    ToClanArea
    ToConsejoCaos
    ToRolesMasters
    ToDeadArea
    ToCiudadanos
    ToCriminales
    ToPartyArea
    ToReal
    ToCaos
    ToCiudadanosYRMs
    ToCriminalesYRMs
    ToRealYRMs
    ToCaosYRMs
    ToHigherAdmins
    ToGMsAreaButRmsOrCounselors
    ToUsersAreaButGMs
    ToUsersAndRmsAndCounselorsAreaButGMs
End Enum

Public Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndData As String)
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus) - Rewrite of original
      'Last Modify Date: 01/08/2007
      'Last modified by: (liquid)
      '**************************************************************
10    On Error Resume Next
          Dim LoopC As Long
          Dim map As Integer
          
20        Select Case sndRoute
              Case SendTarget.ToPCArea
30                Call SendToUserArea(sndIndex, sndData)
40                Exit Sub
              
50            Case SendTarget.ToAdmins
60                For LoopC = 1 To LastUser
70                    If UserList(LoopC).ConnID <> -1 Then
80                        If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
90                            Call EnviarDatosASlot(LoopC, sndData)
100                      End If
110                   End If
120               Next LoopC
130               Exit Sub
              
140           Case SendTarget.ToAll
150               For LoopC = 1 To LastUser
160                   If UserList(LoopC).ConnID <> -1 Then
170                       If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
180                           Call EnviarDatosASlot(LoopC, sndData)
190                       End If
200                   End If
210               Next LoopC
220               Exit Sub
              
230           Case SendTarget.ToAllButIndex
240               For LoopC = 1 To LastUser
250                   If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
260                       If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
270                           Call EnviarDatosASlot(LoopC, sndData)
280                       End If
290                   End If
300               Next LoopC
310               Exit Sub
              
320           Case SendTarget.toMap
330               Call SendToMap(sndIndex, sndData)
340               Exit Sub
                
350           Case SendTarget.ToMapButIndex
360               Call SendToMapButIndex(sndIndex, sndData)
370               Exit Sub
              
380           Case SendTarget.ToGuildMembers
390               LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
400               While LoopC > 0
410                   If (UserList(LoopC).ConnID <> -1) Then
420                       Call EnviarDatosASlot(LoopC, sndData)
430                   End If
440                   LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
450               Wend
460               Exit Sub
              
470           Case SendTarget.ToDeadArea
480               Call SendToDeadUserArea(sndIndex, sndData)
490               Exit Sub
              
500           Case SendTarget.ToPCAreaButIndex
510               Call SendToUserAreaButindex(sndIndex, sndData)
520               Exit Sub
              
530           Case SendTarget.ToClanArea
540               Call SendToUserGuildArea(sndIndex, sndData)
550               Exit Sub
              
560           Case SendTarget.ToPartyArea
570               Call SendToUserPartyArea(sndIndex, sndData)
580               Exit Sub
              
590           Case SendTarget.ToAdminsAreaButConsejeros
600               Call SendToAdminsButConsejerosArea(sndIndex, sndData)
610               Exit Sub
              
620           Case SendTarget.ToNPCArea
630               Call SendToNpcArea(sndIndex, sndData)
640               Exit Sub
              
650           Case SendTarget.ToDiosesYclan
660               LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
670               While LoopC > 0
680                   If (UserList(LoopC).ConnID <> -1) Then
690                       Call EnviarDatosASlot(LoopC, sndData)
700                   End If
710                   LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
720               Wend
                  
730               LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
740               While LoopC > 0
750                   If (UserList(LoopC).ConnID <> -1) Then
760                       Call EnviarDatosASlot(LoopC, sndData)
770                   End If
780                   LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
790               Wend
                  
800               Exit Sub
              
810           Case SendTarget.ToConsejo
820               For LoopC = 1 To LastUser
830                   If (UserList(LoopC).ConnID <> -1) Then
840                       If UserList(LoopC).flags.Privilegios And PlayerType.RoyalCouncil Then
850                           Call EnviarDatosASlot(LoopC, sndData)
860                       End If
870                   End If
880               Next LoopC
890               Exit Sub
              
900           Case SendTarget.ToConsejoCaos
910               For LoopC = 1 To LastUser
920                   If (UserList(LoopC).ConnID <> -1) Then
930                       If UserList(LoopC).flags.Privilegios And PlayerType.ChaosCouncil Then
940                           Call EnviarDatosASlot(LoopC, sndData)
950                       End If
960                   End If
970               Next LoopC
980               Exit Sub
              
990           Case SendTarget.ToRolesMasters
1000              For LoopC = 1 To LastUser
1010                  If (UserList(LoopC).ConnID <> -1) Then
1020                      If UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster Then
1030                          Call EnviarDatosASlot(LoopC, sndData)
1040                      End If
1050                  End If
1060              Next LoopC
1070              Exit Sub
              
1080          Case SendTarget.ToCiudadanos
1090              For LoopC = 1 To LastUser
1100                  If (UserList(LoopC).ConnID <> -1) Then
1110                      If Not criminal(LoopC) Then
1120                          Call EnviarDatosASlot(LoopC, sndData)
1130                      End If
1140                  End If
1150              Next LoopC
1160              Exit Sub
              
1170          Case SendTarget.ToCriminales
1180              For LoopC = 1 To LastUser
1190                  If (UserList(LoopC).ConnID <> -1) Then
1200                      If criminal(LoopC) Then
1210                          Call EnviarDatosASlot(LoopC, sndData)
1220                      End If
1230                  End If
1240              Next LoopC
1250              Exit Sub
              
1260          Case SendTarget.ToReal
1270              For LoopC = 1 To LastUser
1280                  If (UserList(LoopC).ConnID <> -1) Then
1290                      If UserList(LoopC).Faccion.ArmadaReal = 1 Then
1300                          Call EnviarDatosASlot(LoopC, sndData)
1310                      End If
1320                  End If
1330              Next LoopC
1340              Exit Sub
              
1350          Case SendTarget.ToCaos
1360              For LoopC = 1 To LastUser
1370                  If (UserList(LoopC).ConnID <> -1) Then
1380                      If UserList(LoopC).Faccion.FuerzasCaos = 1 Then
1390                          Call EnviarDatosASlot(LoopC, sndData)
1400                      End If
1410                  End If
1420              Next LoopC
1430              Exit Sub
              
1440          Case SendTarget.ToCiudadanosYRMs
1450              For LoopC = 1 To LastUser
1460                  If (UserList(LoopC).ConnID <> -1) Then
1470                      If Not criminal(LoopC) Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
1480                          Call EnviarDatosASlot(LoopC, sndData)
1490                      End If
1500                  End If
1510              Next LoopC
1520              Exit Sub
              
1530          Case SendTarget.ToCriminalesYRMs
1540              For LoopC = 1 To LastUser
1550                  If (UserList(LoopC).ConnID <> -1) Then
1560                      If criminal(LoopC) Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
1570                          Call EnviarDatosASlot(LoopC, sndData)
1580                      End If
1590                  End If
1600              Next LoopC
1610              Exit Sub
              
1620          Case SendTarget.ToRealYRMs
1630              For LoopC = 1 To LastUser
1640                  If (UserList(LoopC).ConnID <> -1) Then
1650                      If UserList(LoopC).Faccion.ArmadaReal = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
1660                          Call EnviarDatosASlot(LoopC, sndData)
1670                      End If
1680                  End If
1690              Next LoopC
1700              Exit Sub
              
1710          Case SendTarget.ToCaosYRMs
1720              For LoopC = 1 To LastUser
1730                  If (UserList(LoopC).ConnID <> -1) Then
1740                      If UserList(LoopC).Faccion.FuerzasCaos = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
1750                          Call EnviarDatosASlot(LoopC, sndData)
1760                      End If
1770                  End If
1780              Next LoopC
1790              Exit Sub
              
1800          Case SendTarget.ToHigherAdmins
1810              For LoopC = 1 To LastUser
1820                  If UserList(LoopC).ConnID <> -1 Then
1830                      If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
1840                          Call EnviarDatosASlot(LoopC, sndData)
1850                     End If
1860                  End If
1870              Next LoopC
1880              Exit Sub
                  
1890          Case SendTarget.ToGMsAreaButRmsOrCounselors
1900              Call SendToGMsAreaButRmsOrCounselors(sndIndex, sndData)
1910              Exit Sub
                  
1920          Case SendTarget.ToUsersAreaButGMs
1930              Call SendToUsersAreaButGMs(sndIndex, sndData)
1940              Exit Sub
1950          Case SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs
1960              Call SendToUsersAndRmsAndCounselorsAreaButGMs(sndIndex, sndData)
1970              Exit Sub
1980      End Select
End Sub

Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal sdData As String)
      '**************************************************************
      'Author: Lucio N. Tourrilhes (DuNga)
      'Last Modify Date: Unknow
      '
      '**************************************************************
          Dim LoopC As Long
          Dim tempIndex As Integer
          
          Dim map As Integer
          Dim AreaX As Integer
          Dim AreaY As Integer
          
10        map = UserList(UserIndex).Pos.map
20        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
30        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
          
40        If Not MapaValido(map) Then Exit Sub
          
50        For LoopC = 1 To ConnGroups(map).CountEntrys
60            tempIndex = ConnGroups(map).UserEntrys(LoopC)
              
70            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
80                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
90                    If UserList(tempIndex).ConnIDValida Then
100                       Call EnviarDatosASlot(tempIndex, sdData)
110                   End If
120               End If
130           End If
140       Next LoopC
End Sub

Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal sdData As String)
      '**************************************************************
      'Author: Lucio N. Tourrilhes (DuNga)
      'Last Modify Date: Unknow
      '
      '**************************************************************
          Dim LoopC As Long
          Dim TempInt As Integer
          Dim tempIndex As Integer
          
          Dim map As Integer
          Dim AreaX As Integer
          Dim AreaY As Integer
          
10        map = UserList(UserIndex).Pos.map
20        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
30        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY

40        If Not MapaValido(map) Then Exit Sub
          
50        For LoopC = 1 To ConnGroups(map).CountEntrys
60            tempIndex = ConnGroups(map).UserEntrys(LoopC)
                  
70            TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
80            If TempInt Then  'Esta en el area?
90                TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
100               If TempInt Then
110                   If tempIndex <> UserIndex Then
120                       If UserList(tempIndex).ConnIDValida Then
130                           Call EnviarDatosASlot(tempIndex, sdData)
140                       End If
150                   End If
160               End If
170           End If
180       Next LoopC
End Sub

Private Sub SendToDeadUserArea(ByVal UserIndex As Integer, ByVal sdData As String)
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: Unknow
      '
      '**************************************************************
          Dim LoopC As Long
          Dim tempIndex As Integer
          
          Dim map As Integer
          Dim AreaX As Integer
          Dim AreaY As Integer
          
10        map = UserList(UserIndex).Pos.map
20        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
30        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
          
40        If Not MapaValido(map) Then Exit Sub
          
50        For LoopC = 1 To ConnGroups(map).CountEntrys
60            tempIndex = ConnGroups(map).UserEntrys(LoopC)
              
70            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
80                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                      'Dead and admins read
90                    If UserList(tempIndex).ConnIDValida = True And (UserList(tempIndex).flags.Muerto = 1 Or (UserList(tempIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) <> 0) Then
100                       Call EnviarDatosASlot(tempIndex, sdData)
110                   End If
120               End If
130           End If
140       Next LoopC
End Sub

Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal sdData As String)
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: Unknow
      '
      '**************************************************************
          Dim LoopC As Long
          Dim tempIndex As Integer
          
          Dim map As Integer
          Dim AreaX As Integer
          Dim AreaY As Integer
          
10        map = UserList(UserIndex).Pos.map
20        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
30        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
          
40        If Not MapaValido(map) Then Exit Sub
          
50        If UserList(UserIndex).GuildIndex = 0 Then Exit Sub
          
60        For LoopC = 1 To ConnGroups(map).CountEntrys
70            tempIndex = ConnGroups(map).UserEntrys(LoopC)
              
80            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
90                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
100                   If UserList(tempIndex).ConnIDValida And (UserList(tempIndex).GuildIndex = UserList(UserIndex).GuildIndex Or ((UserList(tempIndex).flags.Privilegios And PlayerType.Dios) And (UserList(tempIndex).flags.Privilegios And PlayerType.RoleMaster) = 0)) Then
110                       Call EnviarDatosASlot(tempIndex, sdData)
120                   End If
130               End If
140           End If
150       Next LoopC
End Sub

Private Sub SendToUserPartyArea(ByVal UserIndex As Integer, ByVal sdData As String)
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: Unknow
      '
      '**************************************************************
          Dim LoopC As Long
          Dim tempIndex As Integer
          
          Dim map As Integer
          Dim AreaX As Integer
          Dim AreaY As Integer
          
10        map = UserList(UserIndex).Pos.map
20        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
30        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
          
40        If Not MapaValido(map) Then Exit Sub
          
50        If UserList(UserIndex).GroupIndex = 0 Then Exit Sub
          
60        For LoopC = 1 To ConnGroups(map).CountEntrys
70            tempIndex = ConnGroups(map).UserEntrys(LoopC)
              
80            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
90                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
100                   If UserList(tempIndex).ConnIDValida And UserList(tempIndex).GroupIndex = UserList(UserIndex).GroupIndex Then
110                       Call EnviarDatosASlot(tempIndex, sdData)
120                   End If
130               End If
140           End If
150       Next LoopC
End Sub

Private Sub SendToAdminsButConsejerosArea(ByVal UserIndex As Integer, ByVal sdData As String)
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: Unknow
      '
      '**************************************************************
          Dim LoopC As Long
          Dim tempIndex As Integer
          
          Dim map As Integer
          Dim AreaX As Integer
          Dim AreaY As Integer
          
10        map = UserList(UserIndex).Pos.map
20        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
30        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
          
40        If Not MapaValido(map) Then Exit Sub
          
50        For LoopC = 1 To ConnGroups(map).CountEntrys
60            tempIndex = ConnGroups(map).UserEntrys(LoopC)
              
70            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
80                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
90                    If UserList(tempIndex).ConnIDValida Then
100                       If UserList(tempIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then _
                              Call EnviarDatosASlot(tempIndex, sdData)
110                   End If
120               End If
130           End If
140       Next LoopC
End Sub

Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sdData As String)
      '**************************************************************
      'Author: Lucio N. Tourrilhes (DuNga)
      'Last Modify Date: Unknow
      '
      '**************************************************************
          Dim LoopC As Long
          Dim TempInt As Integer
          Dim tempIndex As Integer
          
          Dim map As Integer
          Dim AreaX As Integer
          Dim AreaY As Integer
          
10        map = Npclist(NpcIndex).Pos.map
20        AreaX = Npclist(NpcIndex).AreasInfo.AreaPerteneceX
30        AreaY = Npclist(NpcIndex).AreasInfo.AreaPerteneceY
          
40        If Not MapaValido(map) Then Exit Sub
          
50        For LoopC = 1 To ConnGroups(map).CountEntrys
60            tempIndex = ConnGroups(map).UserEntrys(LoopC)
              
70            TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
80            If TempInt Then  'Esta en el area?
90                TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
100               If TempInt Then
110                   If UserList(tempIndex).ConnIDValida Then
120                       Call EnviarDatosASlot(tempIndex, sdData)
130                   End If
140               End If
150           End If
160       Next LoopC
End Sub

Public Sub SendToAreaByPos(ByVal map As Integer, ByVal AreaX As Integer, ByVal AreaY As Integer, ByVal sdData As String)
      '**************************************************************
      'Author: Lucio N. Tourrilhes (DuNga)
      'Last Modify Date: Unknow
      '
      '**************************************************************
          Dim LoopC As Long
          Dim TempInt As Integer
          Dim tempIndex As Integer
          
10        AreaX = 2 ^ (AreaX \ 9)
20        AreaY = 2 ^ (AreaY \ 9)
          
30        If Not MapaValido(map) Then Exit Sub

40        For LoopC = 1 To ConnGroups(map).CountEntrys
50            tempIndex = ConnGroups(map).UserEntrys(LoopC)
                  
60            TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
70            If TempInt Then  'Esta en el area?
80                TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
90                If TempInt Then
100                   If UserList(tempIndex).ConnIDValida Then
110                       Call EnviarDatosASlot(tempIndex, sdData)
120                   End If
130               End If
140           End If
150       Next LoopC
End Sub

Public Sub SendToMap(ByVal map As Integer, ByVal sdData As String)
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 5/24/2007
      '
      '**************************************************************
          Dim LoopC As Long
          Dim tempIndex As Integer
          
10        If Not MapaValido(map) Then Exit Sub

20        For LoopC = 1 To ConnGroups(map).CountEntrys
30            tempIndex = ConnGroups(map).UserEntrys(LoopC)
              
40            If UserList(tempIndex).ConnIDValida Then
50                Call EnviarDatosASlot(tempIndex, sdData)
60            End If
70        Next LoopC
End Sub

Public Sub SendToMapButIndex(ByVal UserIndex As Integer, ByVal sdData As String)
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 5/24/2007
      '
      '**************************************************************
          Dim LoopC As Long
          Dim map As Integer
          Dim tempIndex As Integer
          
10        map = UserList(UserIndex).Pos.map
          
20        If Not MapaValido(map) Then Exit Sub

30        For LoopC = 1 To ConnGroups(map).CountEntrys
40            tempIndex = ConnGroups(map).UserEntrys(LoopC)
              
50            If tempIndex <> UserIndex And UserList(tempIndex).ConnIDValida Then
60                Call EnviarDatosASlot(tempIndex, sdData)
70            End If
80        Next LoopC
End Sub

Private Sub SendToGMsAreaButRmsOrCounselors(ByVal UserIndex As Integer, ByVal sdData As String)
      '**************************************************************
      'Author: Torres Patricio(Pato)
      'Last Modify Date: 12/02/2010
      '12/02/2010: ZaMa - Restrinjo solo a dioses, admins y gms.
      '15/02/2010: ZaMa - Cambio el nombre de la funcion (viejo: ToGmsArea, nuevo: ToGmsAreaButRMsOrCounselors)
      '**************************************************************
          Dim LoopC As Long
          Dim tempIndex As Integer
          
          Dim map As Integer
          Dim AreaX As Integer
          Dim AreaY As Integer
          
10        map = UserList(UserIndex).Pos.map
20        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
30        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
          
40        If Not MapaValido(map) Then Exit Sub
          
50        For LoopC = 1 To ConnGroups(map).CountEntrys
60            tempIndex = ConnGroups(map).UserEntrys(LoopC)
              
70            With UserList(tempIndex)
80                If .AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
90                    If .AreasInfo.AreaReciveY And AreaY Then
100                       If .ConnIDValida Then
                              ' Exclusivo para dioses, admins y gms
110                           If (.flags.Privilegios And Not PlayerType.User And Not PlayerType.Consejero _
                                  And Not PlayerType.RoleMaster) = .flags.Privilegios Then
120                               Call EnviarDatosASlot(tempIndex, sdData)
130                           End If
140                       End If
150                   End If
160               End If
170           End With
180       Next LoopC
End Sub

Private Sub SendToUsersAreaButGMs(ByVal UserIndex As Integer, ByVal sdData As String)
      '**************************************************************
      'Author: Torres Patricio(Pato)
      'Last Modify Date: 10/17/2009
      '
      '**************************************************************
          Dim LoopC As Long
          Dim tempIndex As Integer
          
          Dim map As Integer
          Dim AreaX As Integer
          Dim AreaY As Integer
          
10        map = UserList(UserIndex).Pos.map
20        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
30        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
          
40        If Not MapaValido(map) Then Exit Sub
          
50        For LoopC = 1 To ConnGroups(map).CountEntrys
60            tempIndex = ConnGroups(map).UserEntrys(LoopC)
              
70            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
80                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
90                    If UserList(tempIndex).ConnIDValida Then
100                       If UserList(tempIndex).flags.Privilegios And PlayerType.User Then
110                           Call EnviarDatosASlot(tempIndex, sdData)
120                       End If
130                   End If
140               End If
150           End If
160       Next LoopC
End Sub

Private Sub SendToUsersAndRmsAndCounselorsAreaButGMs(ByVal UserIndex As Integer, ByVal sdData As String)
      '**************************************************************
      'Author: Torres Patricio(Pato)
      'Last Modify Date: 10/17/2009
      '
      '**************************************************************
          Dim LoopC As Long
          Dim tempIndex As Integer
          
          Dim map As Integer
          Dim AreaX As Integer
          Dim AreaY As Integer
          
10        map = UserList(UserIndex).Pos.map
20        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
30        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
          
40        If Not MapaValido(map) Then Exit Sub
          
50        For LoopC = 1 To ConnGroups(map).CountEntrys
60            tempIndex = ConnGroups(map).UserEntrys(LoopC)
              
70            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
80                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
90                    If UserList(tempIndex).ConnIDValida Then
100                       If UserList(tempIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
110                           Call EnviarDatosASlot(tempIndex, sdData)
120                       End If
130                   End If
140               End If
150           End If
160       Next LoopC
End Sub

Public Sub AlertarFaccionarios(ByVal UserIndex As Integer)
      '**************************************************************
      'Author: ZaMa
      'Last Modify Date: 17/11/2009
      'Alerta a los faccionarios, dandoles una orientacion
      '**************************************************************
          Dim LoopC As Long
          Dim tempIndex As Integer
          Dim map As Integer
          Dim Font As FontTypeNames
          
10        If esCaos(UserIndex) Then
20            Font = FontTypeNames.FONTTYPE_CONSEJOCAOS
30        Else
40            Font = FontTypeNames.FONTTYPE_CONSEJO
50        End If
          
60        map = UserList(UserIndex).Pos.map
          
70        If Not MapaValido(map) Then Exit Sub

80        For LoopC = 1 To ConnGroups(map).CountEntrys
90            tempIndex = ConnGroups(map).UserEntrys(LoopC)
              
100           If UserList(tempIndex).ConnIDValida Then
110               If tempIndex <> UserIndex Then
                      ' Solo se envia a los de la misma faccion
120                   If SameFaccion(UserIndex, tempIndex) Then
130                       Call EnviarDatosASlot(tempIndex, _
                               PrepareMessageConsoleMsg("Escuchas el llamado de un compañero que proviene del " & _
                               GetDireccion(UserIndex, tempIndex), Font))
140                   End If
150               End If
160           End If
170       Next LoopC

End Sub
