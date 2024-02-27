Attribute VB_Name = "Statistics"
'**************************************************************
' modStatistics.bas - Takes statistics on the game for later study.
'
' Implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
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

Private Type trainningData
    startTick As Long
    trainningTime As Long
End Type

Private Type fragLvlRace
    matrix(1 To 50, 1 To 5) As Long
End Type

Private Type fragLvlLvl
    matrix(1 To 50, 1 To 50) As Long
End Type

Private trainningInfo() As trainningData

Private fragLvlRaceData(1 To 7) As fragLvlRace
Private fragLvlLvlData(1 To 7) As fragLvlLvl
Private fragAlignmentLvlData(1 To 50, 1 To 4) As Long

'Currency just in case.... chats are way TOO often...
Private keyOcurrencies(255) As Currency

Public Sub Initialize()
10        ReDim trainningInfo(1 To MaxUsers) As trainningData
End Sub

Public Sub UserConnected(ByVal Userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          'A new user connected, load it's trainning time count
10        trainningInfo(Userindex).trainningTime = val(GetVar(CharPath & UCase$(UserList(Userindex).Name) & ".chr", "RESEARCH", "TrainningTime", 30))
          
20        trainningInfo(Userindex).startTick = (GetTickCount() And &H7FFFFFFF)
End Sub

Public Sub UserDisconnected(ByVal Userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With trainningInfo(Userindex)
              'Update trainning time
20            .trainningTime = .trainningTime + ((GetTickCount() And &H7FFFFFFF) - .startTick) / 1000
              
30            .startTick = (GetTickCount() And &H7FFFFFFF)
              
              'Store info in char file
40            Call WriteVar(CharPath & UCase$(UserList(Userindex).Name) & ".chr", "RESEARCH", "TrainningTime", CStr(.trainningTime))
50        End With
End Sub

Public Sub UserLevelUp(ByVal Userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim handle As Integer
10        handle = FreeFile()
          
20        With trainningInfo(Userindex)
              'Log the data
30            Open App.Path & "\logs\statistics.log" For Append Shared As handle
              
40            Print #handle, UCase$(UserList(Userindex).Name) & " completó el nivel " & CStr(UserList(Userindex).Stats.ELV) & " en " & CStr(.trainningTime + ((GetTickCount() And &H7FFFFFFF) - .startTick) / 1000) & " segundos."
              
50            Close handle
              
              'Reset data
60            .trainningTime = 0
70            .startTick = (GetTickCount() And &H7FFFFFFF)
80        End With
End Sub

Public Sub StoreFrag(ByVal killer As Integer, ByVal victim As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim clase As Integer
          Dim raza As Integer
          Dim alignment As Integer
          
10        If UserList(victim).Stats.ELV > 50 Or UserList(killer).Stats.ELV > 50 Then Exit Sub
          
20        Select Case UserList(killer).clase
              Case eClass.Assasin
30                clase = 1
              
40            Case eClass.Bard
50                clase = 2
              
60            Case eClass.Mage
70                clase = 3
              
80            Case eClass.Paladin
90                clase = 4
              
100           Case eClass.Warrior
110               clase = 5
              
120           Case eClass.Cleric
130               clase = 6
              
140           Case eClass.Hunter
150               clase = 7
              
160           Case Else
170               Exit Sub
180       End Select
          
190       Select Case UserList(killer).raza
              Case eRaza.Elfo
200               raza = 1
              
210           Case eRaza.Drow
220               raza = 2
              
230           Case eRaza.Enano
240               raza = 3
              
250           Case eRaza.Gnomo
260               raza = 4
              
270           Case eRaza.Humano
280               raza = 5
              
290           Case Else
300               Exit Sub
310       End Select
          
320       If UserList(killer).Faccion.ArmadaReal Then
330           alignment = 1
340       ElseIf UserList(killer).Faccion.FuerzasCaos Then
350           alignment = 2
360       ElseIf criminal(killer) Then
370           alignment = 3
380       Else
390           alignment = 4
400       End If
          
410       fragLvlRaceData(clase).matrix(UserList(killer).Stats.ELV, raza) = fragLvlRaceData(clase).matrix(UserList(killer).Stats.ELV, raza) + 1
          
420       fragLvlLvlData(clase).matrix(UserList(killer).Stats.ELV, UserList(victim).Stats.ELV) = fragLvlLvlData(clase).matrix(UserList(killer).Stats.ELV, UserList(victim).Stats.ELV) + 1
          
430       fragAlignmentLvlData(UserList(killer).Stats.ELV, alignment) = fragAlignmentLvlData(UserList(killer).Stats.ELV, alignment) + 1
End Sub

Public Sub DumpStatistics()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim handle As Integer
10        handle = FreeFile()
          
          Dim line As String
          Dim i As Long
          Dim j As Long
          
20        Open App.Path & "\logs\frags.txt" For Output As handle
          
          'Save lvl vs lvl frag matrix for each class - we use GNU Octave's ASCII file format
          
30        Print #handle, "# name: fragLvlLvl_Ase"
40        Print #handle, "# type: matrix"
50        Print #handle, "# rows: 50"
60        Print #handle, "# columns: 50"
          
70        For j = 1 To 50
80            For i = 1 To 50
90                line = line & " " & CStr(fragLvlLvlData(1).matrix(i, j))
100           Next i
              
110           Print #handle, line
120           line = vbNullString
130       Next j
          
140       Print #handle, "# name: fragLvlLvl_Bar"
150       Print #handle, "# type: matrix"
160       Print #handle, "# rows: 50"
170       Print #handle, "# columns: 50"
          
180       For j = 1 To 50
190           For i = 1 To 50
200               line = line & " " & CStr(fragLvlLvlData(2).matrix(i, j))
210           Next i
              
220           Print #handle, line
230           line = vbNullString
240       Next j
          
250       Print #handle, "# name: fragLvlLvl_Mag"
260       Print #handle, "# type: matrix"
270       Print #handle, "# rows: 50"
280       Print #handle, "# columns: 50"
          
290       For j = 1 To 50
300           For i = 1 To 50
310               line = line & " " & CStr(fragLvlLvlData(3).matrix(i, j))
320           Next i
              
330           Print #handle, line
340           line = vbNullString
350       Next j
          
360       Print #handle, "# name: fragLvlLvl_Pal"
370       Print #handle, "# type: matrix"
380       Print #handle, "# rows: 50"
390       Print #handle, "# columns: 50"
          
400       For j = 1 To 50
410           For i = 1 To 50
420               line = line & " " & CStr(fragLvlLvlData(4).matrix(i, j))
430           Next i
              
440           Print #handle, line
450           line = vbNullString
460       Next j
          
470       Print #handle, "# name: fragLvlLvl_Gue"
480       Print #handle, "# type: matrix"
490       Print #handle, "# rows: 50"
500       Print #handle, "# columns: 50"
          
510       For j = 1 To 50
520           For i = 1 To 50
530               line = line & " " & CStr(fragLvlLvlData(5).matrix(i, j))
540           Next i
              
550           Print #handle, line
560           line = vbNullString
570       Next j
          
580       Print #handle, "# name: fragLvlLvl_Cle"
590       Print #handle, "# type: matrix"
600       Print #handle, "# rows: 50"
610       Print #handle, "# columns: 50"
          
620       For j = 1 To 50
630           For i = 1 To 50
640               line = line & " " & CStr(fragLvlLvlData(6).matrix(i, j))
650           Next i
              
660           Print #handle, line
670           line = vbNullString
680       Next j
          
690       Print #handle, "# name: fragLvlLvl_Caz"
700       Print #handle, "# type: matrix"
710       Print #handle, "# rows: 50"
720       Print #handle, "# columns: 50"
          
730       For j = 1 To 50
740           For i = 1 To 50
750               line = line & " " & CStr(fragLvlLvlData(7).matrix(i, j))
760           Next i
              
770           Print #handle, line
780           line = vbNullString
790       Next j
          
          
          
          
          
          'Save lvl vs race frag matrix for each class - we use GNU Octave's ASCII file format
          
800       Print #handle, "# name: fragLvlRace_Ase"
810       Print #handle, "# type: matrix"
820       Print #handle, "# rows: 5"
830       Print #handle, "# columns: 50"
          
840       For j = 1 To 5
850           For i = 1 To 50
860               line = line & " " & CStr(fragLvlRaceData(1).matrix(i, j))
870           Next i
              
880           Print #handle, line
890           line = vbNullString
900       Next j
          
910       Print #handle, "# name: fragLvlRace_Bar"
920       Print #handle, "# type: matrix"
930       Print #handle, "# rows: 5"
940       Print #handle, "# columns: 50"
          
950       For j = 1 To 5
960           For i = 1 To 50
970               line = line & " " & CStr(fragLvlRaceData(2).matrix(i, j))
980           Next i
              
990           Print #handle, line
1000          line = vbNullString
1010      Next j
          
1020      Print #handle, "# name: fragLvlRace_Mag"
1030      Print #handle, "# type: matrix"
1040      Print #handle, "# rows: 5"
1050      Print #handle, "# columns: 50"
          
1060      For j = 1 To 5
1070          For i = 1 To 50
1080              line = line & " " & CStr(fragLvlRaceData(3).matrix(i, j))
1090          Next i
              
1100          Print #handle, line
1110          line = vbNullString
1120      Next j
          
1130      Print #handle, "# name: fragLvlRace_Pal"
1140      Print #handle, "# type: matrix"
1150      Print #handle, "# rows: 5"
1160      Print #handle, "# columns: 50"
          
1170      For j = 1 To 5
1180          For i = 1 To 50
1190              line = line & " " & CStr(fragLvlRaceData(4).matrix(i, j))
1200          Next i
              
1210          Print #handle, line
1220          line = vbNullString
1230      Next j
          
1240      Print #handle, "# name: fragLvlRace_Gue"
1250      Print #handle, "# type: matrix"
1260      Print #handle, "# rows: 5"
1270      Print #handle, "# columns: 50"
          
1280      For j = 1 To 5
1290          For i = 1 To 50
1300              line = line & " " & CStr(fragLvlRaceData(5).matrix(i, j))
1310          Next i
              
1320          Print #handle, line
1330          line = vbNullString
1340      Next j
          
1350      Print #handle, "# name: fragLvlRace_Cle"
1360      Print #handle, "# type: matrix"
1370      Print #handle, "# rows: 5"
1380      Print #handle, "# columns: 50"
          
1390      For j = 1 To 5
1400          For i = 1 To 50
1410              line = line & " " & CStr(fragLvlRaceData(6).matrix(i, j))
1420          Next i
              
1430          Print #handle, line
1440          line = vbNullString
1450      Next j
          
1460      Print #handle, "# name: fragLvlRace_Caz"
1470      Print #handle, "# type: matrix"
1480      Print #handle, "# rows: 5"
1490      Print #handle, "# columns: 50"
          
1500      For j = 1 To 5
1510          For i = 1 To 50
1520              line = line & " " & CStr(fragLvlRaceData(7).matrix(i, j))
1530          Next i
              
1540          Print #handle, line
1550          line = vbNullString
1560      Next j
          
          
          
          
          
          
          'Save lvl vs class frag matrix for each race - we use GNU Octave's ASCII file format
          
1570      Print #handle, "# name: fragLvlClass_Elf"
1580      Print #handle, "# type: matrix"
1590      Print #handle, "# rows: 7"
1600      Print #handle, "# columns: 50"
          
1610      For j = 1 To 7
1620          For i = 1 To 50
1630              line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 1))
1640          Next i
              
1650          Print #handle, line
1660          line = vbNullString
1670      Next j
          
1680      Print #handle, "# name: fragLvlClass_Dar"
1690      Print #handle, "# type: matrix"
1700      Print #handle, "# rows: 7"
1710      Print #handle, "# columns: 50"
          
1720      For j = 1 To 7
1730          For i = 1 To 50
1740              line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 2))
1750          Next i
              
1760          Print #handle, line
1770          line = vbNullString
1780      Next j
          
1790      Print #handle, "# name: fragLvlClass_Dwa"
1800      Print #handle, "# type: matrix"
1810      Print #handle, "# rows: 7"
1820      Print #handle, "# columns: 50"
          
1830      For j = 1 To 7
1840          For i = 1 To 50
1850              line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 3))
1860          Next i
              
1870          Print #handle, line
1880          line = vbNullString
1890      Next j
          
1900      Print #handle, "# name: fragLvlClass_Gno"
1910      Print #handle, "# type: matrix"
1920      Print #handle, "# rows: 7"
1930      Print #handle, "# columns: 50"
          
1940      For j = 1 To 7
1950          For i = 1 To 50
1960              line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 4))
1970          Next i
              
1980          Print #handle, line
1990          line = vbNullString
2000      Next j
          
2010      Print #handle, "# name: fragLvlClass_Hum"
2020      Print #handle, "# type: matrix"
2030      Print #handle, "# rows: 7"
2040      Print #handle, "# columns: 50"
          
2050      For j = 1 To 7
2060          For i = 1 To 50
2070              line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 5))
2080          Next i
              
2090          Print #handle, line
2100          line = vbNullString
2110      Next j
          
          
          
          
          'Save lvl vs alignment frag matrix for each race - we use GNU Octave's ASCII file format
          
2120      Print #handle, "# name: fragAlignmentLvl"
2130      Print #handle, "# type: matrix"
2140      Print #handle, "# rows: 4"
2150      Print #handle, "# columns: 50"
          
2160      For j = 1 To 4
2170          For i = 1 To 50
2180              line = line & " " & CStr(fragAlignmentLvlData(i, j))
2190          Next i
              
2200          Print #handle, line
2210          line = vbNullString
2220      Next j
          
2230      Close handle
          
          
          
          'Dump Chat statistics
2240      handle = FreeFile()
          
2250      Open App.Path & "\logs\huffman.log" For Output As handle
          
          Dim Total As Currency
          
          'Compute total characters
2260      For i = 0 To 255
2270          Total = Total + keyOcurrencies(i)
2280      Next i
          
          'Show each character's ocurrencies
2290      If Total <> 0 Then
2300          For i = 0 To 255
2310              Print #handle, CStr(i) & "    " & CStr(Round(keyOcurrencies(i) / Total, 8))
2320          Next i
2330      End If
          
2340      Print #handle, "TOTAL =    " & CStr(Total)
          
2350      Close handle
End Sub

Public Sub ParseChat(ByRef S As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim i As Long
          Dim key As Integer
          
10        For i = 1 To Len(S)
20            key = Asc(mid$(S, i, 1))
              
30            keyOcurrencies(key) = keyOcurrencies(key) + 1
40        Next i
          
          'Add a NULL-terminated to consider that possibility too....
50        keyOcurrencies(0) = keyOcurrencies(0) + 1
End Sub
