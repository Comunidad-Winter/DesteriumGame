Attribute VB_Name = "PathFinding"
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

'#######################################################
'PathFinding Module
'Coded By Gulfas Morgolock
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'
'Ore is an excellent engine for introducing you not only
'to online game programming but also to general
'game programming. I am convinced that Aaron Perkings, creator
'of ORE, did a great work. He made possible that a lot of
'people enjoy for no fee games made with his engine, and
'for me, this is something great.
'
'I'd really like to contribute to this work, and all the
'projects of free ore-based MMORPGs that are on the net.
'
'I did some basic improvements on the AI of the NPCs, I
'added pathfinding, so now, the npcs are able to avoid
'obstacles. I believe that this improvement was essential
'for the engine.
'
'I'd like to see this as my contribution to ORE project,
'I hope that someone finds this source code useful.
'So, please feel free to do whatever you want with my
'pathfinging module.
'
'I'd really appreciate that if you find this source code
'useful you mention my nickname on the credits of your
'program. But there is no obligation ;).
'
'.........................................................
'Note:
'There is a little problem, ORE refers to map arrays in a
'different manner that my pathfinding routines. When I wrote
'these routines, I did it without thinking in ORE, so in my
'program I refer to maps in the usual way I do it.
'
'For example, suppose we have:
'Map(1 to Y,1 to X) as MapBlock
'I usually use the first coordinate as Y, and
'the second one as X.
'
'ORE refers to maps in converse way, for example:
'Map(1 to X,1 to Y) as MapBlock. As you can see the
'roles of first and second coordinates are different
'that my routines
'
'#######################################################


Option Explicit

Private Const ROWS As Integer = 100
Private Const COLUMS As Integer = 100
Private Const MAXINT As Integer = 1000

Private Type tIntermidiateWork
    Known As Boolean
    DistV As Integer
    PrevV As tVertice
End Type

Dim TmpArray(1 To ROWS, 1 To COLUMS) As tIntermidiateWork

Dim TilePosY As Integer

Private Function Limites(ByVal vfila As Integer, ByVal vcolu As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    Limites = vcolu >= 1 And vcolu <= COLUMS And vfila >= 1 And vfila <= ROWS
End Function

Private Function IsWalkable(ByVal map As Integer, ByVal row As Integer, ByVal Col As Integer, ByVal NpcIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    IsWalkable = MapData(map, row, Col).Blocked = 0 And MapData(map, row, Col).NpcIndex = 0

20    If MapData(map, row, Col).Userindex <> 0 Then
30         If MapData(map, row, Col).Userindex <> Npclist(NpcIndex).PFINFO.TargetUser Then IsWalkable = False
40    End If

End Function

Private Sub ProcessAdjacents(ByVal MapIndex As Integer, ByRef t() As tIntermidiateWork, ByRef vfila As Integer, ByRef vcolu As Integer, ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim V As tVertice
          Dim j As Integer
          'Look to North
10        j = vfila - 1
20        If Limites(j, vcolu) Then
30                If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then
                          'Nos aseguramos que no hay un camino más corto
40                        If t(j, vcolu).DistV = MAXINT Then
                              'Actualizamos la tabla de calculos intermedios
50                            t(j, vcolu).DistV = t(vfila, vcolu).DistV + 1
60                            t(j, vcolu).PrevV.X = vcolu
70                            t(j, vcolu).PrevV.Y = vfila
                              'Mete el vertice en la cola
80                            V.X = vcolu
90                            V.Y = j
100                           Call Push(V)
110                       End If
120               End If
130       End If
140       j = vfila + 1
          'look to south
150       If Limites(j, vcolu) Then
160               If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then
                      'Nos aseguramos que no hay un camino más corto
170                   If t(j, vcolu).DistV = MAXINT Then
                          'Actualizamos la tabla de calculos intermedios
180                       t(j, vcolu).DistV = t(vfila, vcolu).DistV + 1
190                       t(j, vcolu).PrevV.X = vcolu
200                       t(j, vcolu).PrevV.Y = vfila
                          'Mete el vertice en la cola
210                       V.X = vcolu
220                       V.Y = j
230                       Call Push(V)
240                   End If
250               End If
260       End If
          'look to west
270       If Limites(vfila, vcolu - 1) Then
280               If IsWalkable(MapIndex, vfila, vcolu - 1, NpcIndex) Then
                      'Nos aseguramos que no hay un camino más corto
290                   If t(vfila, vcolu - 1).DistV = MAXINT Then
                          'Actualizamos la tabla de calculos intermedios
300                       t(vfila, vcolu - 1).DistV = t(vfila, vcolu).DistV + 1
310                       t(vfila, vcolu - 1).PrevV.X = vcolu
320                       t(vfila, vcolu - 1).PrevV.Y = vfila
                          'Mete el vertice en la cola
330                       V.X = vcolu - 1
340                       V.Y = vfila
350                       Call Push(V)
360                   End If
370               End If
380       End If
          'look to east
390       If Limites(vfila, vcolu + 1) Then
400               If IsWalkable(MapIndex, vfila, vcolu + 1, NpcIndex) Then
                      'Nos aseguramos que no hay un camino más corto
410                   If t(vfila, vcolu + 1).DistV = MAXINT Then
                          'Actualizamos la tabla de calculos intermedios
420                       t(vfila, vcolu + 1).DistV = t(vfila, vcolu).DistV + 1
430                       t(vfila, vcolu + 1).PrevV.X = vcolu
440                       t(vfila, vcolu + 1).PrevV.Y = vfila
                          'Mete el vertice en la cola
450                       V.X = vcolu + 1
460                       V.Y = vfila
470                       Call Push(V)
480                   End If
490               End If
500       End If
         
         
End Sub


Public Sub SeekPath(ByVal NpcIndex As Integer, Optional ByVal MaxSteps As Integer = 30)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      'This Sub seeks a path from the npclist(npcindex).pos
      'to the location NPCList(NpcIndex).PFINFO.Target.
      'The optional parameter MaxSteps is the maximum of steps
      'allowed for the path.
      '***************************************************

          Dim cur_npc_pos As tVertice
          Dim tar_npc_pos As tVertice
          Dim V As tVertice
          Dim NpcMap As Integer
          Dim steps As Integer
          
10        NpcMap = Npclist(NpcIndex).Pos.map
          
20        steps = 0
          
30        cur_npc_pos.X = Npclist(NpcIndex).Pos.Y
40        cur_npc_pos.Y = Npclist(NpcIndex).Pos.X
          
50        tar_npc_pos.X = Npclist(NpcIndex).PFINFO.Target.X '  UserList(NPCList(NpcIndex).PFINFO.TargetUser).Pos.X
60        tar_npc_pos.Y = Npclist(NpcIndex).PFINFO.Target.Y '  UserList(NPCList(NpcIndex).PFINFO.TargetUser).Pos.Y
          
70        Call InitializeTable(TmpArray, cur_npc_pos)
80        Call InitQueue
          
          'We add the first vertex to the Queue
90        Call Push(cur_npc_pos)
          
100       Do While (Not IsEmpty)
110           If steps > MaxSteps Then Exit Do
120           V = Pop
130           If V.X = tar_npc_pos.X And V.Y = tar_npc_pos.Y Then Exit Do
140           Call ProcessAdjacents(NpcMap, TmpArray, V.Y, V.X, NpcIndex)
150       Loop
          
160       Call MakePath(NpcIndex)

End Sub

Private Sub MakePath(ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      'Builds the path previously calculated
      '***************************************************

          Dim Pasos As Integer
          Dim miV As tVertice
          Dim i As Integer
          
10        Pasos = TmpArray(Npclist(NpcIndex).PFINFO.Target.Y, Npclist(NpcIndex).PFINFO.Target.X).DistV
20        Npclist(NpcIndex).PFINFO.PathLenght = Pasos
          
          
30        If Pasos = MAXINT Then
              'MsgBox "There is no path."
40            Npclist(NpcIndex).PFINFO.NoPath = True
50            Npclist(NpcIndex).PFINFO.PathLenght = 0
60            Exit Sub
70        End If
          
80        ReDim Npclist(NpcIndex).PFINFO.Path(0 To Pasos) As tVertice
          
90        miV.X = Npclist(NpcIndex).PFINFO.Target.X
100       miV.Y = Npclist(NpcIndex).PFINFO.Target.Y
          
110       For i = Pasos To 1 Step -1
120           Npclist(NpcIndex).PFINFO.Path(i) = miV
130           miV = TmpArray(miV.Y, miV.X).PrevV
140       Next i
          
150       Npclist(NpcIndex).PFINFO.CurPos = 1
160       Npclist(NpcIndex).PFINFO.NoPath = False
         
End Sub

Private Sub InitializeTable(ByRef t() As tIntermidiateWork, ByRef S As tVertice, Optional ByVal MaxSteps As Integer = 30)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      'Initialize the array where we calculate the path
      '***************************************************


      Dim j As Integer, k As Integer
      Const anymap = 1
10    For j = S.Y - MaxSteps To S.Y + MaxSteps
20        For k = S.X - MaxSteps To S.X + MaxSteps
30            If InMapBounds(anymap, j, k) Then
40                t(j, k).Known = False
50                t(j, k).DistV = MAXINT
60                t(j, k).PrevV.X = 0
70                t(j, k).PrevV.Y = 0
80            End If
90        Next
100   Next

110   t(S.Y, S.X).Known = False
120   t(S.Y, S.X).DistV = 0

End Sub

