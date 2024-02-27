Attribute VB_Name = "AI"
'Argentum Online 0.12.2
'Copyright (C) 2002 Mï¿½rquez Pablo Ignacio
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
'Calle 3 nï¿½mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Cï¿½digo Postal 1900
'Pablo Ignacio Mï¿½rquez

Option Explicit

Public Enum TipoAI
    ESTATICO = 1
    MueveAlAzar = 2
    NpcMaloAtacaUsersBuenos = 3
    NPCDEFENSA = 4
    GuardiasAtacanCriminales = 5
    NpcObjeto = 6
    SigueAmo = 8
    NpcAtacaNpc = 9
    NpcPathfinding = 10
    NpcDagaRusa = 11
    
    'Pretorianos
    SacerdotePretorianoAi = 20
    GuerreroPretorianoAi = 21
    MagoPretorianoAi = 22
    CazadorPretorianoAi = 23
    ReyPretoriano = 24
End Enum

Public Const ELEMENTALFUEGO As Integer = 93
Public Const ELEMENTALTIERRA As Integer = 94
Public Const ELEMENTALAGUA As Integer = 92
Public Const ZOMBIE As Integer = 115
Public Const LOBO As Integer = 512
Public Const OSOS As Integer = 78

'Damos a los NPCs el mismo rango de visiï¿½n que un PJ
Public Const RANGO_VISION_X As Byte = 8
Public Const RANGO_VISION_Y As Byte = 6

'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'                        Modulo AI_NPC
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'AI de los NPC
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½

Private Sub GuardiasAI(ByVal NpcIndex As Integer, ByVal DelCaos As Boolean)
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 12/01/2010 (ZaMa)
      '14/09/2009: ZaMa - Now npcs don't atack protected users.
      '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
      '***************************************************
          Dim nPos As WorldPos
          Dim headingloop As Byte
          Dim UI As Integer
          Dim UserProtected As Boolean
          
   On Error GoTo GuardiasAI_Error

10        With Npclist(NpcIndex)
20            For headingloop = eHeading.NORTH To eHeading.WEST
30                nPos = .Pos
40                If .flags.Inmovilizado = 0 Or headingloop = .Char.Heading Then
50                    Call HeadtoPos(headingloop, nPos)
60                    If InMapBounds(nPos.map, nPos.X, nPos.Y) Then
70                        UI = MapData(nPos.map, nPos.X, nPos.Y).UserIndex
80                        If UI > 0 Then
90                            UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
100                           UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.EnConsulta
                              
110                           If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                                  'ï¿½ES CRIMINAL?
120                               If Not DelCaos Then
130                                   If criminal(UI) Then
140                                       If NpcAtacaUser(NpcIndex, UI) Then
150                                           Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
160                                       End If
170                                       Exit Sub
180                                   ElseIf .flags.AttackedBy = UserList(UI).Name And Not .flags.Follow Then
                                          
190                                       If NpcAtacaUser(NpcIndex, UI) Then
200                                           Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
210                                       End If
220                                       Exit Sub
230                                   End If
240                               Else
250                                   If Not criminal(UI) Then
260                                       If NpcAtacaUser(NpcIndex, UI) Then
270                                           Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
280                                       End If
290                                       Exit Sub
300                                   ElseIf .flags.AttackedBy = UserList(UI).Name And Not .flags.Follow Then
                                            
310                                       If NpcAtacaUser(NpcIndex, UI) Then
320                                           Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
330                                       End If
340                                       Exit Sub
350                                   End If
360                               End If
370                           End If
380                       End If
390                   End If
400               End If  'not inmovil
410           Next headingloop
420       End With
          
430       Call RestoreOldMovement(NpcIndex)

   On Error GoTo 0
   Exit Sub

GuardiasAI_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure GuardiasAI of Módulo AI in line " & Erl
End Sub

''
' Handles the evil npcs' artificial intelligency.
'
' @param NpcIndex Specifies reference to the npc
Private Sub HostilMalvadoAI(ByVal NpcIndex As Integer)
10        On Error GoTo HostilMalvadoAI_Error
      '**************************************************************
      'Author: Unknown
      'Last Modify Date: 12/01/2010 (ZaMa)
      '28/04/2009: ZaMa - Now those NPCs who doble attack, have 50% of posibility of casting a spell on user.
      '14/09/200*: ZaMa - Now npcs don't atack protected users.
      '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
      '**************************************************************
          Dim nPos As WorldPos
          Dim headingloop As Byte
          Dim UI As Integer
          Dim NPCI As Integer
          Dim atacoPJ As Boolean
          Dim UserProtected As Boolean
          
20        atacoPJ = False
          
30        With Npclist(NpcIndex)
40            For headingloop = eHeading.NORTH To eHeading.WEST
50                nPos = .Pos
60                If .flags.Inmovilizado = 0 Or .Char.Heading = headingloop Then
70                    Call HeadtoPos(headingloop, nPos)
80                    If InMapBounds(nPos.map, nPos.X, nPos.Y) Then
90                        UI = MapData(nPos.map, nPos.X, nPos.Y).UserIndex
100                       NPCI = MapData(nPos.map, nPos.X, nPos.Y).NpcIndex
110                       If UI > 0 And Not atacoPJ Then
120                           UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
130                           UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.EnConsulta
                              
140                           If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And (Not UserProtected) Then
                                  
150                               atacoPJ = True
160                               If .Movement = NpcObjeto Then
                                      ' Los npc objeto no atacan siempre al mismo usuario
170                                   If RandomNumber(1, 3) = 3 Then atacoPJ = False
180                               End If
                                  
190                               If atacoPJ Then
200                                   If .flags.LanzaSpells Then
210                                       If .flags.AtacaDoble Then
220                                           If (RandomNumber(0, 1)) Then
230                                               If NpcAtacaUser(NpcIndex, UI) Then
240                                                   Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
250                                               End If
260                                               Exit Sub
270                                           End If
280                                       End If
                                          
290                                       Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
300                                       Call NpcLanzaUnSpell(NpcIndex, UI)
310                                   End If
320                               End If
330                               If NpcAtacaUser(NpcIndex, UI) Then
340                                   Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
350                               End If
360                               Exit Sub

370                           End If
380                       ElseIf NPCI > 0 Then
390                           If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
400                               Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
410                               Call SistemaCombate.NpcAtacaNpc(NpcIndex, NPCI, False)
420                               Exit Sub
430                           End If
440                       End If
450                   End If
460               End If  'inmo
470           Next headingloop
480       End With
          
490       Call RestoreOldMovement(NpcIndex)


    
500       On Error GoTo 0
510       Exit Sub

HostilMalvadoAI_Error:

520       LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HostilMalvadoAI, line " & Erl & "." & " Npc: " & NpcIndex & " , NPCI:" & NPCI & ", MaestroUser: " & Npclist(NPCI).MaestroUser & ": " & UserList(Npclist(NPCI).MaestroUser).Name

End Sub

Private Sub HostilBuenoAI(ByVal NpcIndex As Integer)
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 12/01/2010 (ZaMa)
      '14/09/2009: ZaMa - Now npcs don't atack protected users.
      '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
      '***************************************************
          Dim nPos As WorldPos
          Dim headingloop As eHeading
          Dim UI As Integer
          Dim UserProtected As Boolean
          
   On Error GoTo HostilBuenoAI_Error

10        With Npclist(NpcIndex)
20            For headingloop = eHeading.NORTH To eHeading.WEST
30                nPos = .Pos
40                If .flags.Inmovilizado = 0 Or .Char.Heading = headingloop Then
50                    Call HeadtoPos(headingloop, nPos)
60                    If InMapBounds(nPos.map, nPos.X, nPos.Y) Then
70                        UI = MapData(nPos.map, nPos.X, nPos.Y).UserIndex
80                        If UI > 0 Then
90                            If UserList(UI).Name = .flags.AttackedBy Then
                              
100                               UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
110                               UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.EnConsulta
                                  
120                               If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
130                                   If .flags.LanzaSpells > 0 Then
140                                       Call NpcLanzaUnSpell(NpcIndex, UI)
150                                   End If
                                      
160                                   If NpcAtacaUser(NpcIndex, UI) Then
170                                       Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
180                                   End If
190                                   Exit Sub
200                               End If
210                           End If
220                       End If
230                   End If
240               End If
250           Next headingloop
260       End With
          
270       Call RestoreOldMovement(NpcIndex)

   On Error GoTo 0
   Exit Sub

HostilBuenoAI_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure HostilBuenoAI of Módulo AI in line " & Erl
End Sub

Private Sub IrUsuarioCercano(ByVal NpcIndex As Integer)
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 12/01/2010 (ZaMa)
      '14/09/2009: ZaMa - Now npcs don't follow protected users.
      '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
      '***************************************************
          Dim tHeading As Byte
          Dim UserIndex As Integer
          Dim SignoNS As Integer
          Dim SignoEO As Integer
          Dim i As Long
          Dim UserProtected As Boolean
          
   On Error GoTo IrUsuarioCercano_Error

10        With Npclist(NpcIndex)
20            If .flags.Inmovilizado = 1 Then
30                Select Case .Char.Heading
                      Case eHeading.NORTH
40                        SignoNS = -1
50                        SignoEO = 0
                      
60                    Case eHeading.EAST
70                        SignoNS = 0
80                        SignoEO = 1
                      
90                    Case eHeading.SOUTH
100                       SignoNS = 1
110                       SignoEO = 0
                      
120                   Case eHeading.WEST
130                       SignoEO = -1
140                       SignoNS = 0
150               End Select
                  
160               For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
170                   UserIndex = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                      
                      'Is it in it's range of vision??
180                   If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_X And Sgn(UserList(UserIndex).Pos.X - .Pos.X) = SignoEO Then
190                       If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UserIndex).Pos.Y - .Pos.Y) = SignoNS Then
                              
200                           UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
210                           UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado Or UserList(UserIndex).flags.EnConsulta
                              
220                           If UserList(UserIndex).flags.Muerto = 0 Then
230                               If Not UserProtected Then
240                                   If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UserIndex)
250                                   Exit Sub
260                               End If
270                           End If
                              
280                       End If
290                   End If
300               Next i
                  
              ' No esta inmobilizado
310           Else
                  
                  ' Tiene prioridad de seguir al usuario al que le pertenece si esta en el rango de vision
                  Dim OwnerIndex As Integer
                  
320               OwnerIndex = .Owner
330               If OwnerIndex > 0 Then
                  
                      'Is it in it's range of vision??
340                   If Abs(UserList(OwnerIndex).Pos.X - .Pos.X) <= RANGO_VISION_X Then
350                       If Abs(UserList(OwnerIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                              
                              ' va hacia el si o esta invi ni oculto
360                           If UserList(OwnerIndex).flags.invisible = 0 And UserList(OwnerIndex).flags.Oculto = 0 And Not UserList(OwnerIndex).flags.EnConsulta And Not UserList(OwnerIndex).flags.Ignorado Then
370                               If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, OwnerIndex)
                                      
380                               tHeading = FindDirection(.Pos, UserList(OwnerIndex).Pos)
390                               Call MoveNPCChar(NpcIndex, tHeading)
400                               Exit Sub
410                           End If
420                       End If
430                   End If
                      
440               End If
                  
                  ' No le pertenece a nadie o el dueño no esta en el rango de vision, sigue a cualquiera
450               For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
460                   UserIndex = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                      
                      'Is it in it's range of vision??
470                   If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_X Then
480                       If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                              
490                           With UserList(UserIndex)
                                  
500                               UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And .flags.NoPuedeSerAtacado
510                               UserProtected = UserProtected Or .flags.Ignorado Or .flags.EnConsulta
                                  
520                               If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And _
                                      .flags.AdminPerseguible And Not UserProtected Then
                                      
530                                   If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UserIndex)
                                      
540                                   tHeading = FindDirection(Npclist(NpcIndex).Pos, .Pos)
550                                   Call MoveNPCChar(NpcIndex, tHeading)
560                                   Exit Sub
570                               End If
                                  
580                           End With
                              
590                       End If
600                   End If
610               Next i
                  
                  'Si llega aca es que no habï¿½a ningï¿½n usuario cercano vivo.
                  'A bailar. Pablo (ToxicWaste)
620               If RandomNumber(0, 10) = 0 Then
630                   Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
640               End If
                  
650           End If
660       End With
          
670       Call RestoreOldMovement(NpcIndex)

   On Error GoTo 0
   Exit Sub

IrUsuarioCercano_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure IrUsuarioCercano of Módulo AI in line " & Erl
End Sub

''
' Makes a Pet / Summoned Npc to Follow an enemy
'
' @param NpcIndex Specifies reference to the npc
Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
      '**************************************************************
      'Author: Unknown
      'Last Modify by: Marco Vanotti (MarKoxX)
      'Last Modify Date: 08/16/2008
      '08/16/2008: MarKoxX - Now pets that do melï¿½ attacks have to be near the enemy to attack.
      '**************************************************************
          Dim tHeading As Byte
          Dim UI As Integer
          
          Dim i As Long
          
          Dim SignoNS As Integer
          Dim SignoEO As Integer

   On Error GoTo SeguirAgresor_Error

10        With Npclist(NpcIndex)
20            If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
30                Select Case .Char.Heading
                      Case eHeading.NORTH
40                        SignoNS = -1
50                        SignoEO = 0
                      
60                    Case eHeading.EAST
70                        SignoNS = 0
80                        SignoEO = 1
                      
90                    Case eHeading.SOUTH
100                       SignoNS = 1
110                       SignoEO = 0
                      
120                   Case eHeading.WEST
130                       SignoEO = -1
140                       SignoNS = 0
150               End Select

160               For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
170                   UI = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)

                      'Is it in it's range of vision??
180                   If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X And Sgn(UserList(UI).Pos.X - .Pos.X) = SignoEO Then
190                       If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then

200                           If UserList(UI).Name = .flags.AttackedBy Then
210                               If .MaestroUser > 0 Then
220                                   If Not criminal(.MaestroUser) And Not criminal(UI) And (UserList(.MaestroUser).flags.Seguro Or UserList(.MaestroUser).Faccion.ArmadaReal = 1) Then
230                                       Call WriteShortMsj(.MaestroUser, 27, FontTypeNames.FONTTYPE_INFO)
240                                       Call FlushBuffer(.MaestroUser)
250                                       .flags.AttackedBy = vbNullString
260                                       Exit Sub
270                                   End If
280                               End If

290                               If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
300                                    If .flags.LanzaSpells > 0 Then
310                                         Call NpcLanzaUnSpell(NpcIndex, UI)
320                                    Else
330                                       If Distancia(UserList(UI).Pos, Npclist(NpcIndex).Pos) <= 1 Then
                                              ' TODO : Set this a separate AI for Elementals and Druid's pets
340                                           If Npclist(NpcIndex).Numero <> 92 Then
350                                               Call NpcAtacaUser(NpcIndex, UI)
360                                           End If
370                                       End If
380                                    End If
390                                    Exit Sub
400                               End If
410                           End If
                              
420                       End If
430                   End If
                      
440               Next i
450           Else
460               For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
470                   UI = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                      
                      'Is it in it's range of vision??
480                   If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
490                       If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                              
500                           If UserList(UI).Name = .flags.AttackedBy Then
510                               If .MaestroUser > 0 Then
520                                   If Not criminal(.MaestroUser) And Not criminal(UI) And (UserList(.MaestroUser).flags.Seguro Or UserList(.MaestroUser).Faccion.ArmadaReal = 1) Then
                                          'Call WriteConsoleMsg(.MaestroUser, "La mascota no atacará a ciudadanos si eres miembro del ejército real o tienes el seguro activado.", FontTypeNames.FONTTYPE_INFO)
530                                       Call WriteShortMsj(.MaestroUser, 27, FontTypeNames.FONTTYPE_INFO)
540                                       Call FlushBuffer(.MaestroUser)
550                                       .flags.AttackedBy = vbNullString
560                                       Call FollowAmo(NpcIndex)
570                                       Exit Sub
580                                   End If
590                               End If
                                  
600                               If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
610                                    If .flags.LanzaSpells > 0 Then
620                                           Call NpcLanzaUnSpell(NpcIndex, UI)
630                                    Else
640                                       If Distancia(UserList(UI).Pos, Npclist(NpcIndex).Pos) <= 1 Then
                                              ' TODO : Set this a separate AI for Elementals and Druid's pets
650                                           If Npclist(NpcIndex).Numero <> 92 Then
660                                               Call NpcAtacaUser(NpcIndex, UI)
670                                           End If
680                                       End If
690                                    End If
                                       
700                                    tHeading = FindDirection(.Pos, UserList(UI).Pos)
710                                    Call MoveNPCChar(NpcIndex, tHeading)
                                       
720                                    Exit Sub
730                               End If
740                           End If
                              
750                       End If
760                   End If
                      
770               Next i
780           End If
790       End With
          
800       Call RestoreOldMovement(NpcIndex)

   On Error GoTo 0
   Exit Sub

SeguirAgresor_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure SeguirAgresor of Módulo AI in line " & Erl
End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)
10        With Npclist(NpcIndex)
20            If .MaestroUser = 0 Then
30                .Movement = .flags.OldMovement
40                .Hostile = .flags.OldHostil
50                .flags.AttackedBy = vbNullString
60            End If
70        End With
End Sub

Private Sub PersigueCiudadano(ByVal NpcIndex As Integer)
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 12/01/2010 (ZaMa)
      '14/09/2009: ZaMa - Now npcs don't follow protected users.
      '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs.
      '***************************************************
          Dim UserIndex As Integer
          Dim tHeading As Byte
          Dim i As Long
          Dim UserProtected As Boolean
          
   On Error GoTo PersigueCiudadano_Error

10        With Npclist(NpcIndex)
20            For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
30                UserIndex = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                      
                  'Is it in it's range of vision??
40                If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_X Then
50                    If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                          
60                        If Not criminal(UserIndex) Then
                          
70                            UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
80                            UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado Or UserList(UserIndex).flags.EnConsulta
                              
90                            If UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.invisible = 0 And _
                                  UserList(UserIndex).flags.Oculto = 0 And UserList(UserIndex).flags.AdminPerseguible And Not UserProtected Then
                                  
100                               If .flags.LanzaSpells > 0 Then
110                                   Call NpcLanzaUnSpell(NpcIndex, UserIndex)
120                               End If
130                               tHeading = FindDirection(.Pos, UserList(UserIndex).Pos)
140                               Call MoveNPCChar(NpcIndex, tHeading)
150                               Exit Sub
160                           End If
170                       End If
                          
180                  End If
190               End If
                  
200           Next i
210       End With
          
220       Call RestoreOldMovement(NpcIndex)

   On Error GoTo 0
   Exit Sub

PersigueCiudadano_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PersigueCiudadano of Módulo AI in line " & Erl
End Sub

Private Sub PersigueCriminal(ByVal NpcIndex As Integer)
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 12/01/2010 (ZaMa)
      '14/09/2009: ZaMa - Now npcs don't follow protected users.
      '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs.
      '***************************************************
          Dim UserIndex As Integer
          Dim tHeading As Byte
          Dim i As Long
          Dim SignoNS As Integer
          Dim SignoEO As Integer
          Dim UserProtected As Boolean
          
   On Error GoTo PersigueCriminal_Error

10        With Npclist(NpcIndex)
20            If .flags.Inmovilizado = 1 Then
30                Select Case .Char.Heading
                      Case eHeading.NORTH
40                        SignoNS = -1
50                        SignoEO = 0
                      
60                    Case eHeading.EAST
70                        SignoNS = 0
80                        SignoEO = 1
                      
90                    Case eHeading.SOUTH
100                       SignoNS = 1
110                       SignoEO = 0
                      
120                   Case eHeading.WEST
130                       SignoEO = -1
140                       SignoNS = 0
150               End Select
                  
160               For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
170                   UserIndex = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                      
                      'Is it in it's range of vision??
180                   If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_X And Sgn(UserList(UserIndex).Pos.X - .Pos.X) = SignoEO Then
190                       If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UserIndex).Pos.Y - .Pos.Y) = SignoNS Then
                              
200                           If criminal(UserIndex) Then
210                               With UserList(UserIndex)
                                       
220                                    UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And .flags.NoPuedeSerAtacado
230                                    UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado Or UserList(UserIndex).flags.EnConsulta
                                       
240                                    If .flags.Muerto = 0 And .flags.invisible = 0 And _
                                          .flags.Oculto = 0 And .flags.AdminPerseguible And Not UserProtected Then
                                           
250                                        If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
260                                              Call NpcLanzaUnSpell(NpcIndex, UserIndex)
270                                        End If
280                                        Exit Sub
290                                   End If
300                               End With
310                           End If
                              
320                      End If
330                   End If
340               Next i
350           Else
360               For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
370                   UserIndex = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                      
                      'Is it in it's range of vision??
380                   If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_X Then
390                       If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                              
400                           If criminal(UserIndex) Then
                                  
410                               UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
420                               UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado
                                  
430                               If UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.invisible = 0 And _
                                     UserList(UserIndex).flags.Oculto = 0 And UserList(UserIndex).flags.AdminPerseguible And Not UserProtected Then
440                                   If .flags.LanzaSpells > 0 Then
450                                       Call NpcLanzaUnSpell(NpcIndex, UserIndex)
460                                   End If
470                                   If .flags.Inmovilizado = 1 Then Exit Sub
480                                   tHeading = FindDirection(.Pos, UserList(UserIndex).Pos)
490                                   Call MoveNPCChar(NpcIndex, tHeading)
500                                   Exit Sub
510                              End If
520                           End If
                              
530                      End If
540                   End If
                      
550               Next i
560           End If
570       End With
          
580       Call RestoreOldMovement(NpcIndex)

   On Error GoTo 0
   Exit Sub

PersigueCriminal_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PersigueCriminal of Módulo AI in line " & Erl
End Sub

Private Sub SeguirAmo(ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim tHeading As Byte
          Dim UI As Integer
          
   On Error GoTo SeguirAmo_Error

10        With Npclist(NpcIndex)
20            If .Target = 0 And .TargetNPC = 0 Then
30                UI = .MaestroUser
                  
40                If UI > 0 Then
                      'Is it in it's range of vision??
50                    If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
60                        If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
70                            If UserList(UI).flags.Muerto = 0 _
                                      And UserList(UI).flags.invisible = 0 _
                                      And UserList(UI).flags.Oculto = 0 _
                                      And Distancia(.Pos, UserList(UI).Pos) > 3 Then
80                                tHeading = FindDirection(.Pos, UserList(UI).Pos)
90                                Call MoveNPCChar(NpcIndex, tHeading)
100                               Exit Sub
110                           End If
120                       End If
130                   End If
140               End If
150           End If
160       End With
          
170       Call RestoreOldMovement(NpcIndex)

   On Error GoTo 0
   Exit Sub

SeguirAmo_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure SeguirAmo of Módulo AI in line " & Erl
End Sub

Private Sub AiNpcAtacaNpc(ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim tHeading As Byte
          Dim X As Long
          Dim Y As Long
          Dim NI As Integer
          Dim bNoEsta As Boolean
          
          Dim SignoNS As Integer
          Dim SignoEO As Integer
          
   On Error GoTo AiNpcAtacaNpc_Error

10        With Npclist(NpcIndex)
20            If .flags.Inmovilizado = 1 Then
30                Select Case .Char.Heading
                      Case eHeading.NORTH
40                        SignoNS = -1
50                        SignoEO = 0
                      
60                    Case eHeading.EAST
70                        SignoNS = 0
80                        SignoEO = 1
                      
90                    Case eHeading.SOUTH
100                       SignoNS = 1
110                       SignoEO = 0
                      
120                   Case eHeading.WEST
130                       SignoEO = -1
140                       SignoNS = 0
150               End Select
                  
160               For Y = .Pos.Y To .Pos.Y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
170                   For X = .Pos.X To .Pos.X + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)
180                       If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
190                           NI = MapData(.Pos.map, X, Y).NpcIndex
200                           If NI > 0 Then
210                               If .TargetNPC = NI Then
220                                   bNoEsta = True
230                                   If .Numero = ELEMENTALFUEGO Then
240                                       Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
250                                       If Npclist(NI).NPCtype = Dragon Then
260                                           Npclist(NI).CanAttack = 1
270                                           Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
280                                        End If
290                                    Else
                                          'aca verificamosss la distancia de ataque
300                                       If Distancia(.Pos, Npclist(NI).Pos) <= 2 Then
310                                           Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)
320                                       End If
330                                    End If
340                                    Exit Sub
350                               End If
360                          End If
370                       End If
380                   Next X
390               Next Y
400           Else
410               For Y = .Pos.Y - RANGO_VISION_Y To .Pos.Y + RANGO_VISION_Y
420                   For X = .Pos.X - RANGO_VISION_Y To .Pos.X + RANGO_VISION_Y
430                       If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
440                          NI = MapData(.Pos.map, X, Y).NpcIndex
450                          If NI > 0 Then
460                               If .TargetNPC = NI Then
470                                    bNoEsta = True
480                                    If .Numero = ELEMENTALFUEGO Then
490                                        Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
500                                        If Npclist(NI).NPCtype = Dragon Then
510                                           Npclist(NI).CanAttack = 1
520                                           Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
530                                        End If
540                                    Else
                                          'aca verificamosss la distancia de ataque
550                                       If Distancia(.Pos, Npclist(NI).Pos) <= 3 Then
560                                           Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)
570                                       End If
580                                    End If
590                                    If .flags.Inmovilizado = 1 Then Exit Sub
600                                    If .TargetNPC = 0 Then Exit Sub
610                                    tHeading = FindDirection(.Pos, Npclist(MapData(.Pos.map, X, Y).NpcIndex).Pos)
620                                    Call MoveNPCChar(NpcIndex, tHeading)
630                                    Exit Sub
640                               End If
650                          End If
660                       End If
670                   Next X
680               Next Y
690           End If
              
700           If Not bNoEsta Then
710               If .MaestroUser > 0 Then
720                   Call FollowAmo(NpcIndex)
730               Else
740                   .Movement = .flags.OldMovement
750                   .Hostile = .flags.OldHostil
760               End If
770           End If
780       End With

   On Error GoTo 0
   Exit Sub

AiNpcAtacaNpc_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure AiNpcAtacaNpc of Módulo AI in line " & Erl
End Sub

Public Sub AiNpcObjeto(ByVal NpcIndex As Integer)
      '***************************************************
      'Autor: ZaMa
      'Last Modification: 14/09/2009 (ZaMa)
      '14/09/2009: ZaMa - Now npcs don't follow protected users.
      '***************************************************
          Dim UserIndex As Integer
          Dim tHeading As Byte
          Dim i As Long
          Dim SignoNS As Integer
          Dim SignoEO As Integer
          Dim UserProtected As Boolean
          
   On Error GoTo AiNpcObjeto_Error

10        With Npclist(NpcIndex)
20            For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
30                UserIndex = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                  
                  'Is it in it's range of vision??
40                If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_X Then
50                    If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                          
60                        With UserList(UserIndex)
70                            UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And .flags.NoPuedeSerAtacado
                              
80                            If .flags.Muerto = 0 And .flags.invisible = 0 And _
                                  .flags.Oculto = 0 And .flags.AdminPerseguible And Not UserProtected Then
                                  
                                  ' No quiero que ataque siempre al primero
90                                If RandomNumber(1, 3) < 3 Then
100                                   If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
110                                        Call NpcLanzaUnSpell(NpcIndex, UserIndex)
120                                   End If
                                  
130                                   Exit Sub
140                               End If
150                           End If
160                       End With
170                  End If
180               End If
                  
190           Next i
200       End With

   On Error GoTo 0
   Exit Sub

AiNpcObjeto_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure AiNpcObjeto of Módulo AI in line " & Erl

End Sub
Sub NPCAI(ByVal NpcIndex As Integer)
10        On Error GoTo NPCAI_Error
      '**************************************************************
      'Author: Unknown
      'Last Modify by: ZaMa
      'Last Modify Date: 15/11/2009
      '08/16/2008: MarKoxX - Now pets that do melï¿½ attacks have to be near the enemy to attack.
      '15/11/2009: ZaMa - Implementacion de npc objetos ai.
      '**************************************************************
20        With Npclist(NpcIndex)
              '<<<<<<<<<<< Ataques >>>>>>>>>>>>>>>>
30            If .MaestroUser = 0 Then
                  'Busca a alguien para atacar
                  'ï¿½Es un guardia?
40                If .NPCtype = eNPCType.GuardiaReal Then
50                    Call GuardiasAI(NpcIndex, False)
60                ElseIf .NPCtype = eNPCType.Guardiascaos Then
70                    Call GuardiasAI(NpcIndex, True)
80                ElseIf .Hostile And .Stats.Alineacion <> 0 Then
90                    Call HostilMalvadoAI(NpcIndex)
100               ElseIf .Hostile And .Stats.Alineacion = 0 Then
110                   Call HostilBuenoAI(NpcIndex)
120               End If
130           Else
                  'Evitamos que ataque a su amo, a menos
                  'que el amo lo ataque.
                  'Call HostilBuenoAI(NpcIndex)
140           End If


              '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
150           Select Case .Movement
              Case TipoAI.MueveAlAzar
160               If .flags.Inmovilizado = 1 Then Exit Sub
170               If .NPCtype = eNPCType.GuardiaReal Then
180                   If RandomNumber(1, 12) = 3 Then
190                       Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
200                   End If

210                   Call PersigueCriminal(NpcIndex)

220               ElseIf .NPCtype = eNPCType.Guardiascaos Then
230                   If RandomNumber(1, 12) = 3 Then
240                       Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
250                   End If

260                   Call PersigueCiudadano(NpcIndex)

270               Else
280                   If RandomNumber(1, 12) = 3 Then
290                       Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
300                   End If
310               End If

                  'Va hacia el usuario cercano
320           Case TipoAI.NpcMaloAtacaUsersBuenos
330               Call IrUsuarioCercano(NpcIndex)

                  'Va hacia el usuario que lo ataco(FOLLOW)
340           Case TipoAI.NPCDEFENSA
350               Call SeguirAgresor(NpcIndex)

                  'Persigue criminales
360           Case TipoAI.GuardiasAtacanCriminales
370               Call PersigueCriminal(NpcIndex)

380           Case TipoAI.SigueAmo
390               If .flags.Inmovilizado = 1 Then Exit Sub
400               Call SeguirAmo(NpcIndex)
410               If RandomNumber(1, 12) = 3 Then
420                   Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
430               End If

440           Case TipoAI.NpcAtacaNpc
450               Call AiNpcAtacaNpc(NpcIndex)

460           Case TipoAI.NpcObjeto
470               Call AiNpcObjeto(NpcIndex)

480           Case TipoAI.NpcPathfinding
490               If .flags.Inmovilizado = 1 Then Exit Sub
500               If ReCalculatePath(NpcIndex) Then
510                   Call PathFindingAI(NpcIndex)
                      'Existe el camino?
520                   If .PFINFO.NoPath Then    'Si no existe nos movemos al azar
                          'Move randomly
530                       Call MoveNPCChar(NpcIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))
540                   End If
550               Else
560                   If Not PathEnd(NpcIndex) Then
570                       Call FollowPath(NpcIndex)
580                   Else
590                       .PFINFO.PathLenght = 0
600                   End If
610               End If
            
620         Case TipoAI.NpcDagaRusa
              '  If Events(Npclist(NpcIndex).flags.SlotEvent).TimeCount > 0 Then Exit Sub
                'Call DagaRusa_MoveNpc(NpcIndex)
                
630           End Select
640       End With
650       Exit Sub

NPCAI_Error:

660       LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure NPCAI, line " & Erl & "."
    
    Dim MiNPC  As Npc
670       MiNPC = Npclist(NpcIndex)
680       Call QuitarNPC(NpcIndex)
690       Call ReSpawnNpc(MiNPC)
    
End Sub



Function UserNear(ByVal NpcIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      'Returns True if there is an user adjacent to the npc position.
      '***************************************************

10        With Npclist(NpcIndex)
20            UserNear = Not Int(Distance(.Pos.X, .Pos.Y, UserList(.PFINFO.TargetUser).Pos.X, _
                          UserList(.PFINFO.TargetUser).Pos.Y)) > 1
30        End With
End Function

Function ReCalculatePath(ByVal NpcIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      'Returns true if we have to seek a new path
      '***************************************************

10        If Npclist(NpcIndex).PFINFO.PathLenght = 0 Then
20            ReCalculatePath = True
30        ElseIf Not UserNear(NpcIndex) And Npclist(NpcIndex).PFINFO.PathLenght = Npclist(NpcIndex).PFINFO.CurPos - 1 Then
40            ReCalculatePath = True
50        End If
End Function

Function PathEnd(ByVal NpcIndex As Integer) As Boolean
      '***************************************************
      'Author: Gulfas Morgolock
      'Last Modification: -
      'Returns if the npc has arrived to the end of its path
      '***************************************************
10        PathEnd = Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.PathLenght
End Function

Function FollowPath(ByVal NpcIndex As Integer) As Boolean
      '***************************************************
      'Author: Gulfas Morgolock
      'Last Modification: -
      'Moves the npc.
      '***************************************************
          Dim tmpPos As WorldPos
          Dim tHeading As Byte
          
10        With Npclist(NpcIndex)
20            tmpPos.map = .Pos.map
30            tmpPos.X = .PFINFO.Path(.PFINFO.CurPos).Y ' invertï¿½ las coordenadas
40            tmpPos.Y = .PFINFO.Path(.PFINFO.CurPos).X
              
              'Debug.Print "(" & tmpPos.X & "," & tmpPos.Y & ")"
              
50            tHeading = FindDirection(.Pos, tmpPos)
              
60            MoveNPCChar NpcIndex, tHeading
              
70            .PFINFO.CurPos = .PFINFO.CurPos + 1
80        End With
End Function

Function PathFindingAI(ByVal NpcIndex As Integer) As Boolean
      '***************************************************
      'Author: Gulfas Morgolock
      'Last Modification: -
      'This function seeks the shortest path from the Npc
      'to the user's location.
      '***************************************************
          Dim Y As Long
          Dim X As Long
          
   On Error GoTo PathFindingAI_Error

10        With Npclist(NpcIndex)
20            For Y = .Pos.Y - 10 To .Pos.Y + 10    'Makes a loop that looks at
30                 For X = .Pos.X - 10 To .Pos.X + 10   '5 tiles in every direction
                      
                       'Make sure tile is legal
40                     If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
                          
                           'look for a user
50                         If MapData(.Pos.map, X, Y).UserIndex > 0 Then
                               'Move towards user
                                Dim tmpUserIndex As Integer
60                              tmpUserIndex = MapData(.Pos.map, X, Y).UserIndex
70                              With UserList(tmpUserIndex)
80                                If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible Then
                                      'We have to invert the coordinates, this is because
                                      'ORE refers to maps in converse way of my pathfinding
                                      'routines.
90                                    Npclist(NpcIndex).PFINFO.Target.X = .Pos.Y
100                                   Npclist(NpcIndex).PFINFO.Target.Y = .Pos.X 'ops!
110                                   Npclist(NpcIndex).PFINFO.TargetUser = tmpUserIndex
120                                   Call SeekPath(NpcIndex)
130                                   Exit Function
140                               End If
150                           End With
160                       End If
170                   End If
180               Next X
190           Next Y
200       End With

   On Error GoTo 0
   Exit Function

PathFindingAI_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PathFindingAI of Módulo AI in line " & Erl
End Function

Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
      '**************************************************************
      'Author: Unknown
      'Last Modify by: -
      'Last Modify Date: -
      '**************************************************************
10        With UserList(UserIndex)
20            If .flags.invisible = 1 Or .flags.Oculto = 1 Then Exit Sub
30        End With
          
          Dim k As Integer
40        k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
50        Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, Npclist(NpcIndex).Spells(k))
End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim k As Integer
10        k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
20        Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, Npclist(NpcIndex).Spells(k))
End Sub


Public Sub DagaRusa_MoveNpc(ByVal NpcIndex As Integer, Optional ByVal Init As Boolean = False)

          Dim UserIndex As Integer
          Dim Npc As Npc
          Dim LoopC As Integer
          Dim SlotEvent As Integer
          Dim tHeading As eHeading
          Dim Pos As WorldPos
          
          Static Pasaron As Byte
          
   On Error GoTo DagaRusa_MoveNpc_Error

10        Npc = Npclist(NpcIndex)
20        SlotEvent = Npc.flags.SlotEvent
          
30        If Init Then
40            Pasaron = 0
50            Exit Sub
60        End If
        
70        With Events(SlotEvent)
              
              ' El NPC completa la ronda.
80            If Pasaron >= Npclist(NpcIndex).flags.InscribedPrevio Then
90                DagaRusa_ResetRonda SlotEvent
100               UserIndex = DagaRusa_NextUser(SlotEvent)
                  
110               Pos.map = Npclist(NpcIndex).Pos.map
120               Pos.X = UserList(UserIndex).Pos.X
130               Pos.Y = UserList(UserIndex).Pos.Y - 1
140               tHeading = FindDirection(Npclist(NpcIndex).Pos, Pos)
150               Call MoveNPCChar(NpcIndex, tHeading)
                  
160               If Npclist(NpcIndex).Pos.X = Pos.X Then
170                   Pasaron = 0
180                   Npclist(NpcIndex).flags.InscribedPrevio = .Inscribed
190               End If
                  
                  
200               Exit Sub
210           End If
                      
220           UserIndex = DagaRusa_NextUser(SlotEvent)
              
230           If UserIndex > 0 Then
              
240               If Not (Distancia(Npclist(NpcIndex).Pos, UserList(UserIndex).Pos) <= 1) Then
250                   Pos.map = UserList(UserIndex).Pos.map
260                   Pos.X = UserList(UserIndex).Pos.X
270                   Pos.Y = UserList(UserIndex).Pos.Y - 1
                                  
280                   tHeading = FindDirection(Npclist(NpcIndex).Pos, Pos)
290                   Call MoveNPCChar(NpcIndex, tHeading)
300                   Call ChangeNPCChar(NpcIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.Head, tHeading)
310               Else
320                   Call ChangeNPCChar(NpcIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.Head, eHeading.SOUTH)
330               End If
                  
340               If (Distancia(Npclist(NpcIndex).Pos, UserList(UserIndex).Pos) <= 1) Then
350                   Call ChangeNPCChar(NpcIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.Head, SOUTH)
360                   .Users(UserList(UserIndex).flags.SlotUserEvent).value = 1
370                   Call DagaRusa_AttackUser(UserIndex, NpcIndex)
                      'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(Userindex).Name, FontTypeNames.FONTTYPE_INFO)
380                   Npclist(NpcIndex).Target = UserIndex
390                   Pasaron = Pasaron + 1
400               End If
                  

                  

                  
                     ' If Npclist(NpcIndex).Target <> Userindex Then
                        '  .Users(UserList(Userindex).flags.SlotUserEvent).value = 1
                        '  Call DagaRusa_AttackUser(Userindex, NpcIndex)
                          
                          
                          'Npclist(NpcIndex).Target = Userindex
                         ' Pasaron = Pasaron + 1
                              
                         ' Exit Sub
                      'End If
                 ' End If
              
410           End If
              
420       End With

   On Error GoTo 0
   Exit Sub

DagaRusa_MoveNpc_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure DagaRusa_MoveNpc of Módulo AI in line " & Erl
End Sub

