Attribute VB_Name = "ModQuests"
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
'along with this program; if not, you can find it at [url=http://www.affero.org/oagpl.html]http://www.affero.org/oagpl.html[/url]
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at [email=aaron@baronsoft.com]aaron@baronsoft.com[/email]
'for more information about ORE please visit [url=http://www.baronsoft.com/]http://www.baronsoft.com/[/url]
Option Explicit
 
'Constantes de las quests
Public Const MAXUSERQUESTS As Integer = 2     'Máxima cantidad de quests que puede tener un usuario al mismo tiempo.
 
Public Function TieneQuest(ByVal Userindex As Integer, ByVal QuestNumber As Integer) As Byte
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Devuelve el slot de UserQuests en que tiene la quest QuestNumber. En caso contrario devuelve 0.
      'Last modified: 27/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      Dim i As Integer
       
10        For i = 1 To MAXUSERQUESTS
20            If UserList(Userindex).QuestStats.Quests(i).QuestIndex = QuestNumber Then
30                TieneQuest = i
40                Exit Function
50            End If
60        Next i
         
70        TieneQuest = 0
End Function
 
Public Function FreeQuestSlot(ByVal Userindex As Integer) As Byte
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Devuelve el próximo slot de quest libre.
      'Last modified: 27/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      Dim i As Integer
       
10        For i = 1 To MAXUSERQUESTS
20            If UserList(Userindex).QuestStats.Quests(i).QuestIndex = 0 Then
30                FreeQuestSlot = i
40                Exit Function
50            End If
60        Next i
         
70        FreeQuestSlot = 0
End Function
 
Public Sub HandleQuestAccept(ByVal Userindex As Integer)
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Maneja el evento de aceptar una quest.
      'Last modified: 31/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      Dim NpcIndex As Integer
      Dim QuestSlot As Byte
       
10        Call UserList(Userindex).incomingData.ReadByte
       
20        NpcIndex = UserList(Userindex).flags.TargetNPC
         
30        If NpcIndex = 0 Then Exit Sub
         
          'Está el personaje en la distancia correcta?
40        If Distancia(UserList(Userindex).Pos, Npclist(NpcIndex).Pos) > 5 Then
50            Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
60            Exit Sub
70        End If
         
80        QuestSlot = FreeQuestSlot(Userindex)
         
          'Agregamos la quest.
90        With UserList(Userindex).QuestStats.Quests(QuestSlot)
100           .QuestIndex = Npclist(NpcIndex).QuestNumber
             
110           If QuestList(.QuestIndex).RequiredNPCs Then ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
120           Call WriteConsoleMsg(Userindex, "Has aceptado la misión " & Chr(34) & QuestList(.QuestIndex).Nombre & Chr(34) & ".", FontTypeNames.FONTTYPE_INFO)
             
130       End With
End Sub
 
Public Sub FinishQuest(ByVal Userindex As Integer, ByVal QuestIndex As Integer, ByVal QuestSlot As Byte)
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Maneja el evento de terminar una quest.
      'Last modified: 29/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      Dim i As Integer
      Dim InvSlotsLibres As Byte
      Dim NpcIndex As Integer
       
10        NpcIndex = UserList(Userindex).flags.TargetNPC
         
20        With QuestList(QuestIndex)
              'Comprobamos que tenga los objetos.
30            If .RequiredOBJs > 0 Then
40                For i = 1 To .RequiredOBJs
50                    If TieneObjetos(.RequiredOBJ(i).objindex, .RequiredOBJ(i).Amount, Userindex) = False Then
60                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No has conseguido todos los objetos que te he pedido.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
70                        Exit Sub
80                    End If
90                Next i
100           End If
             
              'Comprobamos que haya matado todas las criaturas.
110           If .RequiredNPCs > 0 Then
120               For i = 1 To .RequiredNPCs
130                   If .RequiredNPC(i).Amount > UserList(Userindex).QuestStats.Quests(QuestSlot).NPCsKilled(i) Then
140                       Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No has matado todas las criaturas que te he pedido.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
150                       Exit Sub
160                   End If
170               Next i
180           End If
         
              'Comprobamos que el usuario tenga espacio para recibir los items.
190           If .RewardOBJs > 0 Then
                  'Buscamos la cantidad de slots de inventario libres.
200               For i = 1 To MAX_INVENTORY_SLOTS
210                   If UserList(Userindex).Invent.Object(i).objindex = 0 Then InvSlotsLibres = InvSlotsLibres + 1
220               Next i
                 
                  'Nos fijamos si entra
230               If InvSlotsLibres < .RewardOBJs Then
240                   Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No tienes suficiente espacio en el inventario para recibir la recompensa. Vuelve cuando hayas hecho más espacio.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
250                   Exit Sub
260               End If
270           End If
         
              'A esta altura ya cumplió los objetivos, entonces se le entregan las recompensas.
280           Call WriteConsoleMsg(Userindex, "¡Has completado la misión " & Chr(34) & QuestList(QuestIndex).Nombre & Chr(34) & "!", FontTypeNames.FONTTYPE_INFO)
             
              'Si la quest pedía objetos, se los saca al personaje.
290           If .RequiredOBJs Then
300               For i = 1 To .RequiredOBJs
310                   Call QuitarObjetos(.RequiredOBJ(i).objindex, .RequiredOBJ(i).Amount, Userindex)
320               Next i
330           End If
             
              'Se entrega la experiencia.
340           If .RewardEXP Then
350               UserList(Userindex).Stats.Exp = UserList(Userindex).Stats.Exp + .RewardEXP
360               Call WriteConsoleMsg(Userindex, "Has ganado " & .RewardEXP & " puntos de experiencia como recompensa.", FontTypeNames.FONTTYPE_INFO)
370           End If
             
              'Se entrega el oro.
380           If .RewardGLD Then
390               UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld + .RewardGLD
400               Call WriteConsoleMsg(Userindex, "Has ganado " & .RewardGLD & " monedas de oro como recompensa.", FontTypeNames.FONTTYPE_INFO)
410           End If
             
              'Si hay recompensa de objetos, se entregan.
420           If .RewardOBJs > 0 Then
430               For i = 1 To .RewardOBJs
440                   If .RewardOBJ(i).Amount Then
450                       Call MeterItemEnInventario(Userindex, .RewardOBJ(i))
460                       Call WriteConsoleMsg(Userindex, "Has recibido " & QuestList(QuestIndex).RewardOBJ(i).Amount & " " & ObjData(QuestList(QuestIndex).RewardOBJ(i).objindex).Name & " como recompensa.", FontTypeNames.FONTTYPE_INFO)
470                   End If
480               Next i
490           End If
         
              'Actualizamos el personaje
500           Call CheckUserLevel(Userindex)
510           Call UpdateUserInv(True, Userindex, 0)
         
              'Limpiamos el slot de quest.
520           Call CleanQuestSlot(Userindex, QuestSlot)
             
              'Ordenamos las quests
530           Call ArrangeUserQuests(Userindex)
         
              'Se agrega que el usuario ya hizo esta quest.
540           Call AddDoneQuest(Userindex, QuestIndex)
550       End With
End Sub
 
Public Sub AddDoneQuest(ByVal Userindex As Integer, ByVal QuestIndex As Integer)
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Agrega la quest QuestIndex a la lista de quests hechas.
      'Last modified: 28/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
10        With UserList(Userindex).QuestStats
20            .NumQuestsDone = .NumQuestsDone + 1
30            ReDim Preserve .QuestsDone(1 To .NumQuestsDone)
40            .QuestsDone(.NumQuestsDone) = QuestIndex
50        End With
End Sub
 
Public Function UserDoneQuest(ByVal Userindex As Integer, ByVal QuestIndex As Integer) As Boolean
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Verifica si el usuario hizo la quest QuestIndex.
      'Last modified: 28/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      Dim i As Integer
10        With UserList(Userindex).QuestStats
20            If .NumQuestsDone Then
30                For i = 1 To .NumQuestsDone
40                    If .QuestsDone(i) = QuestIndex Then
50                        UserDoneQuest = True
60                        Exit Function
70                    End If
80                Next i
90            End If
100       End With
         
110       UserDoneQuest = False
             
End Function
 
Public Sub CleanQuestSlot(ByVal Userindex As Integer, ByVal QuestSlot As Integer)
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Limpia un slot de quest de un usuario.
      'Last modified: 28/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      Dim i As Integer
       
10        With UserList(Userindex).QuestStats.Quests(QuestSlot)
20            If .QuestIndex Then
30                If QuestList(.QuestIndex).RequiredNPCs Then
40                    For i = 1 To QuestList(.QuestIndex).RequiredNPCs
50                        .NPCsKilled(i) = 0
60                    Next i
70                End If
80            End If
90            .QuestIndex = 0
100       End With
End Sub
 
Public Sub ResetQuestStats(ByVal Userindex As Integer)
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Limpia todos los QuestStats de un usuario
      'Last modified: 28/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      Dim i As Integer
       
10        For i = 1 To MAXUSERQUESTS
20            Call CleanQuestSlot(Userindex, i)
30        Next i
         
40        With UserList(Userindex).QuestStats
50            .NumQuestsDone = 0
60            Erase .QuestsDone
70        End With
End Sub
 
Public Sub HandleQuest(ByVal Userindex As Integer)
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Maneja el paquete Quest.
      'Last modified: 28/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      Dim NpcIndex As Integer
      Dim tmpByte As Byte
       
          'Leemos el paquete
10        Call UserList(Userindex).incomingData.ReadByte
       
20        NpcIndex = UserList(Userindex).flags.TargetNPC
         
30        If NpcIndex = 0 Then Exit Sub
         
          'Está el personaje en la distancia correcta?
40        If Distancia(UserList(Userindex).Pos, Npclist(NpcIndex).Pos) > 5 Then
50            Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
60            Exit Sub
70        End If
         
          'El NPC hace quests?
80        If Npclist(NpcIndex).QuestNumber = 0 Then
90           Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No tengo ninguna misión para ti.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
100           Exit Sub
110       End If
         
          'El personaje ya hizo la quest?
120       If UserDoneQuest(Userindex, Npclist(NpcIndex).QuestNumber) Then
              'Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("Ya has hecho una misión para mi.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
              'Exit Sub
130       End If
       
          'El personaje tiene suficiente nivel?
140       If UserList(Userindex).Stats.ELV < QuestList(Npclist(NpcIndex).QuestNumber).RequiredLevel Then
150           Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("Debes ser por lo menos nivel " & QuestList(Npclist(NpcIndex).QuestNumber).RequiredLevel & " para emprender esta misión.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
160           Exit Sub
170       End If
         
          'A esta altura ya analizo todas las restricciones y esta preparado para el handle propiamente dicho
       
180       tmpByte = TieneQuest(Userindex, Npclist(NpcIndex).QuestNumber)
         
190       If tmpByte Then
              'El usuario está haciendo la quest, entonces va a hablar con el NPC para recibir la recompensa.
200           Call FinishQuest(Userindex, Npclist(NpcIndex).QuestNumber, tmpByte)
210       Else
              'El usuario no está haciendo la quest, entonces primero recibe un informe con los detalles de la misión.
220           tmpByte = FreeQuestSlot(Userindex)
             
              'El personaje tiene algun slot de quest para la nueva quest?
230           If tmpByte = 0 Then
240               Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("Estás haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
250               Exit Sub
260           End If
             
              'Enviamos los detalles de la quest
270           Call WriteQuestDetails(Userindex, Npclist(NpcIndex).QuestNumber)
280       End If
End Sub
 
Public Sub LoadQuests()
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Carga el archivo QUESTS.DAT en el array QuestList.
      'Last modified: 27/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
10    On Error GoTo ErrorHandler
      Dim Reader As clsIniManager
      Dim NumQuests As Integer
      Dim TmpStr As String
      Dim i As Integer
      Dim j As Integer
         
          'Cargamos el clsIniManager en memoria
20        Set Reader = New clsIniManager
         
          'Lo inicializamos para el archivo QUESTS.DAT
30        Call Reader.Initialize(DatPath & "QUESTS.DAT")
         
          'Redimensionamos el array
40        NumQuests = Reader.GetValue("INIT", "NumQuests")
50        ReDim QuestList(1 To NumQuests)
         
          'Cargamos los datos
60        For i = 1 To NumQuests
70            With QuestList(i)
80                .Nombre = Reader.GetValue("QUEST" & i, "Nombre")
90                .desc = Reader.GetValue("QUEST" & i, "Desc")
100               .RequiredLevel = val(Reader.GetValue("QUEST" & i, "RequiredLevel"))
                 
                  'CARGAMOS OBJETOS REQUERIDOS
110               .RequiredOBJs = val(Reader.GetValue("QUEST" & i, "RequiredOBJs"))
120               If .RequiredOBJs > 0 Then
130                   ReDim .RequiredOBJ(1 To .RequiredOBJs)
140                   For j = 1 To .RequiredOBJs
150                       TmpStr = Reader.GetValue("QUEST" & i, "RequiredOBJ" & j)
                         
160                       .RequiredOBJ(j).objindex = val(ReadField(1, TmpStr, 45))
170                       .RequiredOBJ(j).Amount = val(ReadField(2, TmpStr, 45))
180                   Next j
190               End If
                 
                  'CARGAMOS NPCS REQUERIDOS
200               .RequiredNPCs = val(Reader.GetValue("QUEST" & i, "RequiredNPCs"))
210               If .RequiredNPCs > 0 Then
220                   ReDim .RequiredNPC(1 To .RequiredNPCs)
230                   For j = 1 To .RequiredNPCs
240                       TmpStr = Reader.GetValue("QUEST" & i, "RequiredNPC" & j)
                         
250                       .RequiredNPC(j).NpcIndex = val(ReadField(1, TmpStr, 45))
260                       .RequiredNPC(j).Amount = val(ReadField(2, TmpStr, 45))
270                   Next j
280               End If
                 
290               .RewardGLD = val(Reader.GetValue("QUEST" & i, "RewardGLD"))
300               .RewardEXP = val(Reader.GetValue("QUEST" & i, "RewardEXP"))
                 
                  'CARGAMOS OBJETOS DE RECOMPENSA
310               .RewardOBJs = val(Reader.GetValue("QUEST" & i, "RewardOBJs"))
320               If .RewardOBJs > 0 Then
330                   ReDim .RewardOBJ(1 To .RewardOBJs)
340                   For j = 1 To .RewardOBJs
350                       TmpStr = Reader.GetValue("QUEST" & i, "RewardOBJ" & j)
                         
360                       .RewardOBJ(j).objindex = val(ReadField(1, TmpStr, 45))
370                       .RewardOBJ(j).Amount = val(ReadField(2, TmpStr, 45))
380                   Next j
390               End If
400           End With
410       Next i
         
          'Eliminamos la clase
420       Set Reader = Nothing
430   Exit Sub
                         
ErrorHandler:
440       MsgBox "Error cargando el archivo QUESTS.DAT.", vbOKOnly + vbCritical
End Sub
 
Public Sub LoadQuestStats(ByVal Userindex As Integer, ByRef UserFile As clsIniManager)
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Carga las QuestStats del usuario.
      'Last modified: 28/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      Dim i As Integer
      Dim j As Integer
      Dim TmpStr As String
       
10        For i = 1 To MAXUSERQUESTS
20            With UserList(Userindex).QuestStats.Quests(i)
30                TmpStr = UserFile.GetValue("QUESTS", "Q" & i)
                 
40                .QuestIndex = val(ReadField(1, TmpStr, 45))
50                If .QuestIndex Then
60                    If QuestList(.QuestIndex).RequiredNPCs Then
70                        ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
                         
80                        For j = 1 To QuestList(.QuestIndex).RequiredNPCs
90                            .NPCsKilled(j) = val(ReadField(j + 1, TmpStr, 45))
100                       Next j
110                   End If
120               End If
130           End With
140       Next i
         
150       With UserList(Userindex).QuestStats
160           TmpStr = UserFile.GetValue("QUESTS", "QuestsDone")
             
170           .NumQuestsDone = val(ReadField(1, TmpStr, 45))
             
180           If .NumQuestsDone Then
190               ReDim .QuestsDone(1 To .NumQuestsDone)
200               For i = 1 To .NumQuestsDone
210                   .QuestsDone(i) = val(ReadField(i + 1, TmpStr, 45))
220               Next i
230           End If
240       End With
                         
End Sub
 
Public Sub SaveQuestStats(ByVal Userindex As Integer, ByVal Manager As clsIniManager)
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Guarda las QuestStats del usuario.
      'Last modified: 29/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      Dim i As Integer
      Dim j As Integer
      Dim TmpStr As String
       
10        For i = 1 To MAXUSERQUESTS
20            With UserList(Userindex).QuestStats.Quests(i)
30                TmpStr = .QuestIndex
                 
40                If .QuestIndex Then
50                    If QuestList(.QuestIndex).RequiredNPCs Then
60                        For j = 1 To QuestList(.QuestIndex).RequiredNPCs
70                            TmpStr = TmpStr & "-" & .NPCsKilled(j)
80                        Next j
90                    End If
100               End If
             
110               Call Manager.ChangeValue("QUESTS", "Q" & i, TmpStr)
120           End With
130       Next i
         
140       With UserList(Userindex).QuestStats
150           TmpStr = .NumQuestsDone
             
160           If .NumQuestsDone Then
170               For i = 1 To .NumQuestsDone
180                   TmpStr = TmpStr & "-" & .QuestsDone(i)
190               Next i
200           End If
              
              
210           Call Manager.ChangeValue("QUESTS", "QuestsDone", TmpStr)
220       End With
End Sub
 
Public Sub HandleQuestListRequest(ByVal Userindex As Integer)
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Maneja el paquete QuestListRequest.
      'Last modified: 30/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
       
          'Leemos el paquete
10        Call UserList(Userindex).incomingData.ReadByte
         
20        Call WriteQuestListSend(Userindex)
End Sub
 
Public Sub ArrangeUserQuests(ByVal Userindex As Integer)
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Ordena las quests del usuario de manera que queden todas al principio del arreglo.
      'Last modified: 30/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      Dim i As Integer
      Dim j As Integer
       
10        With UserList(Userindex).QuestStats
20            For i = 1 To MAXUSERQUESTS - 1
30                If .Quests(i).QuestIndex = 0 Then
40                    For j = i + 1 To MAXUSERQUESTS
50                        If .Quests(j).QuestIndex Then
60                            .Quests(i) = .Quests(j)
70                            Call CleanQuestSlot(Userindex, j)
80                            Exit For
90                        End If
100                   Next j
110               End If
120           Next i
130       End With
End Sub
 
Public Sub HandleQuestDetailsRequest(ByVal Userindex As Integer)
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Maneja el paquete QuestInfoRequest.
      'Last modified: 30/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      Dim QuestSlot As Byte
          
10        If UserList(Userindex).incomingData.length < 2 Then
20            Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
          'Leemos el paquete
50        Call UserList(Userindex).incomingData.ReadByte
         
60        QuestSlot = UserList(Userindex).incomingData.ReadByte
         
70        Call WriteQuestDetails(Userindex, UserList(Userindex).QuestStats.Quests(QuestSlot).QuestIndex, QuestSlot)
End Sub
 
Public Sub HandleQuestAbandon(ByVal Userindex As Integer)
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Maneja el paquete QuestAbandon.
      'Last modified: 31/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

10        If UserList(Userindex).incomingData.length < 2 Then
20            Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
30            Exit Sub
40        End If
          
          'Leemos el paquete.
50        Call UserList(Userindex).incomingData.ReadByte
         
          'Borramos la quest.
60        Call CleanQuestSlot(Userindex, UserList(Userindex).incomingData.ReadByte)
         
          'Ordenamos la lista de quests del usuario.
70        Call ArrangeUserQuests(Userindex)
         
          'Enviamos la lista de quests actualizada.
80        Call WriteQuestListSend(Userindex)
          
          
          
End Sub

