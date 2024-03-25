Attribute VB_Name = "mQuests"
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
Public Const MAXUSERQUESTS As Integer = 30     'Máxima cantidad de quests que puede tener un usuario al mismo tiempo.

Public NumQuests           As Integer
 
Public Function FreeQuestSlot(ByVal UserIndex As Integer) As Integer

    '<EhHeader>
    On Error GoTo FreeQuestSlot_Err

    '</EhHeader>

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Devuelve el próximo slot de quest libre.
    'Last modified: 27/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer
       
    For i = 1 To MAXUSERQUESTS
            
        If UserList(UserIndex).QuestStats(i).QuestIndex = 0 Then
            FreeQuestSlot = i

            Exit Function

        End If

    Next i
         
    FreeQuestSlot = 0
    '<EhFooter>
    Exit Function

FreeQuestSlot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.FreeQuestSlot " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub Quest_SetUserPrincipa(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo Quest_SetUserPrincipa_Err

    '</EhHeader>
        
    Exit Sub
        
    If UserList(UserIndex).Stats.Elv <= 12 Then
        Call Quest_SetUser(UserIndex, 1) '  Newbie
    Else
        Call Quest_SetUser(UserIndex, 2) '  Aledaño

    End If
    
    Call Quest_SetUser(UserIndex, 10) '  En busca del Oasis :: Misión n°10
    Call Quest_SetUser(UserIndex, 12) '  La oscuridad bajo la Ciudad :: Misión n°12
    Call Quest_SetUser(UserIndex, 17) '  Maldición Marabel :: Misión n°17
    Call Quest_SetUser(UserIndex, 23) '  Tesoro del Dragon Mitico :: Misión n°23
    Call Quest_SetUser(UserIndex, 26) '  Explorando nuevas Islas :: Misión n°26
    Call Quest_SetUser(UserIndex, 32) '  Expedición a la Isla Vespar :: Misión n°32
    Call Quest_SetUser(UserIndex, 37) '  Explorando los Mares Ocultos de Nereo :: Misión n°37
    Call Quest_SetUser(UserIndex, 40) '  Guerreros del Laberinto Spectra :: Misión n°40
    Call Quest_SetUser(UserIndex, 44) '  Explorando los Mares de Nueva Esperanza :: Misión n°44
    Call Quest_SetUser(UserIndex, 51) '  Refugiado en la Isla Veril :: Misión n°51
    Call Quest_SetUser(UserIndex, 58) '  Isla de los Sacerdotes y Protectores del Rey :: Misión n°58
    Call Quest_SetUser(UserIndex, 67) '  Descubriendo el Tenebroso Castillo Brezal :: Misión n°67
    Call Quest_SetUser(UserIndex, 74) '  Afueras del Infierno :: Misión n°74
    Call Quest_SetUser(UserIndex, 84) '  En busca del Polo Norte :: Misión n°84
        
    Call WriteQuestInfo(UserIndex, True, 0)
    Call WriteConsoleMsg(UserIndex, "Misiones> Accede al panel de misiones desde la tecla 'ESC' o bien escribiendo /MISIONES", FontTypeNames.FONTTYPE_CRITICO)
        
    '<EhFooter>
    Exit Sub

Quest_SetUserPrincipa_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.Quest_SetUserPrincipa " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

' Setea una nueva mision/objetivo en el personaje
Public Sub Quest_SetUser(ByVal UserIndex As Integer, ByVal QuestIndex As Integer)

    '<EhHeader>
    On Error GoTo Quest_SetUser_Err

    '</EhHeader>
        
    Dim QuestSlot As Integer
        
    Exit Sub
    QuestSlot = FreeQuestSlot(UserIndex)
        
    If QuestSlot > 0 Then

        With UserList(UserIndex).QuestStats(QuestSlot)
            .QuestIndex = QuestIndex
        
            If QuestList(QuestIndex).RequiredNPCs > 0 Then ReDim .NPCsKilled(1 To QuestList(QuestIndex).RequiredNPCs) As Long
            If QuestList(QuestIndex).RequiredChestOBJs > 0 Then ReDim .ObjsPick(1 To QuestList(QuestIndex).RequiredChestOBJs) As Long
            If QuestList(QuestIndex).RequiredSaleOBJs > 0 Then ReDim .ObjsSale(1 To QuestList(QuestIndex).RequiredSaleOBJs) As Long

        End With
        
    Else
        Call WriteConsoleMsg(UserIndex, "Error al otorgar una nueva misión.", FontTypeNames.FONTTYPE_INFORED)

    End If

    '<EhFooter>
    Exit Sub

Quest_SetUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.Quest_SetUser " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Comprueba si tiene los objetos
Public Sub Quests_Check_Objs(ByVal UserIndex As Integer, _
                             ByVal ObjIndex As Integer, _
                             ByVal Amount As Integer)

    '<EhHeader>
    On Error GoTo Quests_Check_ChestObj_Err

    '</EhHeader>

    Dim A          As Long, B As Long

    Dim QuestIndex As Integer
        
    For B = 1 To MAXUSERQUESTS
        QuestIndex = UserList(UserIndex).QuestStats(B).QuestIndex
        
        If QuestIndex = 0 Then Exit Sub
        
        With QuestList(QuestIndex)

            If .RequiredOBJs > 0 Then

                For A = 1 To .RequiredOBJs

                    If ObjIndex = .RequiredObj(A).ObjIndex Then
                        If TieneObjetos(.RequiredObj(A).ObjIndex, .RequiredObj(A).Amount, UserIndex) Then
                            Call Quests_Final(UserIndex, B)
                            Exit For
                        Else
                            Call WriteQuestInfo(UserIndex, False, B)
                                
                        End If

                    End If

                Next A

            End If
                
        End With
        
    Next B

    '<EhFooter>
    Exit Sub

Quests_Check_ChestObj_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.Quests_Check_ChestObj " & "at line " & Erl

    '</EhFooter>
End Sub

' Le otorga la recompensa que merece por haber completado la misión
Public Function Quests_CheckFinish(ByVal UserIndex As Integer, _
                                   ByVal Slot As Integer) As Boolean

    '<EhHeader>
    On Error GoTo Quests_CheckFinish_Err

    '</EhHeader>

    Dim QuestIndex As Integer
    
    With UserList(UserIndex)
        QuestIndex = .QuestStats(Slot).QuestIndex
        
        Dim A As Long
        
        With QuestList(QuestIndex)

            If .RequiredNPCs > 0 Then

                For A = 1 To .RequiredNPCs

                    If .RequiredNpc(A).Amount * .RequiredNpc(A).Hp <> UserList(UserIndex).QuestStats(Slot).NPCsKilled(A) Then
                        Exit Function

                    End If

                Next A

            End If
            
            If .RequiredSaleOBJs > 0 Then

                For A = 1 To .RequiredSaleOBJs

                    If .RequiredSaleObj(A).Amount <> UserList(UserIndex).QuestStats(Slot).ObjsSale(A) Then
                        Exit Function

                    End If

                Next A

            End If
            
            If .RequiredChestOBJs > 0 Then

                For A = 1 To .RequiredChestOBJs

                    If .RequiredChestObj(A).Amount <> UserList(UserIndex).QuestStats(Slot).ObjsPick(A) Then
                        Exit Function

                    End If

                Next A

            End If
            
            If .RequiredOBJs > 0 Then

                For A = 1 To .RequiredOBJs

                    If Not TieneObjetos(.RequiredObj(A).ObjIndex, .RequiredObj(A).Amount, UserIndex) Then
                        Exit Function

                    End If

                Next A

            End If
            
        End With

    End With
    
    Quests_CheckFinish = True

    '<EhFooter>
    Exit Function

Quests_CheckFinish_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.Quests_CheckFinish " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub Quests_Reward(ByVal UserIndex As Integer, ByVal Slot As Integer)

    '<EhHeader>
    On Error GoTo Quests_Reward_Err

    '</EhHeader>
    
    Dim A          As Long

    Dim QuestIndex As Integer
    
    Dim Text       As String
        
    'Dim List()     As String
        
    Dim Obj        As Obj
    
    '    ReDim Preserve List(0) As String
        
    QuestIndex = UserList(UserIndex).QuestStats(Slot).QuestIndex
    
    With UserList(UserIndex)

        If QuestList(QuestIndex).RequiredOBJs > 0 Then

            If QuestList(QuestIndex).Remove > 0 Then

                For A = 1 To QuestList(QuestIndex).RequiredOBJs
                    Call QuitarObjetos(QuestList(QuestIndex).RequiredObj(A).ObjIndex, QuestList(QuestIndex).RequiredObj(A).Amount, UserIndex)
                Next A

            End If
                
        End If
            
        If QuestList(QuestIndex).RewardEXP > 0 Then
            .Stats.Exp = .Stats.Exp + QuestList(QuestIndex).RewardEXP
            Call CheckUserLevel(UserIndex)
            Call WriteUpdateExp(UserIndex)
                  
            'ReDim Preserve List(0 To UBound(List) + 1) As String
            'List(1) = "+" & QuestList(QuestIndex).RewardEXP & " EXP"

        End If
            
        If QuestList(QuestIndex).RewardEldhir > 0 Then
            .Account.Eldhir = .Account.Eldhir + QuestList(QuestIndex).RewardEldhir
            Call WriteUpdateDsp(UserIndex)
                  
            ' ReDim Preserve List(0 To UBound(List) + 1) As String
            'List(UBound(List)) = "+" & QuestList(QuestIndex).RewardEldhir & " DSP"

        End If
            
        If QuestList(QuestIndex).RewardGLD > 0 Then
            .Stats.Gld = .Stats.Gld + QuestList(QuestIndex).RewardGLD
            Call WriteUpdateGold(UserIndex)
            
            '  ReDim Preserve List(0 To UBound(List) + 1) As String
            ' List(UBuond(List)) = "+" & QuestList(QuestIndex).RewardGLD & " ORO"

        End If
        
        If QuestList(QuestIndex).RewardOBJs > 0 Then

            For A = 1 To QuestList(QuestIndex).RewardOBJs
                    
                Obj.ObjIndex = QuestList(QuestIndex).RewardObj(A).ObjIndex
                Obj.Amount = QuestList(QuestIndex).RewardObj(A).Amount
                      
                If ObjData(Obj.ObjIndex).OBJType = otRangeQuest Then
                    Call UseCofrePoder(UserIndex, ObjData(Obj.ObjIndex).Range)
                Else

                    If ClasePuedeUsarItem(UserIndex, Obj.ObjIndex) Then
                        
                        If Not MeterItemEnInventario(UserIndex, Obj) Then
                            Call TirarItemAlPiso(.Pos, Obj)
    
                        End If

                    End If

                End If
                      
                '  ReDim Preserve List(0 To UBound(List) + 1) As String
                ' List(UBound(List)) = "+" & ObjData(Obj.ObjIndex).Name & " (x" & QuestList(QuestIndex).RewardObj(A).Amount & ")"
            Next A
            
        End If
        
        Call WriteUpdateFinishQuest(UserIndex, QuestIndex)
              
        Call SendData(SendTarget.ToOne, UserIndex, PrepareMessagePlayEffect(RandomNumber(eSound.sVictory3, eSound.sVictory5), UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(UserIndex).Char.charindex))

    End With
    
    '<EhFooter>
    Exit Sub

Quests_Reward_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.Quests_Reward " & "at line " & Erl

    '</EhFooter>
End Sub

' Comprueba cuando golpea una criatura
Public Sub Quests_AddNpc(ByVal UserIndex As Integer, _
                         ByVal NpcIndex As Integer, _
                         ByVal Damage As Long)

    '<EhHeader>
    On Error GoTo Quests_AddNpc_Err

    '</EhHeader>
    
    Dim Diferencia As Long

    Dim A          As Long
        
    Dim B          As Long
        
    Dim TempQuest  As tUserQuest
        
    For B = 1 To MAXUSERQUESTS

        With UserList(UserIndex).QuestStats(B)
                
            If .QuestIndex Then
                    
                TempQuest = UserList(UserIndex).QuestStats(B)
                    
                If Damage > Npclist(NpcIndex).Stats.MinHp Then
                    Diferencia = Abs(Npclist(NpcIndex).Stats.MinHp)
                Else
                    Diferencia = Damage

                End If
        
                If QuestList(.QuestIndex).RequiredNPCs Then
        
                    For A = 1 To QuestList(.QuestIndex).RequiredNPCs
        
                        If QuestList(.QuestIndex).RequiredNpc(A).NpcIndex = Npclist(NpcIndex).numero Then
                            .NPCsKilled(A) = .NPCsKilled(A) + Abs(Diferencia)

                            If .NPCsKilled(A) >= QuestList(.QuestIndex).RequiredNpc(A).Amount * Npclist(NpcIndex).Stats.MaxHp Then .NPCsKilled(A) = QuestList(.QuestIndex).RequiredNpc(A).Amount * Npclist(NpcIndex).Stats.MaxHp
                              
                            Call Quests_Final(UserIndex, B)
                               
                            'Exit For
                        
                        End If

                    Next A

                End If

            End If

        End With
        
    Next B

    '<EhFooter>
    Exit Sub

Quests_AddNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.Quests_AddNpc " & "at line " & Erl & " in quest: " & TempQuest.QuestIndex

    '</EhFooter>
End Sub

' Comprueba cuando vende un objeto
Public Sub Quests_AddSale(ByVal UserIndex As Integer, _
                          ByVal ObjIndex As Integer, _
                          ByVal Amount As Long)

    '<EhHeader>
    On Error GoTo Quests_AddSale_Err

    '</EhHeader>
    
    Dim Diferencia As Long

    Dim A          As Long
        
    Dim B          As Long
        
    For B = 1 To MAXUSERQUESTS

        With UserList(UserIndex).QuestStats(B)

            If .QuestIndex Then
                If QuestList(.QuestIndex).RequiredSaleOBJs Then
        
                    For A = 1 To QuestList(.QuestIndex).RequiredSaleOBJs
        
                        If QuestList(.QuestIndex).RequiredSaleObj(A).ObjIndex = ObjIndex Then
                            .ObjsSale(A) = .ObjsSale(A) + Amount

                            If .ObjsSale(A) >= QuestList(.QuestIndex).RequiredSaleObj(A).Amount Then .ObjsSale(A) = QuestList(.QuestIndex).RequiredSaleObj(A).Amount
                              
                            Call Quests_Final(UserIndex, B)
                            Exit For

                        End If

                    Next A

                End If

            End If

        End With

    Next B

    '<EhFooter>
    Exit Sub

Quests_AddSale_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.Quests_AddSale " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Comprueba cuando abre un cofre especifico
Public Sub Quests_AddChest(ByVal UserIndex As Integer, _
                           ByVal ObjIndex As Integer, _
                           ByVal Amount As Long)

    '<EhHeader>
    On Error GoTo Quests_AddChest_Err

    '</EhHeader>
    
    Dim Diferencia As Long

    Dim A          As Long
        
    Dim B          As Long
        
    For B = 1 To MAXUSERQUESTS
         
        With UserList(UserIndex).QuestStats(B)

            If .QuestIndex Then
                If QuestList(.QuestIndex).RequiredChestOBJs Then
        
                    For A = 1 To QuestList(.QuestIndex).RequiredChestOBJs
        
                        If QuestList(.QuestIndex).RequiredChestObj(A).ObjIndex = ObjIndex Then
                            .ObjsPick(A) = .ObjsPick(A) + Amount

                            If .ObjsPick(A) >= QuestList(.QuestIndex).RequiredChestObj(A).Amount Then .ObjsPick(A) = QuestList(.QuestIndex).RequiredChestObj(A).Amount
                              
                            Call Quests_Final(UserIndex, B)
                            Exit For

                        End If

                    Next A

                End If

            End If

        End With
        
    Next B
        
    '<EhFooter>
    Exit Sub

Quests_AddChest_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.Quests_AddChest " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Chequea si pasa la misión
Public Sub Quests_Final(ByVal UserIndex As Integer, ByVal Slot As Integer)

    '<EhHeader>
    On Error GoTo Quests_Final_Err

    '</EhHeader>
    If UserList(UserIndex).QuestStats(Slot).QuestIndex = 0 Then Exit Sub

    If Quests_CheckFinish(UserIndex, Slot) Then
        Call mQuests.Quests_Next(UserIndex, Slot)

    End If
        
    Call WriteQuestInfo(UserIndex, False, Slot)

    '<EhFooter>
    Exit Sub

Quests_Final_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.Quests_Final " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Tipea la próxima misión que debera cumplir de manera automatica
Public Sub Quests_Next(ByVal UserIndex As Integer, ByVal Slot As Integer)

    '<EhHeader>
    On Error GoTo Quests_Next_Err

    '</EhHeader>
    
    Dim NextQuest   As Integer

    Dim NextQuest_1 As Integer
    
    With UserList(UserIndex)
        NextQuest = QuestList(.QuestStats(Slot).QuestIndex).NextQuest
        
        Call Quests_Reward(UserIndex, Slot)
        Call CleanQuestSlot(UserList(UserIndex), Slot)
              
        If NextQuest > 0 Then
            If QuestList(NextQuest).RequiredNPCs > 0 Then
                ReDim .QuestStats(Slot).NPCsKilled(1 To QuestList(NextQuest).RequiredNPCs) As Long

            End If
            
            If QuestList(NextQuest).RequiredSaleOBJs > 0 Then
                ReDim .QuestStats(Slot).ObjsSale(1 To QuestList(NextQuest).RequiredSaleOBJs) As Long

            End If
            
            If QuestList(NextQuest).RequiredChestOBJs > 0 Then
                ReDim .QuestStats(Slot).ObjsPick(1 To QuestList(NextQuest).RequiredChestOBJs) As Long

            End If
            
            .QuestStats(Slot).QuestIndex = NextQuest
            
            'NextQuest_1 = QuestList(.QuestStats(Slot).QuestIndex).NextQuest
                
            'If Len(QuestList(NextQuest).Desc) Then
            'Call WriteConsoleMsg(UserIndex, QuestList(NextQuest).Desc, FontTypeNames.FONTTYPE_INFOGREEN)
            'End If
                  
            Call WriteQuestInfo(UserIndex, False, Slot)
                  
        End If
    
    End With

    '<EhFooter>
    Exit Sub

Quests_Next_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.Quests_Next " & "at line " & Erl
        
    '</EhFooter>
End Sub
 
Public Sub CleanQuestSlot(ByRef IUser As User, ByVal Slot As Integer)

    '<EhHeader>
    On Error GoTo CleanQuestSlot_Err

    '</EhHeader>

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Limpia un slot de quest de un usuario.
    'Last modified: 28/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer

    With IUser.QuestStats(Slot)

        If .QuestIndex Then
            If QuestList(.QuestIndex).RequiredNPCs Then

                For i = 1 To QuestList(.QuestIndex).RequiredNPCs
                    .NPCsKilled(i) = 0
                Next i

            End If

            If QuestList(.QuestIndex).RequiredChestOBJs Then

                For i = 1 To QuestList(.QuestIndex).RequiredChestOBJs
                    .ObjsPick(i) = 0
                Next i

            End If

            If QuestList(.QuestIndex).RequiredSaleOBJs Then

                For i = 1 To QuestList(.QuestIndex).RequiredSaleOBJs
                    .ObjsSale(i) = 0
                Next i

            End If

        End If

        .QuestIndex = 0

    End With

    '<EhFooter>
    Exit Sub

CleanQuestSlot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.CleanQuestSlot " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub
 
Public Sub LoadQuests()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Carga el archivo Quests_FilePath en el array QuestList.
    'Last modified: 27/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    On Error GoTo ErrorHandler

    Dim Reader As clsIniManager

    Dim tmpStr As String

    Dim i      As Integer

    Dim j      As Integer
         
    'Cargamos el clsIniManager en memoria
    Set Reader = New clsIniManager
         
    'Lo inicializamos para el archivo Quests_FilePath
    Call Reader.Initialize(Quests_FilePath)
         
    'Redimensionamos el array
    NumQuests = Reader.GetValue("INIT", "NumQuests")
    ReDim QuestList(1 To NumQuests) As tQuest
         
    'Cargamos los datos
    For i = 1 To NumQuests

        With QuestList(i)
            .Nombre = Reader.GetValue("QUEST" & i, "Nombre")
            .Desc = Reader.GetValue("QUEST" & i, "Desc")
            .DescFinish = Reader.GetValue("QUEST" & i, "DescFinal")
            .RequiredLevel = val(Reader.GetValue("QUEST" & i, "RequiredLevel"))
            .DoneQuestMessage = val(Reader.GetValue("QUEST" & i, "DoneQuestMessage"))
            .RequiredBronce = val(Reader.GetValue("QUEST" & i, "RequiredBronce"))
            .DoneQuest = val(Reader.GetValue("QUEST" & i, "DoneQuest"))
            .RequiredPlata = val(Reader.GetValue("QUEST" & i, "RequiredPlata"))
            .RequiredOro = val(Reader.GetValue("QUEST" & i, "RequiredOro"))
            .RequiredPremium = val(Reader.GetValue("QUEST" & i, "RequiredPremium"))
            .LastQuest = val(Reader.GetValue("QUEST" & i, "LastQuest"))
            .NextQuest = val(Reader.GetValue("QUEST" & i, "NextQuest"))
            
            .Remove = val(Reader.GetValue("QUEST" & i, "Remove"))
            
            .RewardDaily = val(Reader.GetValue("QUEST" & i, "RewardDaily"))
            
            If .RewardDaily > 0 Then
                DailyLast = DailyLast + 1
                
                ReDim Preserve QuestDaily(DailyLast) As Byte
            
                QuestDaily(DailyLast) = i

            End If

            .RequiredOBJs = val(Reader.GetValue("QUEST" & i, "RequiredOBJs"))

            If .RequiredOBJs > 0 Then
                ReDim .RequiredObj(1 To .RequiredOBJs)

                For j = 1 To .RequiredOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredOBJ" & j)
                         
                    .RequiredObj(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredObj(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j

            End If
            
            ' Venta de Objetos
            .RequiredSaleOBJs = val(Reader.GetValue("QUEST" & i, "RequiredSaleOBJs"))

            If .RequiredSaleOBJs > 0 Then
                ReDim .RequiredSaleObj(1 To .RequiredSaleOBJs)

                For j = 1 To .RequiredSaleOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredSaleOBJ" & j)
                         
                    .RequiredSaleObj(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredSaleObj(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j

            End If
                
            ' Requiere:: Abrir Cofres de los Mapas
            .RequiredChestOBJs = val(Reader.GetValue("QUEST" & i, "RequiredChestOBJs"))

            If .RequiredChestOBJs > 0 Then
                ReDim .RequiredChestObj(1 To .RequiredChestOBJs)

                For j = 1 To .RequiredChestOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredChestOBJ" & j)
                         
                    .RequiredChestObj(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredChestObj(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j

            End If
            
            'CARGAMOS NPCS REQUERIDOS
            .RequiredNPCs = val(Reader.GetValue("QUEST" & i, "RequiredNPCs"))

            If .RequiredNPCs > 0 Then
                ReDim .RequiredNpc(1 To .RequiredNPCs)

                For j = 1 To .RequiredNPCs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredNPC" & j)
                         
                    .RequiredNpc(j).NpcIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredNpc(j).Amount = val(ReadField(2, tmpStr, 45))
                    .RequiredNpc(j).Hp = val(LeerNPCs.GetValue("NPC" & .RequiredNpc(j).NpcIndex, "MAXHP"))
                Next j

            End If
                 
            .RewardGLD = val(Reader.GetValue("QUEST" & i, "RewardGLD"))
            .RewardEldhir = val(Reader.GetValue("QUEST" & i, "RewardEldhir"))
            .RewardEXP = val(Reader.GetValue("QUEST" & i, "RewardEXP"))
            
            ' Call WriteVar(Quests_FilePath, "QUEST" & i, "RewardGLD", CStr(.RewardGLD))
            '  Call WriteVar(Quests_FilePath, "QUEST" & i, "RewardEXP", CStr(.RewardEXP))
            
            'CARGAMOS OBJETOS DE RECOMPENSA
            .RewardOBJs = val(Reader.GetValue("QUEST" & i, "RewardOBJs"))

            If .RewardOBJs > 0 Then
                ReDim .RewardObj(1 To .RewardOBJs)

                For j = 1 To .RewardOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RewardOBJ" & j)
                         
                    .RewardObj(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RewardObj(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j

            End If

        End With

    Next i
         
    ' Reader.DumpFile Quests_FilePath
    
    'Eliminamos la clase
    Set Reader = Nothing
    
    Call DataServer_Generate_Quests
    
    Exit Sub
                         
ErrorHandler:
    LogError "Error cargando el archivo " & Quests_FilePath

End Sub
 
Public Sub LoadQuestStats(ByVal UserIndex As Integer, ByRef Userfile As clsIniManager)

    '<EhHeader>
    On Error GoTo LoadQuestStats_Err

    '</EhHeader>

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Carga las QuestStats del usuario.
    'Last modified: 28/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

    Dim j      As Integer

    Dim tmpStr As String
        
    Dim A      As Long
        
    ' Adaptation Chars
    If val(Userfile.GetValue("QUESTS", "Q1")) = 999 Then
        
        For A = 1 To MAXUSERQUESTS
            Call mQuests.CleanQuestSlot(UserList(UserIndex), A)
        Next A
            
        Call Quest_SetUserPrincipa(UserIndex)
        Exit Sub

    End If
                      
    For A = 1 To MAXUSERQUESTS
                
        With UserList(UserIndex).QuestStats(A)

            tmpStr = Userfile.GetValue("QUESTS", "Q" & A)
                 
            .QuestIndex = val(ReadField(1, tmpStr, 45))
                      
            If .QuestIndex Then
                If QuestList(.QuestIndex).RequiredNPCs Then
            
                    ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
                         
                    For j = 1 To QuestList(.QuestIndex).RequiredNPCs
                        .NPCsKilled(j) = val(ReadField(j + 1, tmpStr, 45))
                    Next j

                End If
                
                If QuestList(.QuestIndex).RequiredChestOBJs Then
                    ReDim .ObjsPick(1 To QuestList(.QuestIndex).RequiredChestOBJs)
                         
                    For j = 1 To QuestList(.QuestIndex).RequiredChestOBJs
                        .ObjsPick(j) = val(ReadField(QuestList(.QuestIndex).RequiredNPCs + j + 1, tmpStr, 45))
                    Next j

                End If
                
                If QuestList(.QuestIndex).RequiredSaleOBJs Then
                    ReDim .ObjsSale(1 To QuestList(.QuestIndex).RequiredSaleOBJs)
                         
                    For j = 1 To QuestList(.QuestIndex).RequiredSaleOBJs
                        .ObjsSale(j) = val(ReadField(QuestList(.QuestIndex).RequiredChestOBJs + j + 1, tmpStr, 45))
                    Next j

                End If
                    
            End If

        End With
                         
    Next A

    '<EhFooter>
    Exit Sub

LoadQuestStats_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.LoadQuestStats " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub SaveQuestStats(ByRef IQuest() As tUserQuest, ByRef Manager As clsIniManager)

    '<EhHeader>
    On Error GoTo SaveQuestStats_Err

    '</EhHeader>

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Guarda las QuestStats del usuario.
    'Last modified: 29/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i            As Integer

    Dim j            As Integer

    Dim tmpStr       As String
       
    Dim TempRequired As String
        
    Dim A            As Long
        
    For A = 1 To MAXUSERQUESTS

        With IQuest(A)
            tmpStr = .QuestIndex
            TempRequired = vbNullString
                  
            If .QuestIndex Then
                If QuestList(.QuestIndex).RequiredNPCs Then

                    For j = 1 To QuestList(.QuestIndex).RequiredNPCs
                        TempRequired = TempRequired & "-" & CStr(.NPCsKilled(j))
                    Next j
                        
                    tmpStr = tmpStr & TempRequired
                    TempRequired = vbNullString

                End If
                    
                If QuestList(.QuestIndex).RequiredChestOBJs Then

                    For j = 1 To QuestList(.QuestIndex).RequiredChestOBJs
                        TempRequired = TempRequired & "-" & CStr(.ObjsPick(j))
                    Next j

                    tmpStr = tmpStr & TempRequired
                    TempRequired = vbNullString

                End If
                    
                If QuestList(.QuestIndex).RequiredSaleOBJs Then
                    
                    For j = 1 To QuestList(.QuestIndex).RequiredSaleOBJs
                        TempRequired = TempRequired & "-" & CStr(.ObjsSale(j))
                    Next j
                   
                    tmpStr = tmpStr & TempRequired
                    TempRequired = vbNullString

                End If

            End If
             
            Call Manager.ChangeValue("QUESTS", "Q" & A, tmpStr)

        End With
        
    Next A

    '<EhFooter>
    Exit Sub

SaveQuestStats_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.SaveQuestStats " & "at line " & Erl & " en QuestIndex: " & A

    Resume Next

    '</EhFooter>
End Sub
 
Private Function Quests_SearchQuest(ByVal UserIndex As Integer, _
                                    ByVal QuestIndex As Byte) As Boolean

    '<EhHeader>
    On Error GoTo Quests_SearchQuest_Err

    '</EhHeader>

    Dim A As Long
    
    For A = 1 To MAXUSERQUESTS

        With UserList(UserIndex).QuestStats(A)

            If .QuestIndex = QuestIndex Then
                Quests_SearchQuest = True

                Exit Function

            End If

        End With

    Next A

    '<EhFooter>
    Exit Function

Quests_SearchQuest_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mQuests.Quests_SearchQuest " & "at line " & Erl
        
    '</EhFooter>
End Function
