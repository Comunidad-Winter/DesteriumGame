Attribute VB_Name = "Acciones"
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

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal UserIndex As Integer, _
           ByVal Map As Integer, _
           ByVal X As Integer, _
           ByVal Y As Integer, _
           ByVal Tipo As Byte)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo Accion_Err

    '</EhHeader>

    Dim TempIndex As Integer
    
    '¿Rango Visión? (ToxicWaste)
    If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_x) Then

        Exit Sub

    End If
    
    '¿Posicion valida?
    If InMapBounds(Map, X, Y) Then

        With UserList(UserIndex)

            If MapData(Map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
                TempIndex = MapData(Map, X, Y).NpcIndex
                
                'Set the target NPC
                .flags.TargetNPC = TempIndex
                
                If (Npclist(TempIndex).Comercia = 1 And (Tipo = 1 Or Tipo = 0)) Then

                    '¿Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then

                        Exit Sub

                    End If
                    
                    If Distancia(Npclist(TempIndex).Pos, .Pos) > 5 Then
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                    
                    'Iniciamos la rutina pa' comerciar.
                    Call IniciarComercioNPC(UserIndex)
                    
                ElseIf (Npclist(TempIndex).Quest > 0 And (Tipo = 2 Or Tipo = 0)) Then

                    '¿Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
    
                        Exit Sub
    
                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then

                        Exit Sub

                    End If
                    
                    If Distancia(Npclist(TempIndex).Pos, .Pos) > 5 Then
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                    
                    Call WriteViewListQuest(UserIndex, Npclist(TempIndex).Quests, Npclist(TempIndex).Name)
                    
                ElseIf Npclist(TempIndex).NPCtype = eNPCType.Revividor Or Npclist(TempIndex).NPCtype = eNPCType.ResucitadorNewbie Then

                    If Distancia(.Pos, Npclist(TempIndex).Pos) > 10 Then
                        Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

                        Exit Sub

                    End If
                    
                    'Revivimos si es necesario
                    If .flags.Muerto = 1 And (Npclist(TempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex)) Then
                        Call RevivirUsuario(UserIndex)

                    End If
                    
                    If Npclist(TempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex) Then
                        'curamos totalmente
                        .Stats.MinHp = .Stats.MaxHp
                        Call WriteUpdateUserStats(UserIndex)

                    End If
                    
                ElseIf Npclist(TempIndex).NPCtype = eNPCType.Fundition Then
                    
                ElseIf Npclist(TempIndex).NPCtype = eNPCType.Mascota Then

                    If .MascotaIndex = TempIndex Then
                        Call QuitarPet(UserIndex, .MascotaIndex)
                        Exit Sub
                
                    End If

                ElseIf Npclist(TempIndex).numero = TRAVEL_NPC_HOME Then

                    ' If Distancia(.Pos, Npclist(TempIndex).Pos) > 2 Then
                    '   Call WriteConsoleMsg(UserIndex, "Acercate más y te llevaré de regreso.", FontTypeNames.FONTTYPE_INFO)

                    ' Exit Sub

                    ' End If
                    
                    'Dim Pos As WorldPos
                    
                    ' Pos.Map = Ullathorpe.Map
                    ' Pos.X = Ullathorpe.X
                    '  Pos.Y = Ullathorpe.Y
                    
                    'ClosestStablePos Pos, Pos
                    ' Call WarpUserChar(UserIndex, Pos.Map, Pos.X, Pos.Y, True)
                End If
                
                '¿Es un obj?
            ElseIf MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
                TempIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
                
                .flags.TargetObj = TempIndex
                
                Select Case ObjData(TempIndex).OBJType

                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X, Y, UserIndex)
                    
                    Case eOBJType.otcofre ' Cofres cerrados tirados por el mundo
                        Call AccionParaCofre(Map, X, Y, UserIndex)

                End Select

                '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
            ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
                TempIndex = MapData(Map, X + 1, Y).ObjInfo.ObjIndex
                .flags.TargetObj = TempIndex
                
                Select Case ObjData(TempIndex).OBJType
                    
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
                    
                End Select
            
            ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
                TempIndex = MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = TempIndex
        
                Select Case ObjData(TempIndex).OBJType

                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)

                End Select
            
            ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
                TempIndex = MapData(Map, X, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = TempIndex
                
                Select Case ObjData(TempIndex).OBJType

                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X, Y + 1, UserIndex)

                End Select

            End If

        End With

    End If

    '<EhFooter>
    Exit Sub

Accion_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Acciones.Accion " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub AccionParaPuerta(ByVal Map As Integer, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo AccionParaPuerta_Err

    '</EhHeader>
    
    Dim ObjIndex As Integer
    
    If Not (Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2) Then
        If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
            If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Cerrada = 1 Then

                'Abre la puerta
                If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
                    
                    MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexAbierta
                    
                    ObjIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
                    Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjIndex, ObjData(ObjIndex).GrhIndex, X, Y, ObjData(ObjIndex).Name, 0, ObjData(ObjIndex).Sound))
                    
                    'Desbloquea
                    MapData(Map, X, Y).Blocked = 0
                    MapData(Map, X - 1, Y).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(True, Map, X, Y, 0)
                    Call Bloquear(True, Map, X - 1, Y, 0)
                      
                    'Sonido
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_PUERTA, X, Y))
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                'Cierra puerta
                MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexCerrada
                
                ObjIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
                Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjIndex, ObjData(ObjIndex).GrhIndex, X, Y, ObjData(ObjIndex).Name, 0, ObjData(ObjIndex).Sound))
                                
                MapData(Map, X, Y).Blocked = 1
                MapData(Map, X - 1, Y).Blocked = 1
                
                Call Bloquear(True, Map, X - 1, Y, 1)
                Call Bloquear(True, Map, X, Y, 1)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_PUERTA, X, Y))

            End If
        
            UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex
        Else
            Call WriteConsoleMsg(UserIndex, "La puerta está cerrada con llave.", FontTypeNames.FONTTYPE_INFO)

        End If

    Else
        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

    End If

    '<EhFooter>
    Exit Sub

AccionParaPuerta_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Acciones.AccionParaPuerta " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub AccionParaCofre(ByVal Map As Integer, _
                           ByVal X As Integer, _
                           ByVal Y As Integer, _
                           ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo AccionParaCofre_Err

    '</EhHeader>
    
    Dim ObjIndex   As Integer

    Dim Obj        As ObjData

    Dim ObjAbierto As Obj
    
    Dim DropObj    As Boolean
    
    Dim Time       As Double
    
    Time = GetTime
    DropObj = True

    ObjIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
    Obj = ObjData(ObjIndex)
    
    If UserList(UserIndex).flags.Muerto Then
        Call WriteConsoleMsg(UserIndex, "¡No has logrado abrir el cofre!", FontTypeNames.FONTTYPE_INFO)

        Exit Sub

    End If
    
    If Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2 Then
        Call WriteConsoleMsg(UserIndex, "¡No puedes abrir el cofre desde lejos!", FontTypeNames.FONTTYPE_INFORED)
        
        Exit Sub

    End If
        
    If UserList(UserIndex).Stats.Elv < ObjData(ObjIndex).LvlMin Then
        Call WriteConsoleMsg(UserIndex, "¡Tu nivel no te permite abrir el Cofre!", FontTypeNames.FONTTYPE_INFORED)
        
        Exit Sub

    End If
        
    If (Time - MapData(Map, X, Y).TimeClic) < (Obj.Chest.ClicTime * 1000) Then
        Call WriteConsoleMsg(UserIndex, "¡Haz forzado abrir el cofre antes de tiempo! Debes esperar un poco más...", FontTypeNames.FONTTYPE_INFORED)
        
        Exit Sub

    End If
    
    ' Probabilidad de que el cofre se abra y vuelva a cerrarse
    If RandomNumber(1, 100) <= Obj.Chest.ProbClose Then
        Call WriteConsoleMsg(UserIndex, "No has contado con la suficiente fuerza para abrir el cofre. ¡Se ha cerrado!", FontTypeNames.FONTTYPE_INFORED)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(eSound.sChestClose, X, Y))
        MapData(Map, X, Y).TimeClic = GetTime
        Exit Sub

    End If
                
    ' Probabilidad de que el cofre se abra y se rompa
    If RandomNumber(1, 100) <= Obj.Chest.ProbBreak Then
        Call WriteConsoleMsg(UserIndex, "Parece ser que el cofre se ha roto. ¡Tardará en reconstruirse!", FontTypeNames.FONTTYPE_INFORED)
        DropObj = False ' Se rompe el Cofre
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(eSound.sChestBreak, X, Y))
        GoTo Chesting:

    End If

Chesting:
    
    ' Ponemos el Cofre Abierto/Roto
    If ChestData_Add(Map, X, Y, ObjIndex, Obj.Chest.RespawnTime, DropObj) Then
        If DropObj Then Call Chest_DropObj(UserIndex, ObjIndex, Map, X, Y, False)
        
        ' Chequeamos las Quests
        Call Quests_AddChest(UserIndex, ObjIndex, 1)
            
        ' Quitamos el Cofre Cerrado
        Call EraseObj(MapData(Map, X, Y).ObjInfo.Amount, Map, X, Y)
        
        If DropObj Then

            ObjAbierto.Amount = 1
            ObjAbierto.ObjIndex = ObjData(ObjIndex).IndexAbierta    ' Cofre Abierto
        Else
        
            ObjAbierto.Amount = 1
            ObjAbierto.ObjIndex = ObjData(ObjIndex).IndexCerrada ' Cofre Roto

        End If
        
        Call MakeObj(ObjAbierto, Map, X, Y)
        
    End If
                
    '<EhFooter>
    Exit Sub

AccionParaCofre_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Acciones.AccionParaCofre " & "at line " & Erl
        
    '</EhFooter>
End Sub
