Attribute VB_Name = "mChest"
' // Todo este procesamiento debe ir en un servidor externo el día de mañana.

Option Explicit

Public Type tChestData

    Map As Integer
    X As Byte
    Y As Byte
    
    ObjIndex As Integer ' Objeto que saldrá de la tierra
    Time As Long

End Type

Public Const MAX_CHESTDATA           As Integer = 500

Public ChestLast                     As Integer

Public ChestData(1 To MAX_CHESTDATA) As tChestData

' Busca un Slot libre para agregar le cofre al conteo de respawn
Private Function ChestData_Slot() As Integer

    '<EhHeader>
    On Error GoTo ChestData_Slot_Err

    '</EhHeader>
    Dim A As Long
    
    For A = 1 To MAX_CHESTDATA

        If ChestData(A).Map = 0 Then
            ChestData_Slot = A
            Exit Function
        
        End If

    Next A

    '<EhFooter>
    Exit Function

ChestData_Slot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mChest.ChestData_Slot " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function ChestData_Add(ByVal Map As Integer, _
                              ByVal X As Byte, _
                              ByVal Y As Byte, _
                              ByVal ObjIndex As Integer, _
                              ByVal Time As Long, _
                              ByVal DropObj As Boolean) As Boolean

    '<EhHeader>
    On Error GoTo ChestData_Add_Err

    '</EhHeader>
    
    Dim Slot As Byte
    
    Slot = ChestData_Slot
    
    If Slot = 0 Then
        Call LogError("¡¡ERROR AL AGREGAR UN COFRE EN EL MAPA " & Map & " " & X & " " & Y)
    Else
        
        If Not DropObj Then
            Time = Time * 1.5 ' 50% más de tiempo para que regenere en caso de haber roto

        End If
        
        With ChestData(Slot)
            .Map = Map
            .X = X
            .Y = Y
            .Time = Time
            .ObjIndex = ObjIndex

        End With
        
        ChestData_Add = True

    End If

    '<EhFooter>
    Exit Function

ChestData_Add_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mChest.ChestData_Add " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub ChestLoop()

    '<EhHeader>
    On Error GoTo ChestLoop_Err

    '</EhHeader>
    Dim A         As Long

    Dim ChestNull As tChestData

    Dim Obj       As Obj
    
    For A = 1 To MAX_CHESTDATA

        With ChestData(A)
            
            If .Map > 0 Then
                .Time = .Time - 1

                If .Time = 0 Then
                    Obj.ObjIndex = .ObjIndex
                    Obj.Amount = 1
                    
                    Call EraseObj(MapData(.Map, .X, .Y).ObjInfo.Amount, .Map, .X, .Y)
                    Call MakeObj(Obj, .Map, .X, .Y)
                    Call SendToAreaByPos(.Map, .X, .Y, PrepareMessagePlayEffect(eSound.sChestClose, .X, .Y))
                    ChestData(A) = ChestNull
                
                End If
            
            End If
        
        End With
    
    Next A

    '<EhFooter>
    Exit Sub

ChestLoop_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mChest.ChestLoop " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Chest_DropObj(ByVal UserIndex As Integer, _
                         ByVal ObjIndex As Integer, _
                         ByVal Map As Integer, _
                         ByVal X As Byte, _
                         ByVal Y As Byte, _
                         ByVal DropInv As Boolean)

    '<EhHeader>
    On Error GoTo Chest_DropObj_Err

    '</EhHeader>
    
    Dim DropObj    As Obj

    Dim A          As Long

    Dim RandomDrop As Byte

    Dim nPos       As WorldPos

    Dim Pos        As WorldPos
    
    Dim Sound      As Integer
    
    Dim Random     As Byte
            
    Random = RandomNumber(1, ObjData(ObjIndex).Chest.NroDrop)
            
    RandomDrop = ObjData(ObjIndex).Chest.Drop(Random)
    Pos.Map = Map
    Pos.X = X
    Pos.Y = Y
    
    With DropData(RandomDrop)

        For A = 1 To .Last
            DropObj.ObjIndex = .Data(A).ObjIndex
            DropObj.Amount = RandomNumber(.Data(A).Amount(0), .Data(A).Amount(1))

            If RandomNumber(1, 100) <= .Data(A).Prob Then
                If DropInv Then
                    If Not MeterItemEnInventario(UserIndex, DropObj) Then
                        Call TirarItemAlPiso(UserList(UserIndex).Pos, DropObj)
    
                    End If
    
                Else
                    Call Tilelibre(Pos, nPos, DropObj, False, True)
    
                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call MakeObj(DropObj, nPos.Map, nPos.X, nPos.Y)
                        nPos.Map = 0
                        nPos.X = 0
                        nPos.Y = 0
    
                    End If
    
                End If

            End If
                
        Next A
            
        Call Chest_PlaySound(UserIndex, X, Y)

    End With
    
    '<EhFooter>
    Exit Sub

Chest_DropObj_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mChest.Chest_DropObj " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub Chest_AbreFortuna(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)

    '<EhHeader>
    On Error GoTo Chest_AbreFortuna_Err

    '</EhHeader>
    Dim Num As Long

    Dim Obj As Obj

    Num = RandomNumber(1, ObjData(ObjIndex).MaxFortunas)

    If Not MeterItemEnInventario(UserIndex, ObjData(ObjIndex).Fortuna(Num)) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, ObjData(ObjIndex).Fortuna(Num))

    End If
                    
    Call WriteConsoleMsg(UserIndex, "¡Has recibido " & ObjData(ObjData(ObjIndex).Fortuna(Num).ObjIndex).Name & " (x" & ObjData(ObjIndex).Fortuna(Num).Amount & ")!", FontTypeNames.FONTTYPE_INFOGREEN)
    Call Chest_PlaySound(UserIndex, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    
    '<EhFooter>
    Exit Sub

Chest_AbreFortuna_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mChest.Chest_AbreFortuna " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub Chest_PlaySound(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

    '<EhHeader>
    On Error GoTo Chest_PlaySound_Err

    '</EhHeader>

    Dim Random As Byte

    Dim Sound  As Long
    
    Random = RandomNumber(1, 100)
        
    If Random <= 25 Then
        Sound = eSound.sChestDrop1
    ElseIf Random <= 50 Then
        Sound = eSound.sChestDrop2
    Else
        Sound = eSound.sChestDrop3

    End If
        
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(Sound, X, Y))

    '<EhFooter>
    Exit Sub

Chest_PlaySound_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mChest.Chest_PlaySound " & "at line " & Erl
        
    '</EhFooter>
End Sub
