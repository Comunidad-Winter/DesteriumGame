Attribute VB_Name = "mPesca"
' Sistema de pesca de objetos

Public Type tPesca

    ObjIndex As Integer
    Amount As Integer
    Probability As Byte

End Type

Public Pesca_NumItems As Byte

Public PescaItem()    As tPesca

Public Sub Pesca_LoadItems()

    '<EhHeader>
    On Error GoTo Pesca_LoadItems_Err

    '</EhHeader>

    Dim Manager As clsIniManager

    Dim A       As Long

    Dim Temp    As String
    
    Set Manager = New clsIniManager
    
    Manager.Initialize DatPath & "PESCA.DAT"
    
    Pesca_NumItems = val(Manager.GetValue("INIT", "ITEMS"))
    
    ReDim PescaItem(1 To Pesca_NumItems) As tPesca

    For A = 1 To Pesca_NumItems
        Temp = Manager.GetValue("INIT", A)
        
        With PescaItem(A)
            .ObjIndex = val(ReadField(1, Temp, 45))
            .Amount = val(ReadField(2, Temp, 45))
            .Probability = val(ReadField(3, Temp, 45))

        End With

    Next A
    
    Set Manager = Nothing
    '<EhFooter>
    Exit Sub

Pesca_LoadItems_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mPesca.Pesca_LoadItems " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Pesca_ExtractItem(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo Pesca_ExtractItem_Err

    '</EhHeader>

    Dim A           As Long, B As Long

    Dim RandomItem  As Byte, Random As Byte

    Dim Probability As Byte

    Dim Obj         As Obj
    
    For A = 1 To Pesca_NumItems
        
        With PescaItem(A)

            For B = 1 To .Probability

                ' 10% de ir pasando de etapas
                If RandomNumber(1, 100) <= 10 Then
                    Probability = Probability + 1
                Else

                    Exit For

                End If

            Next B
    
            ' Si cumplimos con la etapa requerida:
            If Probability = .Probability Then
                Obj.ObjIndex = .ObjIndex
                Obj.Amount = .Amount
                            
                If Not MeterItemEnInventario(UserIndex, Obj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, Obj)

                End If
                
                Call WriteConsoleMsg(UserIndex, "Has recolectado de las profundidades del mar " & ObjData(.ObjIndex).Name & " (x" & .Amount & ")", FontTypeNames.FONTTYPE_INFO)

                'Else
                ' Call WriteConsoleMsg(UserIndex, Probability & "/" & .Probability, FontTypeNames.FONTTYPE_INFO)
            End If
            
            Probability = 0

        End With

    Next A

    '<EhFooter>
    Exit Sub

Pesca_ExtractItem_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mPesca.Pesca_ExtractItem " & "at line " & Erl
        
    '</EhFooter>
End Sub
