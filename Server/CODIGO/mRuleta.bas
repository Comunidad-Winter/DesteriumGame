Attribute VB_Name = "mRuleta"
Option Explicit

' Sistema de Ruleta

Public Type tRuletaItem

    ObjIndex As Integer      ' Index del objeto
    Amount As Integer       ' Cantidad de objeto que da
    Prob As Byte                '1,2,3,4,5
    ProbNum As Byte         '10,20,30,40,50,60,70,80,90a99

End Type

Public Type tRuletaConfig

    ItemLast As Integer
    Items() As tRuletaItem
    RuletaGld As Long
    RuletaDsp As Long

End Type

Public RuletaConfig As tRuletaConfig

Public Sub Ruleta_LoadItems()

    Dim Manager  As clsIniManager

    Dim A        As Long

    Dim Temp     As String
    
    Dim FilePath As String
    
    Set Manager = New clsIniManager
    
    FilePath = DatPath & "ruleta.dat"
    
    Manager.Initialize FilePath
    
    With RuletaConfig
        .ItemLast = val(Manager.GetValue("INIT", "LAST"))
        .RuletaDsp = val(Manager.GetValue("INIT", "RULETADSP"))
        .RuletaGld = val(Manager.GetValue("INIT", "RULETAGLD"))
        
        If .ItemLast > 0 Then
            ReDim .Items(1 To .ItemLast) As tRuletaItem
        
            For A = 1 To .ItemLast

                With .Items(A)
                    Temp = Manager.GetValue("LIST", "OBJ" & A)
                
                    .ObjIndex = val(ReadField(1, Temp, 45))
                    .Amount = val(ReadField(2, Temp, 45))
                    .Prob = val(ReadField(3, Temp, 45))
                    .ProbNum = val(ReadField(4, Temp, 45))
                
                End With

            Next A
    
        End If
    
    End With
    
    Manager.DumpFile DatPath & "client\ruleta.dat"
    Set Manager = Nothing

End Sub

Public Sub Ruleta_Tirada(ByVal UserIndex As Integer, ByVal Mode As Byte)

    '<EhHeader>
    On Error GoTo Ruleta_Tirada_Err

    '</EhHeader>

    With UserList(UserIndex)

        Exit Sub

        If Mode = 1 Then ' Monedas de Oro
            If .Stats.Gld < RuletaConfig.RuletaGld Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficientes Monedas de Oro.", FontTypeNames.FONTTYPE_INFORED)
                'TODO: Enter task description here
                Exit Sub

            End If
            
            .Stats.Gld = .Stats.Gld - RuletaConfig.RuletaGld
            Call WriteUpdateGold(UserIndex)
        ElseIf Mode = 2 Then            ' Monedas DSP

            If .Stats.Eldhir < RuletaConfig.RuletaDsp Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficientes DSP.", FontTypeNames.FONTTYPE_INFORED)
                Exit Sub

            End If
            
            .Stats.Eldhir = .Stats.Eldhir - RuletaConfig.RuletaDsp
            Call WriteUpdateDsp(UserIndex)

        End If
        
        Call Ruleta_Tirada_Item(UserIndex)

    End With

    '<EhFooter>
    Exit Sub

Ruleta_Tirada_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mRuleta.Ruleta_Tirada " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Public Sub Ruleta_Tirada_Item(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo Ruleta_Tirada_Item_Err

    '</EhHeader>
    
    Dim RandItem    As Byte

    Dim A           As Long, S As Long
    
    Dim Probability As Long, Sound As Long
    
    Dim MiObj       As Obj
    
    RandItem = RandomNumber(1, RuletaConfig.ItemLast)
    
    With RuletaConfig.Items(RandItem)

        For A = 1 To .Prob

            If RandomNumber(1, 100) <= .ProbNum Then
                Probability = Probability + 1

            End If

        Next A
                
        If Probability = .Prob Then
            MiObj.Amount = .Amount
            MiObj.ObjIndex = .ObjIndex

            If Not MeterItemEnInventario(UserIndex, MiObj, True) Then
                Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

            End If
        
            Exit Sub
            
            S = RandomNumber(1, 100)
        
            If S <= 25 Then
                Sound = eSound.sChestDrop1
            ElseIf S <= 50 Then
                Sound = eSound.sChestDrop2
            Else
                Sound = eSound.sChestDrop3

            End If
        
            Call SendData(SendTarget.ToOne, UserIndex, PrepareMessagePlayEffect(Sound, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

        End If
    
    End With

    '<EhFooter>
    Exit Sub

Ruleta_Tirada_Item_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mRuleta.Ruleta_Tirada_Item " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub
