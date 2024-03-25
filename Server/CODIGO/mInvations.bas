Attribute VB_Name = "mInvations"
Option Explicit

Private FilePath As String

Public Sub Invations_New(ByVal Selected As Byte)

    '<EhHeader>
    On Error GoTo Invations_New_Err

    '</EhHeader>
    
    With Invations(Selected)

        If .Run Then Exit Sub
        .Run = True
        .Time = .Duration
        
        Call Invations_Summon(Selected)
    
        Call Invations_Spam(Selected)

    End With

    '<EhFooter>
    Exit Sub

Invations_New_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mInvations.Invations_New " & "at line " & Erl
        
    '</EhFooter>
End Sub

' MainLoop
Public Sub Invations_MainLoop()

    '<EhHeader>
    On Error GoTo Invations_MainLoop_Err

    '</EhHeader>
    Dim A As Long
    
    For A = 1 To Invations_Last

        With Invations(A)

            If .Time > 0 Then
                .Time = .Time - 1
                
                If .Time = 0 Then
                    Call Invations_Close(A)
                Else

                    If (.Time Mod 300) = 0 Then
                        Call Invations_Spam(A)

                    End If

                End If
                
            End If
                
        End With

    Next A

    '<EhFooter>
    Exit Sub

Invations_MainLoop_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mInvations.Invations_MainLoop " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Cargamos la Matrix
Public Sub Invations_Load()

    On Error GoTo ErrHandler

    Dim Manager As clsIniManager

    Dim A       As Long, B As Long, Temp As String

    Dim Maps()  As String
    
    Set Manager = New clsIniManager
        
    FilePath = DatPath & "INV.DAT"
    Manager.Initialize (FilePath)
        
    Invations_Last = val(Manager.GetValue("INIT", "LAST"))
        
    If Invations_Last = 0 Then
        Call LogError("No le llega agua al tanque")
        Set Manager = Nothing
        Exit Sub

    End If
        
    ReDim Invations(1 To Invations_Last) As tInvasion
        
    For A = 1 To Invations_Last

        With Invations(A)
            .Name = Manager.GetValue(CStr(A), "NAME")
            .Desc = Manager.GetValue(CStr(A), "DESC")
            .Duration = val(Manager.GetValue(CStr(A), "DURATION"))
                
            Temp = Manager.GetValue(CStr(A), "INITIALPOS")
            .InitialMap = val(ReadField(1, Temp, Asc("-")))
            .InitialX = val(ReadField(2, Temp, Asc("-")))
            .InitialY = val(ReadField(3, Temp, Asc("-")))
                
            .Npcs = val(Manager.GetValue(CStr(A), "NPCS"))
                
            Maps = Split(Manager.GetValue(CStr(A), "Maps"), "-")
            ReDim .Maps(LBound(Maps) To UBound(Maps))
                
            For B = LBound(Maps) To UBound(Maps)
                .Maps(B) = val(Maps(B))
            Next B

            If .Npcs = 0 Then
                Call LogError("No le llega agua al tanque")
                Set Manager = Nothing
                Exit Sub

            End If
                
            ReDim .Npc(1 To .Npcs) As tInvasionNpc
                
            For B = 1 To .Npcs
                Temp = Manager.GetValue(CStr(A), "NPC" & B)
                    
                .Npc(B).ID = val(ReadField(1, Temp, Asc("-")))
                .Npc(B).cant = val(ReadField(2, Temp, Asc("-")))
                .Npc(B).Map = val(ReadField(3, Temp, Asc("-")))
            Next B

        End With
        
    Next A
        
    Set Manager = Nothing
        
    Exit Sub
    
ErrHandler:
    Set Manager = Nothing

End Sub

' Enviamos el mensaje a la consola del juego.
Private Sub Invations_Spam(ByVal Selected As Byte)

    '<EhHeader>
    On Error GoTo Invations_Spam_Err

    '</EhHeader>
    With Invations(Selected)
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & "» " & .Desc & ". Tipea '/INVASION " & Selected & "' para ingresar ¡CAEN ITEMS! " & IIf(.Time > 60, "Duración " & (.Time / 60) & " minutos", "¡Ultimo minuto!"), FontTypeNames.FONTTYPE_INVASION))

    End With

    '<EhFooter>
    Exit Sub

Invations_Spam_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mInvations.Invations_Spam " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Invoca las Criaturas en el Mapa
Private Sub Invations_Summon(ByVal Selected As Byte)

    '<EhHeader>
    On Error GoTo Invations_Summon_Err

    '</EhHeader>
        
    Dim A        As Long, B As Long, cant As Integer

    Dim Pos      As WorldPos

    Dim NpcIndex As Integer
        
    For A = LBound(Invations(Selected).Npc) To UBound(Invations(Selected).Npc)
            
        cant = Invations(Selected).Npc(A).cant
        Pos.Map = Invations(Selected).Npc(A).Map
            
        For B = 1 To cant
            Pos.X = RandomNumber(20, 85)
            Pos.Y = RandomNumber(20, 85)
         
            NpcIndex = CrearNPC(Invations(Selected).Npc(A).ID, Pos.Map, Pos)
            Npclist(NpcIndex).flags.Invasion = 1
        Next B
    Next A
    
    '<EhFooter>
    Exit Sub

Invations_Summon_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mInvations.Invations_Summon " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Invations_Close(ByVal Selected As Byte)

    On Error GoTo ErrHandler
        
    Dim A As Long

    With Invations(Selected)
        .Run = False
        
        For A = LBound(.Maps) To UBound(.Maps)
            Call Invations_RemoveCriaturesAndUsers(.Maps(A))
        Next A

        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & "» ha terminado. ¡Agradecemos a todos aquellos personajes que participaron de la invasión!", FontTypeNames.FONTTYPE_INVASION))
    
    End With
    
    Exit Sub
ErrHandler:

End Sub

Private Sub Invations_RemoveCriaturesAndUsers(ByVal Map As Integer)

    '<EhHeader>
    On Error GoTo Invations_RemoveCriaturesAndUsers_Err

    '</EhHeader>

    Dim A As Long, B As Long
    
    For A = YMinMapSize To YMaxMapSize
        For B = XMinMapSize To XMaxMapSize

            If InMapBounds(Map, A, B) Then
                If MapData(Map, A, B).NpcIndex > 0 Then
                    If Npclist(MapData(Map, A, B).NpcIndex).flags.Invasion = 1 Then
                        Call QuitarNPC(MapData(Map, A, B).NpcIndex)

                    End If

                End If
                    
                If MapData(Map, A, B).UserIndex <> 0 Then
                    Call EventWarpUser(MapData(Map, A, B).UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y)

                End If

            End If

        Next B
    Next A
          
    '<EhFooter>
    Exit Sub

Invations_RemoveCriaturesAndUsers_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mInvations.Invations_RemoveCriaturesAndUsers " & "at line " & Erl
        
    '</EhFooter>
End Sub

