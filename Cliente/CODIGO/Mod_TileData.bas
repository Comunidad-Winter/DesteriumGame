Attribute VB_Name = "Mod_TileData"
' Exodo Online 0.13.5
' #Include Wgl_Client.dll

Option Explicit

Sub CargarCabezas()

    Dim N            As Integer

    Dim i            As Long

    Dim Numheads     As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open IniPath & "Cabezas.ind" For Binary Access Read As #N
    Debug.Print IniPath & "Cabezas.ind"
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)

        End If

    Next i
    
    Close #N

End Sub

Sub CargarAuras()

    Dim A          As Long

    Dim b          As Long
    
    Dim NumAuras   As Integer

    Dim MisAuras() As tIndiceCabeza
    
    Dim Manager    As clsIniManager
    
    Set Manager = New clsIniManager
    
    Dim Temp As String
    
    Manager.Initialize IniPath & "auras.INI"
    
    NumAuras = Val(Manager.GetValue("INIT", "Numheads"))
    
    'Resize array
    ReDim AuraAnimData(0 To NumAuras) As AuraData
    ReDim MisAuras(0 To NumAuras) As tIndiceCabeza
    
    For A = 1 To NumAuras
        For b = 1 To 4
            MisAuras(A).Head(b) = Val(Manager.GetValue("HEAD" & A, "HEAD" & b))
            
            Temp = Manager.GetValue("HEAD" & A, "COLOR")
            AuraAnimData(A).Color = ARGB(Val(ReadField(1, Temp, 45)), Val(ReadField(2, Temp, 45)), Val(ReadField(3, Temp, 45)), Val(ReadField(4, Temp, 45)))
            
            If MisAuras(A).Head(b) Then
                Call InitGrh(AuraAnimData(A).Walk(b), MisAuras(A).Head(b), 0)
    
            End If

        Next b
    Next A
    
    Set Manager = Nothing

End Sub

Sub CargarCascos()

    Dim N            As Integer

    Dim i            As Long

    Dim NumCascos    As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open IniPath & "Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)

        End If

    Next i
    
    Close #N

End Sub

Sub CargarCuerpos()

    Dim N            As Integer

    Dim i            As Long, b As Long

    Dim NumCuerpos   As Integer

    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open IniPath & "Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then

            For b = 1 To 4
                InitGrh BodyData(i).Walk(b), MisCuerpos(i).Body(b), 0
                
                BodyData(i).BodyOffSet(b).X = MisCuerpos(i).BodyOffSetX(b)
                BodyData(i).BodyOffSet(b).Y = MisCuerpos(i).BodyOffSetY(b)
            Next b
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
            
            #If ModoBig > 0 Then
                BodyData(i).HeadOffset.Y = BodyData(i).HeadOffset.Y * 2
            #End If

        End If

    Next i
    
    Close #N

End Sub

Sub CargarCuerposAttack()

    Dim N            As Integer

    Dim i            As Long

    Dim NumCuerpos   As Integer

    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open IniPath & "Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyDataAttack(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            InitGrh BodyDataAttack(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyDataAttack(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyDataAttack(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyDataAttack(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyDataAttack(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyDataAttack(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY

        End If

    Next i
    
    Close #N

End Sub

Sub CargarFxs()

    Dim N      As Integer

    Dim i      As Long

    Dim NumFxs As Integer
    
    N = FreeFile()
    Open IniPath & "\Fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N

End Sub

