Attribute VB_Name = "Mod_General"
Option Explicit

'Private Declare Sub InitCommonControls Lib "comctl32" ()

Public bFogata   As Boolean

Public bLluvia() As Byte ' Array para determinar si

'debemos mostrar la animacion de la lluvia
Public Declare Function URLDownloadToFile _
               Lib "urlmon" _
               Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
                                           ByVal szURL As String, _
                                           ByVal szFileName As String, _
                                           ByVal dwReserved As Long, _
                                           ByVal lpfnCB As Long) As Long

'Very percise counter 64bit system counter
Public Declare Function QueryPerformanceFrequency _
               Lib "kernel32" (lpFrequency As Currency) As Long

Public Declare Function QueryPerformanceCounter _
               Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public Const C_CHARACTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890abcdefghijklmnopqrstuvwxyz"

Private Declare Function SetParent _
                Lib "user32.dll" (ByVal hWndChild As Long, _
                                  ByVal hWndNewParent As Long) As Long

Public Sub Setting_MenuInfo(ByVal NpcIndex As Integer, ByVal DoubleClic As Boolean)

    '<EhHeader>
    On Error GoTo Setting_MenuInfo_Err

    '</EhHeader>

    Dim L As Long

    Dim T As Long
        
    SelectedNpcIndex = NpcIndex
    NpcIndex_MouseHover = 0
        
    If DoubleClic Then
        If NpcList(SelectedNpcIndex).MaxHp > 0 Then Exit Sub

    End If
        
    If Not MirandoOpcionesNpc Then
        FrmListAcciones.Show , FrmMain
    Else
        FrmListAcciones.Initial_Form

    End If
            
    L = FrmMain.Left + (FrmMain.MainViewPic.Left * Screen.TwipsPerPixelX) + (FrmMain.MouseX * Screen.TwipsPerPixelX) - 50
    T = FrmMain.Top + (FrmMain.MainViewPic.Top * Screen.TwipsPerPixelX) + (32 * Screen.TwipsPerPixelY) + (FrmMain.MouseY * Screen.TwipsPerPixelY) - 50
    Call FrmListAcciones.Move(L, T)

    '<EhFooter>
    Exit Sub

Setting_MenuInfo_Err:
    LogError err.Description & vbCrLf & "in ARGENTUM.frmMain_Scalled.Setting_MenuInfo " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Public Function GetSystemTime() As Long

    Static Frequency As Currency

    Static offset    As Currency
    
    ' Lazy initialization of timer frequency and offset
    If (Frequency = 0) Then
        Call QueryPerformanceFrequency(Frequency)
        Call QueryPerformanceCounter(offset)
        
        GetSystemTime = 0
    Else

        Dim Value As Currency

        Call QueryPerformanceCounter(Value)
        
        GetSystemTime = ((Value - offset) / Frequency * 1000)

    End If

End Function

Public Function DirGraficos() As String
    DirGraficos = App.path & "\AO\resource\"

End Function

Public Function DirInterface() As String
    DirInterface = App.path & "\AO\resource\interface\"

End Function

Public Function DirSound() As String
    DirSound = App.path & "\AO\resource\" & Config_Inicio.DirSonidos & "\"

End Function

Public Function DirMusic() As String
    DirMusic = App.path & "\AO\resource\" & Config_Inicio.DirMusica & "\"

End Function

Public Function DirMapas() As String
    DirMapas = App.path & "\AO\resource\" & Config_Inicio.DirMapas & "\"

End Function

Public Function DirExtras() As String
    DirExtras = App.path & "\temp\"

End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound

End Function

Private Function PacketID_Change(ByVal Selected As Byte) As Integer
    
    Dim Temp     As Integer

    Dim KeyText  As String

    Dim KeyValue As String
    
    Select Case Selected

        Case 75
            KeyValue = "GAHBDEWIDKFLSQ2DIWJNE"

        Case 150
            KeyValue = "AGSQEFHFFDFSDQETUHFLSJNE"

        Case 99
            KeyValue = "13SDDJS2s"

        Case 105
            KeyValue = "ADSDEWEFFDFGRT"

    End Select
    
    Temp = 127
    Temp = Temp Xor 45
    
    If Len(KeyValue) > 10 Then
        Temp = Temp Xor 4 Xor Selected
    Else
        Temp = Temp Xor 75

    End If
    
End Function

Public Function ReadPacketID(ByRef PacketID As Integer) As Integer
    
    Dim KeyTempOne   As Integer

    Dim KeyTempTwo   As Integer

    Dim KeyTempThree As Integer
    
    Dim KeyOne       As String: KeyOne = "137"

    Dim KeyTwo       As String: KeyTwo = "215"

    Dim KeyThree     As String: KeyThree = "45"

    Dim KeyFour      As String: KeyFour = "12"

    Dim KeyFive      As String: KeyFive = "197"
    
    PacketID = PacketID Xor 127
    KeyTempOne = 127
    PacketID = PacketID Xor 67
    PacketID = PacketID Xor Len(KeyOne)
    KeyTempOne = KeyTempOne Xor 12
    
    PacketID = PacketID Xor PacketID_Change(99)
    
    If PacketID Then
        PacketID = PacketID Xor Len(KeyTwo)
        PacketID = PacketID Xor Len(KeyThree)
        
        PacketID = PacketID Xor PacketID_Change(75)
    Else
        PacketID = PacketID Xor Len(KeyOne)
        PacketID = PacketID Xor Len(KeyThree)
        PacketID = PacketID Xor PacketID_Change(99)

    End If
    
    KeyTempOne = KeyTempOne Xor PacketID
    
    If KeyTempOne > 55 Then
        KeyTempTwo = KeyTempTwo Xor 49
        KeyTempThree = KeyTempThree Xor 75
    ElseIf KeyTempOne > 150 Then
        KeyTempTwo = KeyTempTwo Xor 49
        KeyTempThree = KeyTempThree Xor 75
    ElseIf KeyTempOne > 250 Then
        KeyTempTwo = KeyTempTwo Xor 49

    End If
    
    PacketID = PacketID Xor KeyOne
    KeyTempTwo = KeyTempTwo Xor KeyTempOne Xor PacketID_Change(150)
    KeyTempThree = KeyTempOne Xor KeyTempTwo
    PacketID = PacketID Xor 75 Xor PacketID_Change(105)
    
    KeyTempTwo = PacketID Xor KeyTempThree
    PacketID = PacketID Xor 21
    
    PacketID = PacketID Xor Len(KeyFive)
    
    ReadPacketID = PacketID

End Function

Public Function GetRawName(ByRef sName As String) As String
    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 13/01/2010
    'Last Modified By: -
    'Returns the char name without the clan name (if it has it).
    '***************************************************

    Dim Pos As Integer
    
    Pos = InStr(1, sName, "<")
    
    If Pos > 0 Then
        GetRawName = Trim(Left(sName, Pos - 1))
    Else
        GetRawName = sName

    End If

End Function

Sub CargarAnimArmas()

    On Error Resume Next

    Dim LoopC As Long

    Dim arch  As String
    
    arch = IniPath & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For LoopC = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(LoopC).WeaponWalk(1), Val(GetVar(arch, "ARMA" & LoopC, "Dir1")), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(2), Val(GetVar(arch, "ARMA" & LoopC, "Dir2")), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(3), Val(GetVar(arch, "ARMA" & LoopC, "Dir3")), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(4), Val(GetVar(arch, "ARMA" & LoopC, "Dir4")), 0
    Next LoopC

End Sub

Sub CargarDialogos()

    On Error Resume Next

    '***************************************************
    'Author: Juan Dalmasso (CHOTS)
    'Last Modify Date: 11/06/2011
    '***************************************************
    Dim archivoC As String
    
    archivoC = IniPath & "dialogos.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
        'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los di�logos. Falta el archivo dialogos.dat, reinstale el juego", vbCritical + vbOKOnly)

        Exit Sub

    End If
    
    Dim i As Byte
    
    For i = 1 To MAXCOLORESDIALOGOS
        ColoresDialogos(i).r = CByte(GetVar(archivoC, CStr(i), "R"))
        ColoresDialogos(i).g = CByte(GetVar(archivoC, CStr(i), "G"))
        ColoresDialogos(i).b = CByte(GetVar(archivoC, CStr(i), "B"))
    Next i

End Sub

Sub CargarColores()

    On Error Resume Next

    Dim archivoC As String
    
    archivoC = IniPath & "colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
        'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)

        Exit Sub

    End If
    
    Dim i As Long
    
    For i = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i).r = CByte(GetVar(archivoC, CStr(i), "R"))
        ColoresPJ(i).g = CByte(GetVar(archivoC, CStr(i), "G"))
        ColoresPJ(i).b = CByte(GetVar(archivoC, CStr(i), "B"))
    Next i
    
    ' Crimi
    ColoresPJ(50).r = CByte(GetVar(archivoC, "CR", "R"))
    ColoresPJ(50).g = CByte(GetVar(archivoC, "CR", "G"))
    ColoresPJ(50).b = CByte(GetVar(archivoC, "CR", "B"))
    
    ' Ciuda
    ColoresPJ(49).r = CByte(GetVar(archivoC, "CI", "R"))
    ColoresPJ(49).g = CByte(GetVar(archivoC, "CI", "G"))
    ColoresPJ(49).b = CByte(GetVar(archivoC, "CI", "B"))

End Sub

Sub CargarAnimEscudos()

    On Error Resume Next

    Dim LoopC As Long

    Dim arch  As String
    
    arch = IniPath & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For LoopC = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(LoopC).ShieldWalk(1), Val(GetVar(arch, "ESC" & LoopC, "Dir1")), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(2), Val(GetVar(arch, "ESC" & LoopC, "Dir2")), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(3), Val(GetVar(arch, "ESC" & LoopC, "Dir3")), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(4), Val(GetVar(arch, "ESC" & LoopC, "Dir4")), 0
    Next LoopC

End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, _
                     ByVal Text As String, _
                     Optional ByVal red As Integer = -1, _
                     Optional ByVal green As Integer, _
                     Optional ByVal blue As Integer, _
                     Optional ByVal bold As Boolean = True, _
                     Optional ByVal italic As Boolean = False, _
                     Optional ByVal bCrLf As Boolean = True)
    '******************************************
    'Adds text to a Richtext box at the bottom.
    'Automatically scrolls to new text.
    'Text box MUST be multiline and have a 3D
    'apperance!
    'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
    'Juan Mart�n Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
    '08/02/12 (D'Artagnan) - Divisi�n de consolas
    '******************************************r

    On Error Resume Next

    With RichTextBox
        
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF

        End If
                
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = True 'bold
        .SelItalic = italic
        
        #If ModoBig > 0 Then
            #If FullScreen = 1 Then
                .SelFontSize = 8
            #Else
                .SelFontSize = 10
            #End If
        #Else
            .SelFontSize = 7
        #End If
                
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
                
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        .SelText = Text

    End With

End Sub

Sub AddtoConsole(ByRef RichTextBox As RichTextBox, _
                 ByVal Text As String, _
                 Optional ByVal red As Integer = -1, _
                 Optional ByVal green As Integer, _
                 Optional ByVal blue As Integer, _
                 Optional ByVal bold As Boolean = False, _
                 Optional ByVal italic As Boolean = False, _
                 Optional ByVal bCrLf As Boolean = True)

    '******************************************
    'Author: D'Artagnan (aop.fran@gmail.com)
    'Auxiliar sub for adding console messages
    '******************************************
    With RichTextBox

        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF

        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        .SelText = Text
        
        RichTextBox.Refresh

    End With

End Sub

Public Function GetConsoleIndex(ByVal msg As eMessageType) As Byte

    Select Case msg
    
        Case eMessageType.cEvents_Curso
        
        Case eMessageType.cEvents_General
        
    End Select

End Function

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()

    '*****************************************************************
    'Goes through the charlist and replots all the characters on the map
    'Used to make sure everyone is visible
    '*****************************************************************
    Dim LoopC As Long
    
    For LoopC = 1 To LastChar

        If CharList(LoopC).Active = 1 Then
            MapData(CharList(LoopC).Pos.X, CharList(LoopC).Pos.Y).CharIndex = LoopC

        End If

    Next LoopC

End Sub

Sub SaveGameini()
    'Grabamos los datos del usuario en el Game.ini
    Config_Inicio.Name = "BetaTester"
    Config_Inicio.Password = "DammLamers"
    Config_Inicio.Puerto = UserPort
    
    Call EscribirGameIni(Config_Inicio)

End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim car As Byte

    Dim i   As Integer

    cad = LCase$(cad)

    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
          
        If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
            AsciiValidos = False

            Exit Function

        End If
          
    Next i

    AsciiValidos = True

End Function

Function CheckUserData() As Boolean

    'Validamos los datos del user
    Dim LoopC     As Long

    Dim CharAscii As Integer
        
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")

        Exit Function

    End If
    
    For LoopC = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, LoopC, 1))

        If Not LegalCharacter(CharAscii, False) Then
            MsgBox ("Password inv�lido. El caract�r " & Chr$(CharAscii) & " no est� permitido.")

            Exit Function

        End If

    Next LoopC
    
    If Len(UserName) > 10 Then
        MsgBox ("El nombre debe tener menos de 10 letras.")

        Exit Function

    End If
    
    For LoopC = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, LoopC, 1))

        If Not LegalCharacter(CharAscii, True) Then
            MsgBox ("Nombre inv�lido. El caract�r " & Chr$(CharAscii) & " no est� permitido.")

            Exit Function

        End If

    Next LoopC
    
    CheckUserData = True

End Function

Sub UnloadAllForms_ButPrincipal()

    On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms

        If MirandoObjetos Then
            If (mifrm.Name <> FrmMenu.Name) And mifrm.Name <> FrmMain.Name And mifrm.Name <> FrmObject_Info.Name Then
                Unload mifrm

            End If

        Else

            If (mifrm.Name <> FrmMenu.Name) And mifrm.Name <> FrmMain.Name And mifrm.Name <> FrmObject_Info.Name Then
                Unload mifrm

            End If

        End If

    Next

    FrmMain.SetFocus

End Sub

Sub UnloadAllForms()

    On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms

        Unload mifrm
    Next

End Sub

Function LegalCharacter(ByVal KeyAscii As Integer, ByVal inLogin As Boolean) As Boolean

    '*****************************************************************
    'Only allow characters that are Win 95 filename compatible
    '*****************************************************************
    Dim i As Long

    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True

        Exit Function

    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then

        Exit Function

    End If
    
    If KeyAscii > 126 Then
        If inLogin Then 'Est� chequeando un logueo

            For i = 1 To Len(CAR_ESPECIALES)

                If KeyAscii = Asc(mid$(CAR_ESPECIALES, i, 1)) Then
                    LegalCharacter = True

                    Exit Function

                End If

            Next i

            Exit Function

        Else

            Exit Function

        End If

    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then

        Exit Function

    End If
    
    'else everything is cool
    LegalCharacter = True

End Function

Sub SetConnected()
    '*****************************************************************
    'Sets the client to "Connect" mode
    '*****************************************************************
    'Set Connected
    Connected = True
    
    ' Reset Flags loggin
    UserEvento = False
        
    'TempAccount.Passwd = vbNullString
    'TempAccount.Email = vbNullString
        
    Call SaveGameini
    SecurityKey_Number = 0
    FightOn = False
    
    Call Audio.StopMusic
    
    If CreandoPersonaje Then
    
        CreandoPersonaje = False

        'Account.Chars(SearchFreeChar).Name = UserName
    End If
    
    Account.SelectedChar = 0
    MirandoCuenta = False
    
    'Vaciamos la cola de movimiento
    Call keysMovementPressedQueue.Clear

    Unload FrmConnect_Account

    'Unload the connect form
    'Unload frmCrearPersonaje
    'Unload frmConnect
    
    Dim A As Long

    For A = FrmMain.Label8.LBound To FrmMain.Label8.UBound
        FrmMain.Label8(A).Caption = UserName
    Next A
    
    'Load main form
    FrmMain.visible = True
    
    FrmMain.tUpdateInactive.Enabled = True
    FrmMain.Second.Enabled = True
    FrmMain.tUpdateMS.Enabled = True
    FrmMain.tUpdateInactive.Enabled = True
    
    #If ModoBig = 1 Then
        dockForm FrmMenu.hWnd, FrmMain.PicMenu, True
    #End If

    If Len(CaptionTemp) > 0 Then
        Call WriteDenounce("[SEGURIDAD]: Posible " & CaptionTemp)
        CaptionTemp = vbNullString

    End If
    
    If Len(TempModuleName) > 0 Then
        Call WriteDenounce("[SEGURIDAD]: Posible " & TempModuleName)
        TempModuleName = vbNullString

    End If
    
    Call Draw_MiniMap
    
    'Are we under a roof?
    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, UserPos.Y).Trigger = 2 Or MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)

End Sub

Public Sub Draw_MiniMap()
        
    ' RenderScreen_MiniMap
    ' Call Map_UpdateLabel
    '
    ' Exit Sub
    
    If FileExist(MiniMap_FilePath & UserMap & ".png", vbArchive) Then
        RenderScreen_MiniMapa_PNG UserMap
    Else

        If Not RenderizandoMap Then
            RenderizandoMap = True
            RenderizandoIndex = UserMap
            FrmMain.UpdateMapa.Enabled = True

        End If
        
        RenderScreen_MiniMapa

    End If
    
    ' Call Map_UpdateLabel

End Sub

Sub MoveTo(ByVal Direccion As E_Heading)

    '***************************************************
    'Author: Alejandro Santos (AlejoLp)
    'Last Modify Date: 06/28/2008
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    ' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
    ' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
    ' 06/28/2008: NicoNZ - Saqu� lo que imped�a que si el usuario estaba paralizado se ejecute el sub.
    '***************************************************
    On Error GoTo ErrHandler
    
    Dim LegalOk As Boolean

    Select Case Direccion

        Case E_Heading.NORTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y - 1)

        Case E_Heading.EAST
            LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.Y)

        Case E_Heading.SOUTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)

        Case E_Heading.WEST
            LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.Y)

    End Select
    
    If LegalOk And Not UserParalizado Then
        
        If Not UserDescansar Then

            '  If UserMeditar Then Exit Sub
            If UserEvento Then Exit Sub
            If FrmMain.MacroTrabajo.Enabled Then Call FrmMain.DesactivarMacroTrabajo

            Moviendose = True
            Call MainTimer.Restart(TimersIndex.Walk)
                
            Call WriteWalk(Direccion)
                
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion
           
        Else
        
        End If
    
    Else

        If UserCharIndex > 0 Then
            If CharList(UserCharIndex).Heading <> Direccion Then
                If IntervaloPermiteHeading(True) Then
                    Call WriteChangeHeading(Direccion)

                End If

            End If

        End If

    End If
        
    If MirandoOpcionesNpc Then
        Unload FrmListAcciones

    End If
            
    If MirandoComerciar Then
        Unload frmComerciar

    End If
            
    If MirandoBanco Then
        Unload frmBancoObj

    End If

    Call Map_UpdateLabel(False)

    Exit Sub
ErrHandler:
    
End Sub

Sub RandomMove()
    '***************************************************
    'Author: Alejandro Santos (AlejoLp)
    'Last Modify Date: 06/03/2006
    ' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
    '***************************************************
    Call MoveTo(RandomNumber(NORTH, WEST))

End Sub

Public Sub RequestMeditate()
    'If (Not esGM(UserCharIndex)) Then
    'WaitInput = True
    ' End If
    
    Call WriteMeditate

End Sub

Public Sub AddMovementToKeysMovementPressedQueue()
    
    On Error GoTo AddMovementToKeysMovementPressedQueue_Err
    
    If CustomKeys.BindedKey(mKeyDown) <> 0 And GetKeyState(CustomKeys.BindedKey(mKeyDown)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(mKeyDown)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(mKeyDown))   ' Agrega la tecla al arraylist
    Else

        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(mKeyDown)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(mKeyDown))   ' Remueve la tecla que teniamos presionada

    End If

    If CustomKeys.BindedKey(mKeyLeft) <> 0 And GetKeyState(CustomKeys.BindedKey(mKeyLeft)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(mKeyLeft)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(mKeyLeft)) ' Agrega la tecla al arraylist
    Else

        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(mKeyLeft)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(mKeyLeft))  ' Remueve la tecla que teniamos presionada

    End If

    If CustomKeys.BindedKey(mKeyUp) <> 0 And GetKeyState(CustomKeys.BindedKey(mKeyUp)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(mKeyUp)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(mKeyUp))     ' Agrega la tecla al arraylist
    Else

        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(mKeyUp)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(mKeyUp))     ' Remueve la tecla que teniamos presionada

    End If

    If CustomKeys.BindedKey(mKeyRight) <> 0 And GetKeyState(CustomKeys.BindedKey(mKeyRight)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(mKeyRight)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(mKeyRight))   ' Agrega la tecla al arraylist
    Else

        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(mKeyRight)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(mKeyRight))   ' Remueve la tecla que teniamos presionada

    End If
    
    Exit Sub

AddMovementToKeysMovementPressedQueue_Err:

    Resume Next
    
End Sub

Public Sub CheckKeys()

    Static LastTick As Long

    '
    'If Not Application.IsAppActive() Then Exit Sub
    
    ' No engines
    If Not EngineRun Then Exit Sub
    
    If MirandoSkins Then Exit Sub

    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'No walking when fmr cantidad is visible
    If MirandoCantidad Then Exit Sub

    If MirandoConcentracion Then Exit Sub
    
    If MirandoCuenta Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Then Exit Sub
    
    'TODO: Deber�a informarle por consola?
    If Traveling Then Exit Sub

    'If FrmViajes.Visible Then Exit Sub
    If MirandoTravel Then Exit Sub
    
    ' If (FrameTime - LastTick >= 0) Then
    'LastTick = FrameTime

    If Not UserMoving Then 'And Not WaitInput
        If ClientSetup.bConfig(eSetupMods.SETUP_MOVERSEHABLAR) = 0 And FrmMain.SendTxt.visible Then Exit Sub
        If FrmMain.TrainingMacro.Enabled Then FrmMain.DesactivarMacroHechizos
                
        If Not UserEstupido Then
                    
            Call AddMovementToKeysMovementPressedQueue
                    
            Select Case keysMovementPressedQueue.GetLastItem()
            
                    ' Prevenimos teclas sin asignar... Te deja moviendo para siempre
                Case 0: Exit Sub
                                
                    'Move Up
                Case CustomKeys.BindedKey(mKeyUp)
                    Call MoveTo(E_Heading.NORTH)
            
                    'Move Right
                Case CustomKeys.BindedKey(mKeyRight)
                    Call MoveTo(E_Heading.EAST)
                                    
                    'Move down
                Case CustomKeys.BindedKey(mKeyDown)
                    Call MoveTo(E_Heading.SOUTH)
                                    
                    'Move left
                Case CustomKeys.BindedKey(mKeyLeft)
                    Call MoveTo(E_Heading.WEST)
                                    
            End Select
                    
        Else
            Call RandomMove

        End If

    End If

    'End If

End Sub

Sub SwitchMap(ByVal Map As Integer)

    '<EhHeader>
    On Error GoTo SwitchMap_Err

    '</EhHeader>

    '**************************************************************
    'Formato de mapas optimizado para reducir el espacio que ocupan.
    'Dise�ado y creado por Juan Mart�n Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
    '**************************************************************
    Dim Y       As Long

    Dim X       As Long

    Dim tempint As Integer

    Dim ByFlags As Byte

    Dim handle  As Integer
    
    handle = FreeFile()
    
    Open App.path & Maps_FilePath & "Mapa" & Map & ".map" For Binary As handle
    Seek handle, 1
            
    'map Header
    Get handle, , MapInfo.MapVersion
    Get handle, , MiCabecera
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    
    g_Swarm.Clear
    
    Dim NullMap As MapBlock

    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(X, Y).SoundSource > 0 Then
                Call Audio.DeleteSource(MapData(X, Y).SoundSource, True)

            End If
            
            ' MapData(X, Y) = NullMap
            
            Get handle, , ByFlags
            
            MapData(X, Y).Blocked = (ByFlags And 1)
            
            Get handle, , MapData(X, Y).Graphic(1).GrhIndex

            If MapData(X, Y).Graphic(1).GrhIndex = 71706 Then
                Debug.Print "lautaro"

            End If
                
            InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get handle, , MapData(X, Y).Graphic(2).GrhIndex
                InitGrh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex

                With GrhData(MapData(X, Y).Graphic(2).GrhIndex)
                    Call g_Swarm.Insert(1, -1, X, Y, .TileWidth, .TileHeight)

                End With

                If MapData(X, Y).Graphic(2).GrhIndex = 71706 Then
                    Debug.Print "lautaro"

                End If

            Else
                MapData(X, Y).Graphic(2).GrhIndex = 0

            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get handle, , MapData(X, Y).Graphic(3).GrhIndex
                InitGrh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex

                With GrhData(MapData(X, Y).Graphic(3).GrhIndex)
                    Call g_Swarm.Insert(2, -1, X, Y, .TileWidth, .TileHeight)
                        
                    If MapData(X, Y).Graphic(3).GrhIndex = 71706 Then
                        Debug.Print "lautaro"

                    End If

                End With

            Else
                MapData(X, Y).Graphic(3).GrhIndex = 0

            End If
                
            ' Layer 4 used?
            If ByFlags And 8 Then
                Get handle, , MapData(X, Y).Graphic(4).GrhIndex
                InitGrh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndex

                With GrhData(MapData(X, Y).Graphic(4).GrhIndex)
                    Call g_Swarm.Insert(3, -1, X, Y, .TileWidth, .TileHeight)
                        
                    If MapData(X, Y).Graphic(4).GrhIndex = 71706 Then
                        Debug.Print "lautaro"

                    End If

                End With

            Else
                MapData(X, Y).Graphic(4).GrhIndex = 0

            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get handle, , MapData(X, Y).Trigger
            Else
                MapData(X, Y).Trigger = 0

            End If
            
            'Erase NPCs
            If MapData(X, Y).CharIndex > 0 Then
                Call EraseChar(MapData(X, Y).CharIndex)

            End If
            
            'Erase OBJs
            MapData(X, Y).ObjGrh.GrhIndex = 0
        Next X
    Next Y
    
    Close handle
    
    ' Sonidos del Mapa ambientales
    Call IMapInitial_Sound(Map)
    
    MapInfo.Name = ""
    MapInfo.Music = ""
    
    CurMap = Map

    '<EhFooter>
    Exit Sub

SwitchMap_Err:
    LogError err.Description & vbCrLf & "in ARGENTUM.Mod_General.SwitchMap " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

'TODO : Si bien nunca estuvo all�, el mapa es algo independiente o a lo sumo dependiente del engine, no va ac�!!!
Sub SwitchMap_Copy(ByVal Map As Integer)

    '**************************************************************
    'Formato de mapas optimizado para reducir el espacio que ocupan.
    'Dise�ado y creado por Juan Mart�n Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
    '**************************************************************
    '<EhHeader>
    On Error GoTo SwitchMap_Copy_Err

    '</EhHeader>
    Dim Y        As Long

    Dim X        As Long

    Dim tempint  As Integer

    Dim ByFlags  As Byte

    Dim hFile    As Integer

    Dim Reader   As Network.Reader

    Dim Buffer() As Byte

    Buffer = LoadBytes(Maps_FilePath & "Mapa" & Map & ".map")

    Set Reader = New Network.Reader
    Call Reader.SetData(Buffer)
 
    'map Header
    Reader.ReadInt16
 
    Call Reader.Skip(255 + 8 + 8) ' MiCabecera + Double
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
 
            With MapData_Copy(X, Y)
 
                ByFlags = Reader.ReadInt8
 
                .Blocked = (ByFlags And 1)
 
                .Graphic(1).GrhIndex = Reader.ReadInt32
                InitGrh .Graphic(1), .Graphic(1).GrhIndex
 
                'Layer 2 used?
                If ByFlags And 2 Then
                    .Graphic(2).GrhIndex = Reader.ReadInt32
                    InitGrh .Graphic(2), .Graphic(2).GrhIndex
                Else
                    .Graphic(2).GrhIndex = 0

                End If
 
                'Layer 3 used?
                If ByFlags And 4 Then
                    .Graphic(3).GrhIndex = Reader.ReadInt32
                    InitGrh .Graphic(3), .Graphic(3).GrhIndex
                Else
                    .Graphic(3).GrhIndex = 0

                End If
 
                'Layer 4 used?
                If ByFlags And 8 Then
                    .Graphic(4).GrhIndex = Reader.ReadInt32
                    InitGrh .Graphic(4), .Graphic(4).GrhIndex
                Else
                    .Graphic(4).GrhIndex = 0

                End If
 
                'Trigger used?
                If ByFlags And 16 Then
                    .Trigger = Reader.ReadInt16
                Else
                    .Trigger = 0

                End If

            End With

        Next X
    Next Y

    '<EhFooter>
    Exit Sub

SwitchMap_Copy_Err:
    LogError err.Description & vbCrLf & "in SwitchMap_Copy " & "at line " & Erl

    '</EhFooter>
End Sub

Function ReadField(ByVal Pos As Integer, _
                   ByRef Text As String, _
                   ByVal SepASCII As Byte) As String

    '*****************************************************************
    'Gets a field from a delimited string
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/15/2004
    '*****************************************************************
    Dim i          As Long

    Dim lastPos    As Long

    Dim CurrentPos As Long

    Dim delimiter  As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)

    End If

End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long

    '*****************************************************************
    'Gets the number of fields in a delimited string
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modify Date: 07/29/2007
    '*****************************************************************
    Dim Count     As Long

    Dim curPos    As Long

    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count

End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(File, FileType) <> "")

End Function

Private Sub Load_Messages()
    ReDim MessagesSpam(1 To 14) As String
    
    MessagesSpam(1) = "Tu personaje comenzar� en mapa principal de los newbies. Al alcanzar el nivel 13 ir�s a la ciudad principal para comenzar un recorrido por los mundos."
    MessagesSpam(2) = "Las pociones iniciales te acompa�aran hasta Nivel 12 inclusive. Tu vestimenta y daga desaparecer�n al Nivel 18 y ser� momento de que ya dispongas de objetos propios y algunas monedas de oro..."
    MessagesSpam(3) = "Tu personaje podr� ser borrado �nicamente siendo menor a Nivel 30. Luego no podr�s quitarlo de tu cuenta por seguridad."
    MessagesSpam(4) = "La primer embarcaci�n que podr�s utilizar para recorrer los mares la usar�s antes del Nivel 25 �Adelantate a la Aventura!"
    MessagesSpam(5) = "Ten cuidado con aquellos Ladrones de Objetos y Oro. Te quitaran todas tus pertenencias en Zona Insegura �Te recomendamos que vayas con Cuidado!"
    MessagesSpam(6) = "Cuanto m�s te alejes de la Ciudad Principal, mas peligro correr�s. Haz clic sobre el MiniMapa para poder ver el Mundo Desterium"
    MessagesSpam(7) = "El comando /HOGAR te llevar� a la Ciudad Principal. A medida que tu personaje sea de mayor nivel requerir� m�s Monedas de Oro. �Ten Cuidado, podr�as quedarte sin oro! Adem�s cuanto m�s lejos de la Ciudad Principal est�s, m�s tiempo tardar�s en llegar. Tampoco podr�s moverte en viaje, pero si podr�s cancelarlo tipeando nuevamente /HOGAR."
    MessagesSpam(8) = "Las Criaturas de Ullathorpe se encargar�n de fabricar los objetos necesarios para tu personaje. Para ello debes apretar el bot�n 'FABRICACION' y previamente debes conseguir los recursos necesarios..."
    MessagesSpam(9) = "Tu personaje 'Trabajador' extraer� una mayor cantidad de recursos a medida que adquiere mayor Nivel. Por eso entrenarlo es una buena manera de destacarse en la obtenci�n de recursos naturales"
    MessagesSpam(10) = "Para fundar un nuevo clan debes obtener las Gemas Veril en sus diferentes variantes y podr�s obtenerlas acabando con varias Medusas Reinas..."
    MessagesSpam(11) = "Para sumar miembros a un clan, debes tipear el comando /CLAN seguido del nombre del usuario al que desees enviar la invitaci�n."
    MessagesSpam(12) = "Puedes formar grupos con los dem�s usuarios del juego y as� dividir la experiencia obtenida. Para ello debes tipear la tecla F7"
    MessagesSpam(13) = "Cuando alcances un nivel considerado podr�s disputar de enfrentamientos privados contra otros personajes. �Podr�s apostar tus objetos!. Para ello debes tipear la tecla F5"
    MessagesSpam(14) = "En los Pasillos de Veril podr�s obtener un poder m�gico proveniente de las medusas y con �l hacer que tu personaje sea m�s fuerte e inmue a la inmovilidad �Usalo para entrenar!"
    
End Sub

Public Sub SelectedSpamMessage()
    'MessagesSpam_Last = MessagesSpam_Last + 1
    
    'Call ShowConsoleMsg(MessagesSpam(MessagesSpam_Last), RandomNumber(100, 200), RandomNumber(200, 250), RandomNumber(200, 250), True, True)
    
    ' If MessagesSpam_Last >= UBound(MessagesSpam) Then
    ' MessagesSpam_Last = 0
    ' End If
End Sub

Sub Main()

    '<EhHeader>
    On Error GoTo Main_Err

    '</EhHeader>
        
    Dim TimeUpdate As Long
        
    Folder_Renew MiniMap_FilePath
    ' StartMonitoring
   
    frmCargando.Show
        
    'Call InitCommonControls
    Call Commands_Load
    'frmCargando.imgLoading.Width = 0
          
    frmCargando.lblLoad.Caption = "Cargando, Espera..."

    'Esto es para el movimiento suave de pjs, para que el pj termine de hacer el movimiento antes de empezar otro
    Set keysMovementPressedQueue = New clsArrayList
    Call keysMovementPressedQueue.Initialize(1, 4)
    
    Dim Temp As String

    With AccountSec
        .IP_Local = Application.System_GetIP_Local
        .SERIAL_BIOS = Application.System_GetSerial_BIOS
        .SERIAL_DISK = Application.System_GetSerial_DISK
        .SERIAL_MAC = Application.System_GetMAC
        .SERIAL_MOTHERBOARD = Application.System_GetSerial_Motherboard
        .SERIAL_PROCESSOR = Application.System_GetSerial_Processor
        .SYSTEM_DATA = Application.System_GetData
        .IP_Public = IP_Publica

    End With
        
    'Lorwik> Busca Cheats
    Call SearchDesterium
 
    'Lorwik>
    'Call ILoad_Temporal
    
    Call mParticle.Initialize
    Call Load_Messages
    
    Call LoadListPasswd
    'frmCargando.Refresh
    IniPath = Init_FilePath
    
    '#If Testeo = 0 Then
    'If Get_ValidExecute = "0" Then
    'Call MsgBox("Debes ejecutar el juego desde el Launcher, para estar siempre actualizado.")
    ' Exit Sub
    ' End If
    '#End If
    
    'Call Set_ValidExecute
    
    Call GenerateContra("Gracia$Totales")
    Call Initialize_Security
    Call SetIntervalos

    'UserIpExternal = IP_Publica
    
    'Load config file
    If FileExist(IniPath & "Inicio.con", vbNormal) Then
        Config_Inicio = LeerGameIni()
        Config_Inicio.DirMusica = "MP3"

    End If
  
    If FileExist(Init_FilePath & CustomPath, vbArchive) Then LoadCustomConsole

    Call ILoadClientSetup

    #If Testeo = 0 Then

'        If FindPreviousInstance Then
'            MsgBox "No puedes tenerm m�s de 1 cliente abierto."
'            End
'
'        End If

    #End If
    
    'usaremos esto para ayudar en los parches
    'Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.path & "\")
    
    'ChDrive App.path
    ' ChDir App.path

    MD5HushYo = "0123456789abcdef"  'We aren't using a real MD5

    'Set resolution BEFORE the loading form is displayed, therefore it will be centered
        
    Call Resolution.SetResolution
        
    ' Load constants, classes, flags, graphics..
    LoadInitialConfig
 
    frmCargando.visible = False
 
    frmConnectAcc.Show

    'Inicializaci�n de variables globales
    prgRun = True
    pausa = False

    lFrameTimer = FrameTime
    'Call SetElapsedTime(True)
    
    Do While prgRun
        
        Call RenderStarted
        
        ' If there is anything to be sent, we send it and poll all network events
        Call modNetwork.Poll

        ' Update audio thread
        Call Audio.SetListener(UserPos.X, UserPos.Y)
        Call Audio.Update(FrameTime)

    Loop
    
    Call CloseClient
    '<EhFooter>
    Exit Sub

Main_Err:
    LogError err.Description & vbCrLf & "in ARGENTUM.Mod_General.Main " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Private Sub LoadEngine()

    '<EhHeader>
    On Error GoTo LoadEngine_Err

    '</EhHeader>

    Dim pixelWidth As Long, pixelHeight As Long

    Dim TileAlto   As Long, TileAncho As Long

    Dim ScrollX    As Single, ScrollY As Single

    Dim Speed      As Single
    
    #If ModoBig = 1 Then
        pixelWidth = 64
        pixelHeight = 64
        TileAlto = 13
        TileAncho = 17
        ScrollX = 11.5
        ScrollY = 11.5
        Speed = 0.025
        
    #Else
        pixelWidth = 32
        pixelHeight = 32
        TileAlto = 13
        TileAncho = 17
        ScrollX = 8.5
        ScrollY = 8.5
        Speed = 0.018
                
    #End If

    #If FullScreen = 1 Then
        TileAlto = 17
        TileAncho = 25
    #End If

    If Not InitTileEngine(pixelHeight, pixelWidth, TileAlto, TileAncho, ScrollX, ScrollY, Speed) Then
        Call CloseClient

    End If

    '<EhFooter>
    Exit Sub

LoadEngine_Err:
    LogError err.Description & vbCrLf & "in ARGENTUM.Mod_General.LoadEngine " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Private Sub LoadInitialConfig()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/03/2011
    '15/03/2011: ZaMa - Initialize classes lazy way.
    '***************************************************
    '<EhHeader>
    On Error GoTo LoadInitialConfig_Err

    '</EhHeader>

    Dim i As Long
        
    'frmCargando.imgLoading.Width = 0
    'frmCargando.lblLoad.Caption = "Cargando Juego..."
    '###########
    ' CONSTANTES
    ' frmCargando.imgLoading.Width = 50
    'frmCargando.lblLoad.Caption = "Cargando Textos Predeterminados..."
    'Call AddtoRichTextBox(frmCargando.Status, "Cargando textos predeterminados... ", 255, 255, 255, True, False, True)
    Call InicializarNombres
    
    ' Initialize FONTTYPES
    Call Protocol.InitFonts
    
    UserMap = 1
    
    ' Mouse Pointer (Loaded before opening any form with buttons in it)
    If FileExist(DirExtras & "Hand.ico", vbArchive) Then Set picMouseIcon = LoadPicture(DirExtras & "Hand.ico")

    'Call AddtoRichTextBox(frmCargando.Status, "Hecho", 50, 255, 0, True, False, False)

    '#######
    ' CLASES
    'frmCargando.imgLoading.Width = 80
    'frmCargando.lblLoad.Caption = "Iniciando clases predeterminadas..."
    'Call AddtoRichTextBox(frmCargando.Status, "Iniciando clases predeterminadas... ", 255, 255, 255, True, False, True)
    Set Dialogos = New clsDialogs
    Set DialogosClanes = New clsGuildDlg
    Set Audio = New clsAudio
    Set CustomKeys = New clsCustomKeys
    Set CustomMessages = New clsCustomMessages
    Set MainTimer = New clsTimer
          
    DialogosClanes.Activo = True
          
    'Call AddtoRichTextBox(frmCargando.Status, "Hecho", 50, 255, 0, True, False, False)

    '##############
    ' MOTOR GR�FICO
    ' frmCargando.imgLoading.Width = 125
    ' frmCargando.lblLoad.Caption = "Iniciando motor gr�fico..."
    'Call AddtoRichTextBox(frmCargando.Status, "Iniciando motor gr�fico... ", 255, 255, 255, True, False, True)
 
    Call LoadEngine
        
    'Call AddtoRichTextBox(frmCargando.Status, "Hecho", 50, 255, 0, True, False, False)
    
    '##############
    ' NETWORKING
    'frmCargando.imgLoading.Width = 170
    ' frmCargando.lblLoad.Caption = "Iniciando conexiones..."
    'Call AddtoRichTextBox(frmCargando.Status, "Iniciando conexiones... ", 255, 255, 255, True, False, True)
    Call modNetwork.Initialise
    'Call AddtoRichTextBox(frmCargando.Status, "Hecho", 50, 255, 0, True, False, False)
        
    ' Carga los Mapas tEMPORALMENTE
    Dim Temp As Long

    ' frmCargando.imgLoading.Width = 202
    ' frmCargando.lblLoad.Caption = "Cargando animaciones..."
        
    'Call AddtoRichTextBox(frmCargando.Status, "Cargando animaciones... ", 255, 255, 255, True, False, True)
    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCuerposAttack
    Call CargarAuras
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    Call DataServer_LoadAll
    Call mMaps.ILoad_MapInfo                        ' @ Carga todos los sonidos en los mapas.
    Call DB_LoadSkills                                    ' @ Carga la info de los skills y skills especiales
    Call DataServer_Load_Maps                     ' @ Info de los Mapas
    Call DataServer_Load_Spells                   ' @ Datos de los hechizos
    Call Drops_Load                                       ' @ Drops
    Call Ruleta_LoadItems
    ' ###################
    ' ANIMACIONES EXTRAS
    
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    Call CargarDialogos
    Call Load_Balance
    'Call AddtoRichTextBox(frmCargando.Status, "Hecho", 50, 255, 0, True, False, False)

    '#############
    ' DIRECT SOUND
    '   frmCargando.imgLoading.Width = 250
    '  frmCargando.lblLoad.Caption = "Iniciando motor de audio..."
    'Call AddtoRichTextBox(frmCargando.Status, "Iniciando motor de audio... ", 255, 255, 255, True, False, True)

    'Inicializamos el sonido
    Call Audio.Initialize(DirMusic, DirSound)
             
    'Enable / Disable audio
    Audio.MusicActivated = (ClientSetup.bSoundMusic > 0)
    Audio.EffectActivated = (ClientSetup.bSoundEffect > 0)
    Audio.InterfaceActivated = (ClientSetup.bSoundInterface > 0)
    Audio.MasterActivated = (ClientSetup.bMasterSound > 0)
          
    Audio.MasterVolume = ClientSetup.bValueSoundMaster
    Audio.MusicVolume = ClientSetup.bValueSoundMusic
    Audio.EffectVolume = ClientSetup.bValueSoundEffect
    Audio.InterfaceVolume = ClientSetup.bValueSoundInterface
    Audio.Effect3D = ClientSetup.bConfig(eSetupMods.SETUP_SOUND3D)
        
    'Inicializamos el inventario gr�fico

    '#If Testeo = 0 Then
    Call Audio.PlayMusic(MP3_Inicio & ".mp3")
    '#End If
        
    'Call AddtoRichTextBox(frmCargando.Status, "Hecho", 50, 255, 0, True, False, False)
    
    ' frmCargando.imgLoading.Width = 307
    '  frmCargando.lblLoad.Caption = "�Comenzando Nueva Aventura!"
    'Call AddtoRichTextBox(frmCargando.Status, "        �Bienvenido a DS Exodo III!", 200, 255, 200, True, False, True)
       
    'Give the user enough time to read the welcome text
    ' Call Sleep(1000)
    
    ' Unload frmCargando
    
    '<EhFooter>
    Exit Sub

LoadInitialConfig_Err:
    LogError err.Description & vbCrLf & "in LoadInitialConfig " & "at line " & Erl

    '</EhFooter>
End Sub

Sub WriteVar(ByVal File As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal Value As String)
    '*****************************************************************
    'Writes a var to a text file
    '*****************************************************************
    writeprivateprofilestring Main, Var, Value, File

End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String

    '*****************************************************************
    'Gets a Var from a text file
    '*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function

'[CODE 002]:MatuX
'
'  Funci�n para chequear el email
'
'  Corregida por Maraxus para que reconozca como v�lidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean

    On Error GoTo errHnd

    Dim lPos As Long

    Dim lX   As Long

    Dim iAsc As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")

    If (lPos <> 0) Then

        '2do test: Busca un simbolo . despu�s de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then Exit Function
        
        '3er test: Recorre todos los caracteres y los val�da
        For lX = 0 To Len(sString) - 1

            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))

                If Not CMSValidateChar_(iAsc) Then Exit Function

            End If

        Next lX
        
        'Finale
        CheckMailString = True

    End If

errHnd:

End Function

'  Corregida por Maraxus para que reconozca como v�lidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 90) Or (iAsc >= 97 And iAsc <= 122) Or (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)

End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer ac�....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = (((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520) Or (MapData(X, Y).Graphic(1).GrhIndex >= 5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or (MapData(X, Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= 13562)) Or (MapData(X, Y).Graphic(1).GrhIndex >= 39568 And MapData(X, Y).Graphic(1).GrhIndex <= 39583) Or (MapData(X, Y).Graphic(1).GrhIndex >= 39584 And MapData(X, Y).Graphic(1).GrhIndex <= 39599) Or (MapData(X, Y).Graphic(1).GrhIndex >= 39600 And MapData(X, Y).Graphic(1).GrhIndex <= 39615)) And MapData(X, Y).Graphic(2).GrhIndex = 0
                
End Function

Public Sub ShowSendTxt()

    If Not FrmCantidad.visible Then
        FrmMain.SendTxt.visible = True
        FrmMain.SendTxt.SetFocus

    End If

End Sub

Private Sub InicializarNombres()
    '**************************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/27/2005
    'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
    '**************************************************************
    Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
    Ciudades(eCiudad.cNix) = "Nix"
    Ciudades(eCiudad.cBanderbill) = "Banderbill"
    Ciudades(eCiudad.cLindos) = "Lindos"
    Ciudades(eCiudad.cArghal) = "Argh�l"
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    
    ListaRazasShort(eRaza.Humano) = "H"
    ListaRazasShort(eRaza.Elfo) = "Elf"
    ListaRazasShort(eRaza.ElfoOscuro) = "Eo"
    ListaRazasShort(eRaza.Gnomo) = "G"
    ListaRazasShort(eRaza.Enano) = "E"
    
    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Thief) = "Ladron"
    
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasi�n en combate"
    SkillsNames(eSkill.Armas) = "Combate con Armas"
    SkillsNames(eSkill.Apu�alar) = "Apu�alar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Talar) = "Talar �rboles"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Comercio) = "Comercio"
    SkillsNames(eSkill.Proyectiles) = "Armas con Proyectiles"
    SkillsNames(eSkill.Navegacion) = "Navegacion"
    SkillsNames(eSkill.Resistencia) = "Resistencia m�gica"
    SkillsNames(eSkill.Domar) = "Domar Animales"
    
    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"

End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
    '**************************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/27/2005
    'Removes all text from the console and dialogs
    '**************************************************************
    'Clean console and dialogs
    FrmMain.RecTxt.Text = vbNullString
    
    Call DialogosClanes.RemoveDialogs
    
    Call Dialogos.RemoveAllDialogs

End Sub

Public Sub CloseClient()
    '**************************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modify Date: 8/14/2007
    'Frees all used resources, cleans up and leaves
    '**************************************************************
    
    ' Quitamos los Cursores gr�ficos
    Call Cursores_ResotreDefault
    
    ' Allow new instances of the client to be opened
    Call PrevInstance.ReleaseInstance
    
    EngineRun = False
    'frmCargando.Show
    'frmCargando.imgLoading.Width = 0
    'frmCargando.lblLoad.Caption = "Liberando Recursos..."
    'Call AddtoRichTextBox(frmCargando.Status, "Liberando recursos...", 0, 0, 0, 0, 0, 0)
    
    Call Resolution.ResetResolution
    
    'Stop tile engine
    Call DeinitTileEngine
    
    Call UnloadAllForms
    
    'Actualizar tip
    Call EscribirGameIni(Config_Inicio)
    
    'Call RemoveFont(App.path & "\AO\resource\FONT\" & "booterfz.ttf")
    
    'Destruimos los objetos p�blicos creados
    Set CustomMessages = Nothing
    Set CustomKeys = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    
    ' StopMonitoring
    
    End

End Sub

Public Function esGM(CharIndex As Integer) As Boolean
    esGM = False

    If CharList(CharIndex).Priv >= 1 And CharList(CharIndex).Priv < 5 Or CharList(CharIndex).Priv = 25 Then esGM = True

End Function

Public Function esAdmin(CharIndex As Integer) As Boolean
    esAdmin = False

    If CharList(CharIndex).Priv = 3 Or CharList(CharIndex).Priv = 4 Then esAdmin = True

End Function

Public Function getTagPosition(ByVal Nick As String) As Integer

    Dim buf As Integer

    buf = InStr(Nick, "<")

    If buf > 0 Then
        getTagPosition = buf

        Exit Function

    End If

    buf = InStr(Nick, "[")

    If buf > 0 Then
        getTagPosition = buf

        Exit Function

    End If

    getTagPosition = Len(Nick) + 2

End Function

Public Function getStrenghtColor(ByVal yFuerza As Byte) As Long

    Dim m As Long

    m = Int(255 / MAXATRIBUTOS)
    getStrenghtColor = RGB(255 - (m * yFuerza), (m * yFuerza), 0)

End Function

Public Function getDexterityColor(ByVal yAgilidad As Byte) As Long

    Dim m As Long

    m = 255 / MAXATRIBUTOS
    getDexterityColor = RGB(255, m * yAgilidad, 0)

End Function

Public Function getCharIndexByName(ByVal Name As String) As Integer

    Dim i As Long

    For i = 1 To LastChar

        If CharList(i).Nombre = Name Then
            getCharIndexByName = i

            Exit Function

        End If

    Next i

End Function

Public Sub ResetAllInfo(Optional ByVal ResetAccount As Boolean)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/06/2011
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ResetAllInfo_Err

    '</EhHeader>

    ' Save config.ini
    SaveGameini
        
    ' Disable timers
    FrmMain.Second.Enabled = False
    FrmMain.tUpdateMS.Enabled = False
    FrmMain.tUpdateInactive.Enabled = False
    FrmMain.tmrBlink.Enabled = False
    FrmMain.MacroTrabajo.Enabled = False
          
    FrmMain.lblDopa.visible = False
    FrmMain.imgDopa.visible = False
    FrmMain.lblInvi.visible = False
    FrmMain.imgInvisible.visible = False
    FrmMain.imgHome.visible = False
    FrmMain.lblHome.visible = False
    FrmMain.imgParalisis.visible = False
    FrmMain.lblParalisis.visible = False
         
    If ResetAccount Then
        MirandoCuenta = False
        Account = NullAccount
        FrmMain.tUpdateInactive.Enabled = False

    Else
        MirandoCuenta = True

    End If
    
    Connected = False
    
    'Unload all forms except frmMain, frmConnect
    Dim frm As Form

    For Each frm In Forms

        If frm.Name <> FrmMain.Name And frm.Name <> FrmConnect_Account.Name And frm.Name <> FrmConectando.Name Then
            frm.Hide

        End If

    Next
    
    ' Eliminamos el sonido que est� en la Interfaz en caso de existir
    Call Audio.DeleteSource(SOURCE_INTERFACE, True)
    
    On Local Error GoTo 0
    
    ' Return to connection screen
    'frmConnect.MousePointer = vbNormal
    'frmconnect.Visible=
    
    If Not ResetAccount Then
              
        FrmConnect_Account.Show
        FrmConnect_Account.SelectedPanelAccount (ePanelAccount)

    End If
    
    'If Not frmCrearPersonaje.Visible Then frmConnect.Visible = True
    FrmMain.visible = False
    
    'Stop audio
    If (ResetAccount) Then
        Call Audio.Halt

    End If
    
    FrmMain.IsPlaying = 0
        
    ' Reset flags
    pausa = False
    UserMeditar = False
    UserEstupido = False
    UserCiego = False
    UserDescansar = False
    UserParalizado = False
    UserEnvenenado = False
    UserLeader = False
    Traveling = False
    UserNavegando = False
    UserMontando = False
    bFogata = False
    Comerciando = False
    UserEvento = False
    UserMontando = False
    
    MirandoEstadisticas = False
    MirandoCantidad = False
    MirandoForo = False
    MirandoParty = False
    
    MirandoRank = False
    MirandoGuildPanel = False
    MirandoTravel = False
    MirandoComerciarUsu = False
    MirandoBanco = False
    MirandoConcentracion = False
    MirandoComerciar = False
    MirandoSkins = False
    MirandoObjetos = False
    MirandoOpcionesNpc = False
    MirandoPartidas = False
        
    MirandoMercader = False
    MirandoStatsUser = False
    MirandoOffer = False
        
    MirandoObj = False
    MirandoNpc = False
        
    UserHelmEqpSlot = 0
    UserMagicEqpSlot = 0
    UserArmourEqpSlot = 0
    UserShieldEqpSlot = 0
    UserAnilloEqpSlot = 0
    UserWeaponEqpSlot = 0
        
    'Delete all kind of dialogs
    Call CleanDialogs

    'Reset some char variables...
    Dim i As Long

    For i = 1 To LastChar
        CharList(i).Invisible = False
        CharList(i).Speeding = 0
    Next i

    ' Reset stats
    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    SkillPoints = 0
    Alocados = 0
    UserFaccion = 0
    
    ' Reset skills
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    ' Reset attributes
    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    
    ' Clear inventory slots
    Inventario.ClearAllSlots
    
    ResetKeyPackets
        
    '<EhFooter>
    Exit Sub

ResetAllInfo_Err:
    LogError err.Description & vbCrLf & "in ARGENTUM.Mod_General.ResetAllInfo " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Function complexNameToSimple(ByVal str As String, ByVal isGuild As Boolean) As String

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 06/12/2011
    '
    '***************************************************
    Dim i   As Long

    Dim aux As String

    For i = 1 To Len(CAR_ESPECIALES)
        aux = mid$(CAR_ESPECIALES, i, 1)

        If InStr(1, str, aux) Then
            str = Replace(str, aux, mid$(CAR_COMUNES, i, 1))

        End If

    Next i
    
    If isGuild Then
    
        For i = 1 To Len(CAR_ESPECIALES_CLANES)
            aux = mid$(CAR_ESPECIALES_CLANES, i, 1)

            If InStr(1, str, aux) Then
                str = Replace(str, aux, mid$(CAR_COMUNES_CLANES, i, 1))

            End If

        Next i

    End If
    
    complexNameToSimple = str

End Function

Public Sub LoadCustomConsole()
    '***************************************************
    'Author: D'Artagnan (aop.fran@gmail.com)
    'Last Modification: 01/27/2012
    'Load user's custom console config.
    '***************************************************
    AdminMsg = Val(GetVar(Init_FilePath & CustomPath, "CONFIG", "Admin"))
    GuildMsg = Val(GetVar(Init_FilePath & CustomPath, "CONFIG", "Clan"))
    PartyMsg = Val(GetVar(Init_FilePath & CustomPath, "CONFIG", "Party"))
    CombateMsg = Val(GetVar(Init_FilePath & CustomPath, "CONFIG", "Combate"))
    TrabajoMsg = Val(GetVar(Init_FilePath & CustomPath, "CONFIG", "Trabajo"))
    InfoMsg = Val(GetVar(Init_FilePath & CustomPath, "CONFIG", "Info"))

End Sub

Public Sub Invalidate(ByVal hWnd As Long)
 
    Call RedrawWindow(hWnd, &H0, &H0, &H1)
 
End Sub

Public Sub ResetKeyPackets()

    Dim A As Byte
    
    For A = 0 To MAX_KEY_PACKETS
        KeyPackets(A) = 0
    Next A

End Sub

Public Function CurServerIp() As String

    Dim filePath As String

    filePath = IniPath & "Sinfo.DAT"

    CurServerIp = GetVar(filePath, "INIT", "IP")

End Function

Public Function CurServerPort() As String

    Dim filePath As String

    Dim Temp     As Long
    
    filePath = IniPath & "Sinfo.DAT"

    Temp = Val(GetVar(filePath, "INIT", "PORT"))
    
    CurServerPort = Temp

End Function

Public Function CharTieneClan() As Boolean

    Dim tPos As Integer

    tPos = InStr(CharList(UserCharIndex).Nombre, "<")

    If tPos = 0 Then
        CharTieneClan = False

        Exit Function

    End If

    CharTieneClan = True

End Function

Public Function GetExternalIP() As String

    Dim ArchN As Integer
    
    ArchN = FreeFile()
    Open IniPath & "IP.ini" For Input As #ArchN
         
    Do While Not EOF(ArchN)
        Line Input #ArchN, GetExternalIP
    Loop
         
    Close #ArchN

End Function

Public Sub Set_ValidExecute()

    Dim ArchN As Integer
    
    ArchN = FreeFile()
    Open IniPath & "Execute.ini" For Output As #ArchN

    Write #ArchN, 0
         
    Close #ArchN

End Sub

Public Function Get_ValidExecute() As String

    Dim ArchN As Integer
    
    ArchN = FreeFile()
    Open IniPath & "Execute.ini" For Input As #ArchN
         
    Do While Not EOF(ArchN)
        Line Input #ArchN, Get_ValidExecute
    Loop
         
    Close #ArchN

End Function

Public Sub ScreenCapture(Optional ByVal Autofragshooter As Boolean = False, _
                         Optional ByVal boton As Boolean = False)

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 11/16/2006
    '11/16/2010: Amraphen - Now the FragShooter screenshots are stored in different directories.
    '**************************************************************
    On Error GoTo err:

    Dim File    As String

    Dim sI      As String

    Dim i       As Long
    
    Dim dirFile As String
    
    ' Primero chequea si existe la carpeta Screenshots
    dirFile = App.path & "\Screenshots"

    If Not FileExist(dirFile, vbDirectory) Then Call MkDir(dirFile)

    'Si no es screenshot del FragShooter, entonces se usa el formato "DD-MM-YYYY hh-mm-ss.jpg"
    File = dirFile & "\" & Format(Now, "DD-MM-YYYY hh-mm-ss") & ".jpg"
    
    LastCapture = Format(Now, "DD-MM-YYYY hh-mm-ss") & ".jpg"
    
    Call wGL_Graphic.Capture(FrmMain.hWnd, File)
    
    If boton Then
        AddtoRichTextBox FrmMain.RecTxt, "�Screen capturada correctamente!", 200, 200, 200, False, False, True

    End If
    
    Exit Sub

err:
    Call AddtoRichTextBox(FrmMain.RecTxt, err.Number & "-" & err.Description, 200, 200, 200, False, False, True)

End Sub

Public Sub General_Drop_X_Y(ByVal X As Byte, ByVal Y As Byte)

    ' /  Author  : Dunkan
    ' /  Note    : Calcular la posici�n de donde va a tirar el item

    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
              
        If Inventario.SelectedItem < 1 Then
            Call ShowConsoleMsg("No tienes esa cantidad.", 65, 190, 156, False, False)
            Inventario.sMoveItem = False
            Inventario.uMoveItem = False
            FrmMain.PicInv.MouseIcon = Nothing

            Exit Sub

        End If
              
        ' - Hay que pasar estas funciones al servidor
        If MapData(X, Y).Blocked = 1 Then
            Call ShowConsoleMsg("Elige una posici�n v�lida para tirar tus objetos.", 65, 190, 156, False, False)
            Inventario.sMoveItem = False
            FrmMain.PicInv.MouseIcon = Nothing

            Exit Sub

        End If

        If UserEstado = 1 Then
            Call ShowConsoleMsg("�Est�s muerto!", 65, 190, 156, False, False)
            Inventario.sMoveItem = False
            FrmMain.PicInv.MouseIcon = Nothing

            Exit Sub

        End If
            
        If UserMontando = True Then
            Call ShowConsoleMsg("Debes bajarte de la montura para poder arrojar items.", 65, 190, 156, False, False)
            Inventario.sMoveItem = False
            FrmMain.PicInv.MouseIcon = Nothing

            Exit Sub

        End If
                
        If Comerciando Then
            Call ShowConsoleMsg("No puedes arrojar items mientras te encuentres comerciando.", 65, 190, 156, False, False)
            Inventario.sMoveItem = False
            Inventario.uMoveItem = False
            FrmMain.PicInv.MouseIcon = Nothing

            Exit Sub

        End If
                
        ' - Hay que pasar estas funciones al servidor
              
        If GetKeyState(vbKeyShift) < 0 Then
        
            If Not FrmCantidad.visible Then
                
                FrmCantidad.Show , FrmMain
                Call FrmCantidad.SetDropDragged(FrmMain.MouseX, FrmMain.MouseY)

            End If

        Else
            Call WriteDragToPos(X, Y, Inventario.SelectedItem, 1)

        End If

    End If
           
    Inventario.sMoveItem = False
    FrmMain.PicInv.MouseIcon = Nothing

End Sub

Public Sub Guilds_FounderNew()

    Dim A As Long
  
End Sub

Function AsciiValidos_Name(ByVal cad As String) As Boolean

    Dim car As Byte

    Dim i   As Long

    Dim j   As Long
    
    cad = LCase$(cad)

    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If Not ((car >= 97 And car <= 122)) And Not (car = 32) Then
            AsciiValidos_Name = False
            Exit Function

        End If

    Next i
    
    AsciiValidos_Name = True
    
End Function

Public Function Guilds_PrepareRangeName(ByVal Range As eGuildRange) As String

    Select Case Range
    
        Case eGuildRange.rNone
            Guilds_PrepareRangeName = "Miembro"
            
        Case eGuildRange.rFound
            Guilds_PrepareRangeName = "Fundador"
            
        Case eGuildRange.rLeader
            Guilds_PrepareRangeName = "Lider"
            
        Case eGuildRange.rVocero
            Guilds_PrepareRangeName = "Vocero"

    End Select

End Function

Public Function Guilds_SearchIndex(ByVal GuildName As String)

    Dim A As Long
    
    For A = 1 To MAX_GUILDS

        If StrComp(UCase$(GuildsInfo(A).Name), UCase$(GuildName)) = 0 Then
            Guilds_SearchIndex = GuildsInfo(A).Index
            Exit For
        
        End If

    Next A

End Function

Private Function SearchFreeChar() As Byte

    Dim A As Long
    
    With Account
    
        For A = 1 To ACCOUNT_MAX_CHARS

            If .Chars(A).Name = vbNullString Then
                SearchFreeChar = A
                Exit For

            End If

        Next A
        
    End With
    
End Function

Public Function SearchSlotChar(ByVal UserName As String) As Byte

    Dim A As Long
    
    With Account
    
        For A = 1 To ACCOUNT_MAX_CHARS

            If StrComp(UCase$(.Chars(A).Name), UserName) = 0 Then
                SearchSlotChar = A
                Exit For

            End If

        Next A
        
    End With
    
End Function

Public Function SearchFX_Default(ByVal Text As String) As Integer

    Select Case Text
            'Case "Remover Paralisis"
            'SearchFX_Default = FX_REMOVER

            'Case "Inmovilizar"
            'SearchFX_Default = FX_INMOVILIZAR
            
        Case "Tormenta de Fuego"
            SearchFX_Default = FX_TORMENTA
            
        Case "Descarga electrica"
            SearchFX_Default = FX_DESCARGA
            
        Case "Apocalipsis"
            SearchFX_Default = FX_APOCALIPSIS
            
            'Case "Warp"
            'SearchFX_Default = FX_WARP
            
    End Select

End Function

Public Function AscU(ByRef Text As String) As Long

    Dim lngChar As Long, lngChar2 As Long, lngLen As Long

    lngLen = LenB(Text)

    If lngLen Then
        If lngLen <= 2 Then
            lngChar = AscW(Left$(Text, 1)) And &HFFFF&

            If lngChar < &HD800& Or lngChar > &HDBFF& Then
                AscU = lngChar
                Exit Function

            End If

        Else
            lngChar = AscW(Left$(Text, 1)) And &HFFFF&

            If lngChar < &HD800& Or lngChar > &HDBFF& Then
                AscU = lngChar
                Exit Function
            Else
                lngChar2 = AscW(mid$(Text, 2, 1)) And &HFFFF&

                If lngChar2 >= &HDC00& And lngChar2 <= &HDFFF& Then
                    AscU = &H10000 + (((lngChar And &H3FF&) * 1024&) Or (lngChar2 And &H3FF&))
                    Exit Function

                End If

            End If

        End If

    End If

End Function

Public Function SearchObjIndex(ByVal Slot As Byte) As Integer

    Dim A As Long
    
    For A = LBound(ObjBlacksmith) To UBound(ObjBlacksmith)

        If Not ObjBlacksmith(A).ObjIndex = BlacksmithInv.ObjIndex(Slot) Then
            SearchObjIndex = A
            Exit Function

        End If

    Next A

End Function

Public Sub LogError(Desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.path & "\errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

' Transforma los segundos en un tiempo determinado en (Horas Minutos & Segundos)
Public Function SecondsToHMS(ByVal Seconds As Long) As String

    '<EhHeader>
    On Error GoTo SecondsToHMS_Err

    '</EhHeader>

    Dim HR As Integer
    
    Dim MS As Integer
    
    Dim SS As Integer
        
    Dim DS As Integer
        
    DS = (Seconds \ 3600) \ 24
        
    HR = (Seconds \ 3600) Mod 24
    
    MS = (Seconds Mod 3600) \ 60
    
    SS = (Seconds Mod 3600) Mod 60
        
    SecondsToHMS = IIf(DS > 0, DS & " d�as ", vbNullString) & IIf(HR > 0, HR & " hs ", vbNullString) & IIf(MS > 0, MS & " ms ", vbNullString) & IIf(SS > 0, SS & " ss. ", vbNullString)

    '<EhFooter>
    Exit Function

SecondsToHMS_Err:
    LogError err.Description & vbCrLf & "in SecondsToHMS " & "at line " & Erl

    '</EhFooter>
End Function

' Centra el formulario
Sub CentrarFormulario(Formulario As Form, FormularioPadre As Form)
    Formulario.Left = (FormularioPadre.ScaleWidth - Formulario.Width) / 2
    Formulario.Top = (FormularioPadre.ScaleHeight - Formulario.Height) / 2

End Sub

Public Function Porcentaje_String(ByVal Probability As Byte) As String

    Select Case Probability

        Case 1
            Porcentaje_String = "10%"

        Case 2
            Porcentaje_String = "1%"

        Case 3
            Porcentaje_String = "0.1%"

        Case 4
            Porcentaje_String = "0.01%"

        Case 5
            Porcentaje_String = "0.001%"

    End Select

End Function

Public Sub SelectedObjIndex_Update()
    
    Exit Sub

    Dim ParentForm As Form
    
    Char_InfoObj.Head = CharList(UserCharIndex).iHead
    
    ' La inicializaci�n es cuando los valores son nulos, luego queda guardado el �ltimo cargado...
    If Char_InfoObj.Body = 0 Then Char_InfoObj.Body = CharList(UserCharIndex).iBody
    If Char_InfoObj.Helm.Head(1).GrhIndex = 0 Then Char_InfoObj.Helm = CharList(UserCharIndex).Casco
    If Char_InfoObj.Shield.ShieldWalk(1).GrhIndex = 0 Then Char_InfoObj.Shield = CharList(UserCharIndex).Escudo
    If Char_InfoObj.Weapon.WeaponWalk(1).GrhIndex = 0 Then Char_InfoObj.Weapon = CharList(UserCharIndex).Arma
    
    Select Case ObjData(SelectedObjIndex).ObjType
        
        Case eOBJType.otarmadura

            If (UserRaza = Gnomo Or UserRaza = Enano) And ObjData(SelectedObjIndex).AnimBajos > 0 Then
                Char_InfoObj.Body = ObjData(SelectedObjIndex).AnimBajos
            Else
                Char_InfoObj.Body = ObjData(SelectedObjIndex).Anim

            End If
        
        Case eOBJType.otWeapon
            Char_InfoObj.Weapon = WeaponAnimData(ObjData(SelectedObjIndex).Anim)
           
        Case eOBJType.otcasco
            Char_InfoObj.Helm = CascoAnimData(ObjData(SelectedObjIndex).Anim)
            
        Case eOBJType.otescudo
            Char_InfoObj.Shield = ShieldAnimData(ObjData(SelectedObjIndex).Anim)

    End Select
    
    If Not MirandoObjetos Then
            
        If MirandoComerciar Then
            FrmObject_Info.Show , frmComerciar
        ElseIf MirandoBanco Then
            FrmObject_Info.Show , frmBancoObj
        Else
            FrmObject_Info.Show , FrmMain

        End If

    Else
        
        Call FrmObject_Info.Close_Form
        Call FrmObject_Info.Initial_Form

        ' FrmObject_Info.Move (frmComerciar.Left)
    End If
    
End Sub

Public Sub Map_UpdateLabel(Optional ByVal Blocked As Boolean = False)
    
    If FrmMain.CoordBloqued Or Blocked Then
        FrmMain.lblMap(0).Caption = "Mapa " & UserMap & " [" & UserPos.X & "," & UserPos.Y & "]"
    Else
        FrmMain.lblMap(0).Caption = UserMapName

    End If
    
End Sub

Function TieneObjetos(ByVal ItemIndex As Integer) As Long

    Dim A     As Integer

    Dim Total As Long

    For A = 1 To Inventario.MaxObjs

        If Inventario.ObjIndex(A) = ItemIndex Then
            Total = Total + Inventario.Amount(A)

        End If

    Next A
    
    '  If Cant <= Total Then
    TieneObjetos = Total

    'Exit Function

    '  End If
        
End Function

Private Sub Form_CargateControls()

    Dim ThisControl As Control
  
    ReDim Preserve ControlInfos(0 To 0)
    ControlInfos(0).Width = frmMain_Scalled.Width
    ControlInfos(0).Height = frmMain_Scalled.Height

    For Each ThisControl In frmMain_Scalled.Controls

        ReDim Preserve ControlInfos(0 To UBound(ControlInfos) + 1)

        On Error Resume Next  ' hack to bypass controls with no size or position properties

        With ControlInfos(UBound(ControlInfos))
            .Left = ThisControl.Left
            .Top = ThisControl.Top
            .Width = ThisControl.Width
            .Height = ThisControl.Height
            .FontSize = ThisControl.FontSize

        End With

        On Error GoTo 0

    Next

End Sub

Private Sub Form_Resize()

    Dim ThisControl As Control, Iter As Integer

    Iter = 0

    For Each ThisControl In FrmMain.Controls

        Iter = Iter + 1

        On Error Resume Next  ' hack to bypass controls

        With ThisControl
            .Left = ControlInfos(Iter).Left
            .Top = ControlInfos(Iter).Top
            .Width = ControlInfos(Iter).Width
            .Height = ControlInfos(Iter).Height
            .FontSize = ControlInfos(Iter).FontSize

        End With

        On Error GoTo 0

    Next
  
End Sub

Public Sub Render_Exp(ByVal Porcentaje As Boolean)

    Dim A As Long

    #If ModoBig = 0 Then

        If Porcentaje Then
        
            If FrmMain.PorcBloqued Then
                If UserPasarNivel > 0 Then
                    FrmMain.lblporclvl(0).Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
                Else
                    FrmMain.lblporclvl(0).Caption = UserLvl

                End If

            Else
                FrmMain.lblporclvl(0).Caption = UserLvl

            End If
        
            For A = FrmMain.lblporclvl.LBound To FrmMain.lblporclvl.UBound

                If UserPasarNivel > 0 Then
                    FrmMain.imgExp.Width = Round(((UserExp / 100) / (UserPasarNivel / 100)) * WIDTH_EXP)
                Else
                    FrmMain.imgExp.Width = 138

                End If
            
            Next A
    
        Else

            If UserPasarNivel > 0 Then
                FrmMain.lblporclvl(0).Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
            Else
                FrmMain.lblporclvl(0).Caption = UserLvl

            End If
    
        End If
    
    #Else
    
        If Porcentaje Then
        
            If FrmMain.PorcBloqued Then
                If UserPasarNivel > 0 Then
                    FrmMain.lblporclvl(2).Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
                    FrmMain.lblporclvl(1).visible = True
                    FrmMain.lblporclvl(1).Caption = Format$(UserExp, "#,###") & "/" & Format$(UserPasarNivel, "#,###")
                Else
                    FrmMain.lblporclvl(2).Caption = UserLvl
                    
                    If UserPasarNivel > 0 Then
                        FrmMain.lblporclvl(1).Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"

                    End If

                End If
                
            Else
                FrmMain.lblporclvl(2).Caption = UserLvl

                If UserPasarNivel > 0 Then
                    FrmMain.lblporclvl(1).Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"

                End If

            End If
    
        Else
        
            FrmMain.lblporclvl(2).Caption = UserLvl

            'FrmMain.lblporclvl(1).visible = False
            If UserPasarNivel > 0 Then
                FrmMain.lblporclvl(1).Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"

            End If

        End If
    
    #End If

    If UserPasarNivel > 0 Then
        FrmMain.imgExp.Width = Round(((UserExp / 100) / (UserPasarNivel / 100)) * WIDTH_EXP)
    Else
        FrmMain.imgExp.Width = WIDTH_EXP
        
    End If
 
End Sub

Function EliminarEspeciales(ByVal Text As String, _
                            Optional ByVal Filtro As String = vbNullString) As String

    Dim i        As Integer

    Dim TempText As String
    
    TempText = Text
    
    For i = 1 To Len(Filtro)
        Text = Replace(Text, mid(Filtro, i, 1), "")
    Next
    
    EliminarEspeciales = Text

End Function

Function ValidarNombre(Nombre As String) As Boolean
    
    Dim Temp As String

    Temp = UCase$(Nombre)
    
    Dim i           As Long, Char As Integer, LastChar As Integer

    Dim CantEspacio As Byte
        
    For i = 1 To Len(Temp)
        Char = Asc(mid$(Temp, i, 1))
        
        If (Char < 65 Or Char > 90) And Char <> 32 Then
            Exit Function
        
        ElseIf Char = 32 Then
                
            If LastChar = 32 Then
                Exit Function

            End If
                
            CantEspacio = CantEspacio + 1
                
            If CantEspacio > 1 Then Exit Function

        End If
        
        LastChar = Char

    Next
          
    If Asc(mid$(Temp, 1, 1)) = 32 Then
        Exit Function

    End If

    '
    ValidarNombre = True

End Function

Public Function PonerPuntos(Numero As Long) As String
    
    On Error GoTo PonerPuntos_Err

    Dim i     As Integer

    Dim Cifra As String
 
    Cifra = str(Numero)
    Cifra = Right$(Cifra, Len(Cifra) - 1)

    For i = 0 To 4

        If Len(Cifra) - 3 * i >= 3 Then
            If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
                PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos

            End If

        Else

            If Len(Cifra) - 3 * i > 0 Then
                PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos

            End If

            Exit For

        End If

    Next
 
    PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)
    
    Exit Function

PonerPuntos_Err:

    'Call RegistrarError(err.Number, err.Description, "ModLadder.PonerPuntos", Erl)
    Resume Next
    
End Function

Public Function IsActionParaCliente(ByVal ObjIndex As Integer) As Boolean

    If ObjIndex = 0 Then Exit Function
    IsActionParaCliente = True
    
    Select Case ObjData(ObjIndex).ObjType
    
        Case eOBJType.otActaNick
            Call FrmChangeNick.Show(, FrmMain)
            Exit Function
            
        Case eOBJType.otActaGuild
            Call FrmChangeNickGuild.Show(, FrmMain)
            Exit Function
            
    End Select
    
    IsActionParaCliente = False

End Function

' @ Devuelve el Color del objeto seg�n espec�ficaciones del servidor (tier)
Public Function DataServer_ColorObj(ByVal Tier As Byte) As Long
    
    Select Case Tier
    
        Case 0      ' Sin Color
            DataServer_ColorObj = 0

        Case 1      ' Color Gris-neutro
            DataServer_ColorObj = ARGB(230, 230, 230, 255)

        Case 2      ' Color Verde
            DataServer_ColorObj = ARGB(145, 245, 118, 255)

        Case 3      ' Color Cyan
            DataServer_ColorObj = ARGB(118, 245, 245, 255)

        Case 4      ' Color Violeta
            DataServer_ColorObj = ARGB(180, 118, 245, 255)

        Case 5      ' Color Amarillo Power
            DataServer_ColorObj = ARGB(250, 250, 0, 255)

    End Select
    
End Function

Public Function IntervaloPermiteHeading(Optional ByVal Actualizar As Boolean = True) As Boolean
    
    On Error GoTo IntervaloPermiteHeading_Err

    If FrameTime - IntervaloHeading >= CONST_INTERVALO_HEADING Then
        If Actualizar Then
            IntervaloHeading = FrameTime

        End If

        IntervaloPermiteHeading = True
        'Call AddToConsole(  "Golpe - Magia OK.", 255, 0, 0, True, False, False)
    Else
        IntervaloPermiteHeading = False

        'Call AddToConsole(  "Golpe - Magia NO.", 255, 0, 0, True, False, False)
    End If
    
    Exit Function

IntervaloPermiteHeading_Err:

    Resume Next
    
End Function

' # Buscamos si el usuario la tiene comprada.

Public Function Skin_SearchUser(ByVal ObjIndex As Integer) As Integer

    Dim A As Long
    
    With ClientInfo.Skin

        For A = 1 To .Last

            If .ObjIndex(A) = ObjIndex Then
                Skin_SearchUser = A
                Exit Function

            End If

        Next A

    End With
    
End Function

' # Chequea si tiene equipado algo para remarcar
Public Function Skins_CheckingItems(ByVal ObjIndex As Integer) As Boolean

    With ClientInfo.Skin

        If .ArmourIndex = ObjIndex Or .HelmIndex = ObjIndex Or .WeaponIndex = ObjIndex Or .WeaponDagaIndex = ObjIndex Or .WeaponArcoIndex = ObjIndex Or .ShieldIndex = ObjIndex Then
            
            Skins_CheckingItems = True

        End If
        
    End With

End Function

''
' Calculates the selling price of an item (The price that a merchant will sell you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.

Public Function CalculateSellPrice(ByRef objValue As Single, _
                                   ByVal ObjAmount As Long) As Long

    '*************************************************
    'Author: Marco Vanotti (MarKoxX)
    'Last modified: 19/08/2008
    'Last modify by: Franco Zeoli (Noich)
    '*************************************************
    On Error GoTo error

    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateSellPrice = CCur(objValue * 1000000) / 1000000 * ObjAmount + 0.5
          
    Exit Function

error:
    MsgBox err.Description, vbExclamation, "Error: " & err.Number

End Function
