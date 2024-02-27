Attribute VB_Name = "Mod_General"
'Desterium AO 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'Desterium AO is based on Baronsoft's VB6 Online RPG
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


Private Type OSVERSIONINFO
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long
    PlatformID As Long
    szCSDVersion As String * 128
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
                                      (lpVersionInformation As OSVERSIONINFO) As Long
                                      
'Hardwareid
Private Const READ_CONTROL As Long = &H20000
Private Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)
Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
Private Const SYNCHRONIZE As Long = &H100000
Private Const KEY_WOW64_64KEY As Long = &H100  '// 64-bit Key
Private Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const REG_SZ = 1
Private Const ERROR_SUCCESS = 0&

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'HDD
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

'Build
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Declare Function GetVolumeSerialNumber Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public iplst As String
Public Cpu_ID  As String
Public Cpu_SSN As String
Public bFogata As Boolean

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Public alphaT As Double

Private lFrameTimer As Long

Public Function DirGraficos() As String
    DirGraficos = App.path & "\RECURSOS\"
End Function

Public Function DirSound() As String
    DirSound = App.path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
    DirMidi = App.path & "\" & Config_Inicio.DirMusica & "\"
End Function

Public Function DirMapas() As String
    DirMapas = App.path & "\RECURSOS\"
End Function

Public Function DirExtras() As String
    DirExtras = App.path & "\EXTRAS\"
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
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

    Dim loopC As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopC = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopC).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopC, "Dir1")), 0
        InitGrh WeaponAnimData(loopC).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopC, "Dir2")), 0
        InitGrh WeaponAnimData(loopC).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopC, "Dir3")), 0
        InitGrh WeaponAnimData(loopC).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopC, "Dir4")), 0
    Next loopC
End Sub

Sub CargarColores()
On Error Resume Next
    Dim archivoC As String
    
    archivoC = App.path & "\init\colores.dat"
    
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
    
    ' Atacable
    ColoresPJ(48).r = CByte(GetVar(archivoC, "AT", "R"))
    ColoresPJ(48).g = CByte(GetVar(archivoC, "AT", "G"))
    ColoresPJ(48).b = CByte(GetVar(archivoC, "AT", "B"))
    
    ' @@Admins vbWhite.
    
    ColoresPJ(4).r = 255
    ColoresPJ(4).g = 255
    ColoresPJ(4).b = 255
    
End Sub

#If SeguridadAlkon Then
Sub InitMI()
    Dim alternativos As Integer
    Dim CualMITemp As Integer
    
    alternativos = RandomNumber(1, 7368)
    CualMITemp = RandomNumber(1, 1233)
    

    Set mi(CualMITemp) = New clsManagerInvisibles
    Call mi(CualMITemp).Inicializar(alternativos, 10000)
    
    If CualMI <> 0 Then
        Call mi(CualMITemp).CopyFrom(mi(CualMI))
        Set mi(CualMI) = Nothing
    End If
    CualMI = CualMITemp
End Sub
#End If

Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopC As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopC = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopC).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopC, "Dir1")), 0
        InitGrh ShieldAnimData(loopC).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopC, "Dir2")), 0
        InitGrh ShieldAnimData(loopC).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopC, "Dir3")), 0
        InitGrh ShieldAnimData(loopC).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopC, "Dir4")), 0
    Next loopC
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = True)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'******************************************r
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

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopC As Long
    
    For loopC = 1 To LastChar
        If charlist(loopC).Active = 1 Then
            MapData(charlist(loopC).Pos.X, charlist(loopC).Pos.Y).CharIndex = loopC
        End If
    Next loopC
End Sub
Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopC As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Dirección de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopC = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopC, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopC
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 15 Then
        MsgBox ("El nombre debe tener menos de 15 letras.")
        Exit Function
    End If
    
    For loopC = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, loopC, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopC
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

#If SeguridadAlkon Then
    Call UnprotectForm
#End If

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
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
        Exit Function
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
    
#If SeguridadAlkon Then
    'Unprotect character creation form
    Call UnprotectForm
#End If
    
    'Unload the connect form
    Unload frmCrearPersonaje
    Unload frmConnect
    Dim F As Long
    For F = 0 To 6
    Next F
    
    frmMain.Label8(0).Caption = UserName
    frmMain.Label8(1).Caption = UserName
    frmMain.Label8(2).Caption = UserName
    frmMain.Label8(3).Caption = UserName
    frmMain.Label8(4).Caption = UserName
    'Load main form
    frmMain.Visible = True
    If charlist(UserCharIndex).priv > 0 Then
frmMain.Labelgm3.Visible = True
frmMain.Labelgm4.Visible = True
frmMain.Labelgm44.Visible = True
frmMain.Label5.Visible = True

'frmMain.Labelgm6.Visible = True
'frmMain.Labelgm7.Visible = True
'frmMain.Labelgm8.Visible = True
'frmMain.Labelgm9.Visible = True
'frmMain.Labelgm10.Visible = True
'frmMain.Labelgm11.Visible = True
'frmMain.Labelgm12.Visible = True
frmMain.Label1.Visible = True
'frmMain.Line1.Visible = True

Else
frmMain.Labelgm3.Visible = False
frmMain.Labelgm4.Visible = False
frmMain.Labelgm44.Visible = False
frmMain.Label5.Visible = False

'frmMain.Labelgm6.Visible = False
'frmMain.Labelgm7.Visible = False
'frmMain.Labelgm8.Visible = False
'frmMain.Labelgm9.Visible = False
'frmMain.Labelgm10.Visible = False
'frmMain.Labelgm11.Visible = False
'frmMain.Labelgm12.Visible = False
frmMain.Label1.Visible = False
'frmMain.Line1.Visible = False
End If
    
    Call frmMain.ControlSM(eSMType.mSpells, False)
    Call frmMain.ControlSM(eSMType.mWork, False)
    
    FPSFLAG = True
#If SeguridadAlkon Then
    'Protect the main form
    Call ProtectForm(frmMain)
#End If

End Sub

Sub CargarTip()
    Dim n As Integer
    n = RandomNumber(1, UBound(Tips))
    
    frmtip.tip.Caption = Tips(n)
End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/28/2008
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
'***************************************************
    Dim LegalOk As Boolean
    
    If Cartel Then Cartel = False
    
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
   If UserMeditar Then UserMeditar = False
        Call WriteWalk(Direccion)
        Call DibujarMiniMapa(frmMain.Minimap)
 
        If Not UserDescansar And Not UserMeditar Then
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion
        End If
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            Call WriteChangeHeading(Direccion)
        End If
    End If
    
    If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
    
    ' Update 3D sounds!
    Call Audio.MoveListener(UserPos.X, UserPos.Y)
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo'dame 1seg
'***************************************************
    Call MoveTo(RandomNumber(NORTH, WEST))
End Sub

Public Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
    Static LastMovement As Long
    
    'No input allowed while Argentum is not the active window
    If Not Application.IsAppActive() Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'No walking while writting in the forum.
    If MirandoForo Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Then Exit Sub
    
    'TODO: Debería informarle por consola?
    If Traveling Then Exit Sub

    'Control movement interval (this enforces the 1 step loss when meditating / resting client-side)
    If GetTickCount - LastMovement > 56 Then
        LastMovement = GetTickCount
    Else
        Exit Sub
    End If
    
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(NORTH)
                frmMain.Coord.Caption = "Mapa " & UserMap & " [" & UserPos.X & "," & UserPos.Y & "]"
                frmMain.lblmapaname.Caption = MapaActual
                Exit Sub
            End If
            
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(EAST)
                'frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.y & ")"
                frmMain.Coord.Caption = "Mapa " & UserMap & " [" & UserPos.X & "," & UserPos.Y & "]"
                frmMain.lblmapaname.Caption = MapaActual
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(SOUTH)
                frmMain.Coord.Caption = "Mapa " & UserMap & " [" & UserPos.X & "," & UserPos.Y & "]"
                frmMain.lblmapaname.Caption = MapaActual
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(WEST)
                frmMain.Coord.Caption = "Mapa " & UserMap & " [" & UserPos.X & "," & UserPos.Y & "]"
                frmMain.lblmapaname.Caption = MapaActual
                Exit Sub
            End If
            
            ' We haven't moved - Update 3D sounds!
            Call Audio.MoveListener(UserPos.X, UserPos.Y)
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
            If kp Then
                Call RandomMove
            Else
                ' We haven't moved - Update 3D sounds!
                Call Audio.MoveListener(UserPos.X, UserPos.Y)
            End If
            
            If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
            'frmMain.Coord.Caption = "(" & UserPos.x & "," & UserPos.y & ")"
            frmMain.Coord.Caption = "Mapa " & UserMap & " [" & UserPos.X & "," & UserPos.Y & "]"
            frmMain.lblmapaname.Caption = MapaActual
        End If
    End If
End Sub
Private Sub Char_CleanAll()
 '// Borramos los obj y char que esten
 
        Dim X As Long, Y As Long
 
        For X = XMinMapSize To XMaxMapSize
                For Y = YMinMapSize To YMaxMapSize
 
                        With MapData(X, Y)
 
                                If (.CharIndex) Then
Call EraseChar(.CharIndex)
End If

If (.ObjGrh.GrhIndex) Then
MapData(X, Y).ObjGrh.GrhIndex = 0
End If
 
                        End With
 
                Next Y
        Next X
 
End Sub
'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!
Sub SwitchMap(ByVal map As Integer)
    Dim Y As Long
    Dim X As Long
    Dim tempint As Integer
    Dim ByFlags As Byte
    Dim handle As Integer
    
    handle = FreeFile()
    
    
    If LenB(Dir(DirMapas & "tmp.map", vbArchive)) <> 0 Then Kill DirMapas & "tmp.map"
    
    Dim data() As Byte
    Call Get_File_Data(DirMapas, "MAPA" & CStr(map) & ".MAP", data, 1)
       '// Borramos todos los char y objetos de mapa, antes de cargar el nuevo mapa
      Call Char_CleanAll
    
    Open DirMapas & "tmp.map" For Binary As handle
        Put handle, , data
    Close handle
    
    ' Cargamos el mapa...
    Open DirMapas & "tmp.map" For Binary As handle
    Seek handle, 1
            
    'map Header
    Get handle, , MapInfo.MapVersion
    Get handle, , MiCabecera
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            Get handle, , ByFlags
            
            MapData(X, Y).Blocked = (ByFlags And 1)
            
            Get handle, , MapData(X, Y).Graphic(1).GrhIndex
            InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get handle, , MapData(X, Y).Graphic(2).GrhIndex
                InitGrh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex
            Else
                MapData(X, Y).Graphic(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get handle, , MapData(X, Y).Graphic(3).GrhIndex
                InitGrh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex
            Else
                MapData(X, Y).Graphic(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get handle, , MapData(X, Y).Graphic(4).GrhIndex
                InitGrh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndex
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
    
      If LenB(Dir(DirMapas & "tmp.map", vbArchive)) <> 0 Then Kill DirMapas & "tmp.map"
    
    MapInfo.name = ""
    MapInfo.Music = ""
    
    CurMap = map
    
    Call GenerarMiniMapa
    Call DibujarMiniMapa(frmMain.Minimap)
    

    Exit Sub
errorH:     ' Mapas
    
    If LenB(Dir(DirMapas & "tmp.map", vbArchive)) <> 0 Then Kill DirMapas & "tmp.map"
    Call MsgBox("Error en el formato del Mapa " & map, vbCritical + vbOKOnly, "Argentum Online")
    CloseClient ' Cerramos el cliente
End Sub

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
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
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim Count As Long
    Dim curPos As Long
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

Public Function IsIp(ByVal Ip As String) As Boolean
    Dim i As Long
    
    For i = 1 To UBound(ServersLst)
        If ServersLst(i).Ip = Ip Then
            IsIp = True
            Exit Function
        End If
    Next i
End Function
Private Function SystemSerialNumber32() As String
Dim mother_boards As Variant
Dim board As Variant
Dim wmi As Variant
Dim serial_numbers As String

    ' Get the Windows Management Instrumentation object.
    Set wmi = GetObject("WinMgmts:")

    ' Get the "base boards" (mother boards).
    Set mother_boards = wmi.InstancesOf("Win32_BaseBoard")
    For Each board In mother_boards
        serial_numbers = serial_numbers & ", " & _
            board.SerialNumber
    Next board
    If Len(serial_numbers) > 0 Then serial_numbers = _
        mid$(serial_numbers, 3)

    SystemSerialNumber32 = serial_numbers
End Function

Private Function SystemSerialNumber64() As String

    Dim sVolLabel As String, sSerial As Long, sName As String
    Dim objCPUItem As Object, objCPU As Object
    Dim objBaseBoard As Object, objBoard As Object
    Dim i As Long, stmp As Long, sresult As String

    If GetVolumeSerialNumber(Left(Environ("windir"), 3), sVolLabel, 0, sSerial, 0, 0, sName, 0) Then
  SystemSerialNumber64 = Hex(sSerial)
    Else
  SystemSerialNumber64 = "00"
    End If

    Set objCPUItem = GetObject("winmgmts:").InstancesOf("Win32_Processor")

    If Err.number = 0 Then
  For Each objCPU In objCPUItem
    SystemSerialNumber64 = SystemSerialNumber64 & objCPU.MaxClockSpeed
    SystemSerialNumber64 = SystemSerialNumber64 & objCPU.ProcessorId
  Next
  Set objCPUItem = Nothing
    End If

    Set objBaseBoard = GetObject("WinMgmts:").InstancesOf("Win32_BaseBoard")

    For Each objBoard In objBaseBoard
  SystemSerialNumber64 = SystemSerialNumber64 & objBoard.SerialNumber
  Exit For
    Next

    For i = 1 To Len(SystemSerialNumber64)
  If IsNumeric(mid(SystemSerialNumber64, i, 1)) = True Then
    stmp = stmp + mid(SystemSerialNumber64, i, 1)
  End If
    Next

    For i = 1 To Len(stmp)
  sresult = sresult & mid(stmp, i, 1) Mod 12
    Next

    SystemSerialNumber64 = StrReverse(sresult)

End Function

Private Function CpuId() As String
Dim computer As String
Dim wmi As Variant
Dim processors As Variant
Dim cpu As Variant
Dim cpu_ids As String

    computer = "."
    Set wmi = GetObject("winmgmts:" & _
        "{impersonationLevel=impersonate}!\\" & _
        computer & "\root\cimv2")
    Set processors = wmi.ExecQuery("Select * from " & _
        "Win32_Processor")

    For Each cpu In processors
        cpu_ids = cpu_ids & ", " & cpu.ProcessorId
    Next cpu
    If Len(cpu_ids) > 0 Then cpu_ids = mid$(cpu_ids, 3)

    CpuId = cpu_ids
End Function


Public Function GWV() As String

    Dim osv As OSVERSIONINFO
    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.PlatformID
            Case VER_PLATFORM_WIN32s
                GWV = "Win32s on W 3.1"
            Case VER_PLATFORM_WIN32_NT
                GWV = "W NT"

                Select Case osv.dwVerMajor
                    Case 3
                        GWV = "W NT 3.5"
                    Case 4
                        GWV = "W NT 4.0"
                    Case 5
                        Select Case osv.dwVerMinor
                            Case 0
                                GWV = "W 2000"
                            Case 1
                                GWV = "W XP"
                            Case 2
                                GWV = "W Server 2003"
                        End Select
                    Case 6
                        Select Case osv.dwVerMinor
                            Case 0
                                GWV = "W Vista"
                            Case 1
                                GWV = "W 7"
                            Case 2
                                GWV = "W 8"
                            Case 3
                                GWV = "W 8.1"
                        End Select
                End Select

            Case VER_PLATFORM_WIN32_WINDOWS:
                Select Case osv.dwVerMinor
                    Case 0
                        GWV = "W95"
                    Case 90
                        GWV = "w90"
                    Case Else
                        GWV = "w98"
                End Select
        End Select
    Else
        GWV = "Windows desconocido."
    End If
End Function

Sub Main()

DirButtons = App.path & "\GRAFICOS\"

Dim MySO As String
MySO = GWV

Select Case MySO

Case "W 8.1", "W 8"
    Windows = 32
Case "W 7", "W XP"
    Windows = 16
Case Else
    Windows = 32
End Select


Cpu_ID = CpuId

Cpu_SSN = SystemSerialNumber32

If Cpu_SSN = vbNullString Then Cpu_SSN = SystemSerialNumber64


    Call m_Damages.Initialize
    Call FotoD_Initialize
    Call LoadRecup
    Call ModResolution.Generate_Array
     
    'Load config file
    If FileExist(App.path & "\init\Inicio.con", vbNormal) Then
        Config_Inicio = LeerGameIni()
    End If
      If FileExist(App.path & "\Init\Config.con", vbNormal) Then
        TSetup = ReadOptionIni()
    End If
      ' Call GenCM("G2W8H9364") ' GS - Iniciamos la clave maestra!
    Call modCompression.GenerateContra("DGraps1", 0) ' graficos
   Call modCompression.GenerateContra("DMapas1", 1) ' mapas
    

    'Load ao.dat config file
    Call LoadClientSetup
    
    If ClientSetup.bDinamic Then
        Set SurfaceDB = New clsSurfaceManDyn
    Else
        Set SurfaceDB = New clsSurfaceManStatic
    End If
    


    'Read command line. Do it AFTER config file is loaded to prevent this from
    'canceling the effects of "/nores" option.
    Call LeerLineaComandos
    
        'If Not App.EXEName = "DS" Then End
    
    If FindPreviousInstance Then
      '  Call MsgBox("¡Desterium AO ya está corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
      '  End
    End If
    
    
    'usaremos esto para ayudar en los parches
    Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.path & "\")
    
    ChDrive App.path
    ChDir App.path

#If SeguridadAlkon Then
    'Obtener el HushMD5
    Dim fMD5HushYo As String * 32
    
    fMD5HushYo = md5.GetMD5File(App.path & "\" & App.EXEName & ".exe")
    Call md5.MD5Reset
    MD5HushYo = txtOffset(hexMd52Asc(fMD5HushYo), 55)
    
    Debug.Print fMD5HushYo
#Else
    ' Por esto explota... Debe hacer un MD5 a cualquier cosa jajajaj.
    'MD5HushYo = MD5File(App.path & "\" & App.EXEName & ".exe")  'We aren't using a real MD5
#End If
    
    tipf = Config_Inicio.tip
    
    'Set resolution BEFORE the loading form is displayed, therefore it will be centered.
    
    ' Mouse Pointer (Loaded before opening any form with buttons in it)
    If FileExist(DirExtras & "Hand.ico", vbArchive) Then _
        Set picMouseIcon = LoadPicture(DirExtras & "Hand.ico")
    
    frmCargando.Show
        'frmCargando.Analizar
    frmCargando.Refresh
    
    frmConnect.version = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    Call AddtoRichTextBox(frmCargando.status, "Buscando servidores... ", 255, 255, 255, True, False, True)
frmCargando.barra.Width = frmCargando.barra.Width + 25
    
'TODO : esto de ServerRecibidos no se podría sacar???
    ServersRecibidos = True
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    Call AddtoRichTextBox(frmCargando.status, "Iniciando constantes... ", 255, 255, 255, True, False, True)
    frmCargando.barra.Width = frmCargando.barra.Width + 35
    Call InicializarNombres
    
    ' Initialize FONTTYPES
    Call Protocol.InitFonts
    
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.status, "Iniciando motor gráfico... ", 255, 255, 255, True, False, True)
    frmCargando.barra.Width = frmCargando.barra.Width + 45
    
    If Not InitTileEngine(frmMain.hwnd, frmMain.MainViewShp.Top, frmMain.MainViewShp.Left, 32, 32, frmMain.MainViewShp.Height / 32, frmMain.MainViewShp.Width / 32, 9, 6, 6, 0.018) Then
        Call CloseClient
    End If
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.status, "Creando animaciones extra... ", 255, 255, 255, True, False, True)
    


    frmCargando.barra.Width = frmCargando.barra.Width + 65
    Call CargarTips
    
UserMap = 1
    
    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    CargarAuras
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectSound... ", 255, 255, 255, True, False, True)
    frmCargando.barra.Width = frmCargando.barra.Width + 85
    'Inicializamos el sonido
    Call Audio.Initialize(DirectX, frmMain.hwnd, App.path & "\" & Config_Inicio.DirSonidos & "\", App.path & "\" & Config_Inicio.DirMusica & "\")
    'Enable / Disable audio
    Audio.MusicActivated = Not ClientSetup.bNoMusic
    Audio.SoundActivated = Not ClientSetup.bNoSound
    Audio.SoundEffectsActivated = Not ClientSetup.bNoSoundEffects
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(DirectDraw, frmMain.PicInv, MAX_INVENTORY_SLOTS, , , , , , , , , True)
    
    Call Audio.MusicMP3Play(App.path & "\MP3\" & MP3_Inicio & ".mp3")
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    
#If SeguridadAlkon Then
    CualMI = 0
    Call InitMI
#End If
    
    Call AddtoRichTextBox(frmCargando.status, "                    ¡Bienvenido a Desterium AO!", 255, 255, 255, True, False, True)
    frmCargando.barra.Width = frmCargando.barra.Width + 85
    
    'Give the user enough time to read the welcome text
    Call Sleep(250)
    
    Call Resolution.SetResolution
    
    frmConnect.version = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision


'TODO : esto de ServerRecibidos no se podría sacar???
    ServersRecibidos = True
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)

    Call InicializarNombres
    
    ' Initialize FONTTYPES
    Call Protocol.InitFonts
    
    
    If Not InitTileEngine(frmMain.hwnd, frmMain.MainViewShp.Top, frmMain.MainViewShp.Left, 32, 32, frmMain.MainViewShp.Height / 32, frmMain.MainViewShp.Width / 32, 9, 8, 8, 0.018) Then
        Call CloseClient
    End If
    

    Call CargarTips
    
UserMap = 1
    
    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    'Inicializamos el sonido
    Call Audio.Initialize(DirectX, frmMain.hwnd, App.path & "\" & Config_Inicio.DirSonidos & "\", App.path & "\" & Config_Inicio.DirMusica & "\")
    'Enable / Disable audio
    Audio.MusicActivated = Not ClientSetup.bNoMusic
    Audio.SoundActivated = Not ClientSetup.bNoSound
    Audio.SoundEffectsActivated = Not ClientSetup.bNoSoundEffects
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(DirectDraw, frmMain.PicInv, MAX_INVENTORY_SLOTS, , , , , , , , , True)
    
    
    
    
#If SeguridadAlkon Then
    CualMI = 0
    Call InitMI
#End If
    
    
    Unload frmCargando
    
#If UsarWrench = 1 Then
    frmMain.Socket1.Startup
#End If

    frmConnect.Visible = True
    
    'Inicialización de variables globales
    PrimeraVez = True
    prgRun = True
    pausa = False
    
    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
    
    frmMain.macrotrabajo.Interval = INT_MACRO_TRABAJO
    frmMain.macrotrabajo.Enabled = False
    
   'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
    
    'Set the dialog's font
    Dialogos.font = frmMain.font
    DialogosClanes.font = frmMain.font
    
    lFrameTimer = GetTickCount
    
    ' Load the form for screenshots
    Call Load(frmScreenshots)
    Call CargarAnimsExtra
        
    Do While prgRun
        'Sólo dibujamos si la ventana no está minimizada
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY)
            
            If bTecho Then
                If alphaT >= 3 Then
                alphaT = alphaT - 3
                End If
            Else
                If alphaT <= 252 Then
                alphaT = alphaT + 3
                End If
            End If
            
            'Play ambient sounds
            Call RenderSounds
            
            Call CheckKeys
        End If
 'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then
      '  loopC = gsStatus
      '  If loopC <> 0 Then
        '   If loopC = 2 Then
        '     If Connected = True Then Call gsWriteCheatInfo
        '      Sleep 5
        '       prgRun = False
        '   End If
    '   Else
     '    prgRun = False
     ' End If
            If FPSFLAG Then frmMain.lblFPS.Caption = Mod_TileEngine.FPS
            
            lFrameTimer = GetTickCount
        End If
        
#If SeguridadAlkon Then
        Call CheckSecurity
#End If
        
        ' If there is anything to be sent, we send it
        Call FlushBuffer
        
        DoEvents
    Loop
    
    Call CloseClient
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, File
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
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= 13562)) And _
                MapData(X, Y).Graphic(2).GrhIndex = 0
                
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

Public Sub ShowSendCMSGTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendCMSTXT.Visible = True
        frmMain.SendCMSTXT.SetFocus
    End If
End Sub

''
' Checks the command line parameters, if you are running Ao with /nores command and checks the AoUpdate parameters
'
'

Public Sub LeerLineaComandos()
'*************************************************
'Author: Unknown
'Last modified: 25/11/2008 (BrianPr)
'
'*************************************************
    Dim t() As String
    Dim i As Long
    
    Dim UpToDate As Boolean
    Dim Patch As String
    
    'Parseo los comandos
    t = Split(Command, " ")
    For i = LBound(t) To UBound(t)
        Select Case UCase$(t(i))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
            Case "/UPTODATE"
                UpToDate = True
        End Select
    Next i
    
    'Call AoUpdate(UpToDate, NoRes) ' www.gs-zone.org
End Sub

''
' Runs AoUpdate if we haven't updated yet, patches aoupdate and runs Client normally if we are updated.
'
' @param UpToDate Specifies if we have checked for updates or not
' @param NoREs Specifies if we have to set nores arg when running the client once again (if the AoUpdate is executed).

Private Sub AoUpdate(ByVal UpToDate As Boolean, ByVal NoRes As Boolean)
'*************************************************
'Author: BrianPr
'Created: 25/11/2008
'Last modified: 25/11/2008
'
'*************************************************
On Error GoTo error
    Dim extraArgs As String
    If Not UpToDate Then
        'No recibe update, ejecutar AU
        'Ejecuto el AoUpdate, sino me voy
        If Dir(App.path & "\AoUpdate.exe", vbArchive) = vbNullString Then
            MsgBox "No se encuentra el archivo de actualización AoUpdate.exe por favor descarguelo y vuelva a intentar", vbCritical
            End
        Else
            FileCopy App.path & "\AoUpdate.exe", App.path & "\AoUpdateTMP.exe"
            
            If NoRes Then
                extraArgs = " /nores"
            End If
            
            Call ShellExecute(0, "Open", App.path & "\AoUpdateTMP.exe", App.EXEName & ".exe" & extraArgs, App.path, SW_SHOWNORMAL)
            End
        End If
    Else
        If FileExist(App.path & "\AoUpdateTMP.exe", vbArchive) Then Kill App.path & "\AoUpdateTMP.exe"
    End If
Exit Sub

error:
    If Err.number = 75 Then 'Si el archivo AoUpdateTMP.exe está en uso, entonces esperamos 5 ms y volvemos a intentarlo hasta que nos deje.
        Sleep 5
        Resume
    Else
        MsgBox Err.Description & vbCrLf, vbInformation, "[ " & Err.number & " ]" & " Error "
        End
    End If
End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/19/09
'11/19/09: Pato - Is optional show the frmGuildNews form
'**************************************************************
    Dim fHandle As Integer
    
    If FileExist(App.path & "\init\ao.dat", vbArchive) Then
        fHandle = FreeFile
        
        Open App.path & "\init\ao.dat" For Binary Access Read Lock Write As fHandle
            Get fHandle, , ClientSetup
        Close fHandle
    Else
        'Use dynamic by default
        ClientSetup.bDinamic = True
    End If
    
    NoRes = ClientSetup.bNoRes
    
    If InStr(1, ClientSetup.sGraficos, "Graficos") Then
        GraphicsFile = ClientSetup.sGraficos
    Else
        GraphicsFile = "Graficos3.ind"
    End If
    
    ClientSetup.bGuildNews = Not ClientSetup.bGuildNews
    DialogosClanes.Activo = Not ClientSetup.bGldMsgConsole
    DialogosClanes.CantidadDialogos = ClientSetup.bCantMsgs
End Sub

Private Sub SaveClientSetup()
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 03/11/10
'
'**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    
    ClientSetup.bNoMusic = Not Audio.MusicActivated
    ClientSetup.bNoSound = Not Audio.SoundActivated
    ClientSetup.bNoSoundEffects = Not Audio.SoundEffectsActivated
    ClientSetup.bGuildNews = Not ClientSetup.bGuildNews
    ClientSetup.bGldMsgConsole = Not DialogosClanes.Activo
    ClientSetup.bCantMsgs = DialogosClanes.CantidadDialogos
    
    Open App.path & "\init\ao.dat" For Binary As fHandle
        Put fHandle, , ClientSetup
    Close fHandle
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
    Ciudades(eCiudad.cNix) = "Nix"
    Ciudades(eCiudad.cBanderbill) = "Banderbill"
    Ciudades(eCiudad.cLindos) = "Lindos"
    Ciudades(eCiudad.cArghal) = "Arghâl"
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"

    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Worker) = "Trabajador"
    ListaClases(eClass.Pirat) = "Pirata"
    
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Tácticas de combate"
    SkillsNames(eSkill.Armas) = "Combate cuerpo a cuerpo"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar árboles"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"
    SkillsNames(eSkill.Equitacion) = "Equitacion"
    SkillsNames(eSkill.Resistencia) = "Resistencia Mágica"

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
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Removes all text from the console and dialogs
'**************************************************************
    'Clean console and dialogs
    frmMain.RecTxt.Text = vbNullString
    
    Call DialogosClanes.RemoveDialogs
    
    Call Dialogos.RemoveAllDialogs
End Sub

Public Sub CloseClient()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 8/14/2007
'Frees all used resources, cleans up and leaves
'**************************************************************
    ' Allow new instances of the client to be opened
    Call PrevInstance.ReleaseInstance
    
    EngineRun = False
    
    'Call AddtoRichTextBox(frmCargando.status, "Liberando recursos...", 0, 0, 0, 0, 0, 0)
    
    
    'Stop tile engine
    Call DeinitTileEngine
    
    Call SaveClientSetup
    
    'Destruimos los objetos públicos creados
    Set CustomMessages = Nothing
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    
#If SeguridadAlkon Then
    Set md5 = Nothing
#End If
    
    Call UnloadAllForms
    
    'Actualizar tip
    Config_Inicio.tip = tipf
    Call EscribirGameIni(Config_Inicio)
    End
End Sub
Public Function esGM(CharIndex As Integer) As Boolean
esGM = False
If charlist(CharIndex).priv >= 1 And charlist(CharIndex).priv <= 5 Or charlist(CharIndex).priv = 25 Then _
    esGM = True

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
Public Sub checkText(ByVal Text As String)
Dim Nivel As Integer
If Right(Text, Len(MENSAJE_FRAGSHOOTER_TE_HA_MATADO)) = MENSAJE_FRAGSHOOTER_TE_HA_MATADO Then
    Call ScreenCapture(True)
    Exit Sub
End If
If Left(Text, Len(MENSAJE_FRAGSHOOTER_HAS_MATADO)) = MENSAJE_FRAGSHOOTER_HAS_MATADO Then
    EsperandoLevel = True
    Exit Sub
End If
If EsperandoLevel Then
    If Right(Text, Len(MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA)) = MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA Then
        If CInt(mid(Text, Len(MENSAJE_FRAGSHOOTER_HAS_GANADO), (Len(Text) - (Len(MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA) + Len(MENSAJE_FRAGSHOOTER_HAS_GANADO))))) / 2 > ClientSetup.byMurderedLevel Then
            Call ScreenCapture(True)
        End If
    End If
End If
EsperandoLevel = False
End Sub

Public Function getStrenghtColor() As Long
Dim M As Long
M = 255 / MAXATRIBUTOS
getStrenghtColor = RGB(255 - (M * UserFuerza), (M * UserFuerza), 0)
End Function
Public Function getDexterityColor() As Long
Dim M As Long
M = 255 / MAXATRIBUTOS
getDexterityColor = RGB(255, M * UserAgilidad, 0)
End Function

Public Function getCharIndexByName(ByVal name As String) As Integer
Dim i As Long
For i = 1 To LastChar
    If charlist(i).nombre = name Then
        getCharIndexByName = i
        Exit Function
    End If
Next i
End Function

Public Function EsAnuncio(ByVal ForumType As Byte) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Returns true if the post is sticky.
'***************************************************
    Select Case ForumType
        Case eForumMsgType.ieCAOS_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieGENERAL_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieREAL_STICKY
            EsAnuncio = True
            
    End Select
    
End Function

Public Function ForumAlignment(ByVal yForumType As Byte) As Byte
'***************************************************
'Author: ZaMa
'Last Modification: 01/03/2010
'Returns the forum alignment.
'***************************************************
    Select Case yForumType
        Case eForumMsgType.ieCAOS, eForumMsgType.ieCAOS_STICKY
            ForumAlignment = eForumType.ieCAOS
            
        Case eForumMsgType.ieGeneral, eForumMsgType.ieGENERAL_STICKY
            ForumAlignment = eForumType.ieGeneral
            
        Case eForumMsgType.ieREAL, eForumMsgType.ieREAL_STICKY
            ForumAlignment = eForumType.ieREAL
            
    End Select
    
End Function
Public Sub General_Drop_X_Y(ByVal X As Byte, ByVal Y As Byte)

' /  Author  : Dunkan
' /  Note    : Calcular la posición de donde va a tirar el item

    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
        
        If Inventario.SelectedItem < 1 Then
        Call ShowConsoleMsg("No tienes esa cantidad.", 65, 190, 156, False, False)
            Inventario.sMoveItem = False
            Inventario.uMoveItem = False
            Exit Sub
        End If
        
        ' - Hay que pasar estas funciones al servidor
        If MapData(X, Y).Blocked = 1 Then
            Call ShowConsoleMsg("Elige una posición válida para tirar tus objetos.", 65, 190, 156, False, False)
            Inventario.sMoveItem = False
            Exit Sub
        End If
        
        If HayAgua(X, Y) = True Then
            Call ShowConsoleMsg("No está permitido tirar objetos en el agua.", 65, 190, 156, False, False)
            Inventario.sMoveItem = False
            Exit Sub
        End If

      If UserEstado = 1 Then
      Call ShowConsoleMsg("¡Estás muerto!", 65, 190, 156, False, False)
      Inventario.sMoveItem = False
      Exit Sub
      End If
      
      If UserMontando = True Then
       Call ShowConsoleMsg("Debes bajarte de la montura para poder arrojar items.", 65, 190, 156, False, False)
      Inventario.sMoveItem = False
        Exit Sub
        End If
        
        If UserNavegando = True Then
          Call ShowConsoleMsg("No puedes arrojar items mientras te encuentres navegando.", 65, 190, 156, False, False)
          Inventario.sMoveItem = False
          Exit Sub
          End If
          
          If Comerciando Then
                    Call ShowConsoleMsg("No puedes arrojar items mientras te encuentres comerciando.", 65, 190, 156, False, False)
          Inventario.sMoveItem = False
          Inventario.uMoveItem = False
          Exit Sub
          End If
          
        ' - Hay que pasar estas funciones al servidor
        
  If GetKeyState(vbKeyShift) < 0 Then
            FrmDrag.Show vbModal
        Else
            Call Protocol.WriteDragToPos(X, Y, Inventario.SelectedItem, 1)
        End If
    End If
     
    Inventario.sMoveItem = False
    
End Sub
