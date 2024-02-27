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

Public Canjes() As tCanjes
Public NumCanjes As Byte

Public Type tCanjes
    NumRequired As Byte
    
    ObjRequired(1 To 30) As Obj
    ObjCanje As Obj
    Points As Long
    GrhIndex As Long
End Type

Private Type tConfig
    NotWalkToConsole As Boolean
    ClickDerecho As Boolean
    
End Type

Public Config As tConfig

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
Private Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or _
    KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const REG_SZ = 1
Private Const ERROR_SUCCESS = 0&

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal _
    samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal _
    lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData _
    As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As _
    Long

'HDD
Private Declare Function GetVolumeInformation Lib "kernel32" Alias _
    "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal _
    lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
    lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal _
    nFileSystemNameSize As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
    (ByVal nDrive As String) As Long

'Build
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, _
    pSrc As Any, ByVal ByteLen As Long)

Private Declare Function GetVolumeSerialNumber Lib "kernel32.dll" Alias _
    "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal _
    lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
    lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal _
    nFileSystemNameSize As Long) As Long

Public iplst As String
Public Cpu_ID  As String
Public Cpu_SSN As String
Public bFogata As Boolean
Public alphaT As Double

Private lFrameTimer As Long
Public Const QS_HOTKEY = &H80
Public Const QS_KEY = &H1
Public Const QS_MOUSEBUTTON = &H4
Public Const QS_MOUSEMOVE = &H2
Public Const QS_PAINT = &H20
Public Const QS_POSTMESSAGE = &H8
Public Const QS_SENDMESSAGE = &H40
Public Const QS_TIMER = &H10
Public Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or _
    QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
Public Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Public Const QS_INPUT = (QS_MOUSE Or QS_KEY)
Public Const QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT _
    Or QS_HOTKEY)

Public Declare Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As _
    Long


Public Function cGetInputState()
      Dim qsRet As Long
10        qsRet = GetQueueStatus(QS_ALLEVENTS)
20        cGetInputState = qsRet
End Function

Public Function DirGraficos() As String
10        DirGraficos = App.path & "\RECURSOS\"
End Function

Public Function DirSound() As String
10        DirSound = App.path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
10        DirMidi = App.path & "\" & Config_Inicio.DirMusica & "\"
End Function

Public Function DirMapas() As String
10        DirMapas = App.path & "\RECURSOS\"
End Function

Public Function DirExtras() As String
10        DirExtras = App.path & "\EXTRAS\"
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As _
    Long) As Long
          'Initialize randomizer
10        Randomize Timer
          
          'Generate random number
20        RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Public Function GetRawName(ByRef sName As String) As String
      '***************************************************
      'Author: ZaMa
      'Last Modify Date: 13/01/2010
      'Last Modified By: -
      'Returns the char name without the clan name (if it has it).
      '***************************************************

          Dim Pos As Integer
          
10        Pos = InStr(1, sName, "<")
          
20        If Pos > 0 Then
30            GetRawName = Trim(Left(sName, Pos - 1))
40        Else
50            GetRawName = sName
60        End If

End Function

Sub CargarAnimArmas()
10    On Error Resume Next

          Dim LoopC As Long
          Dim arch As String
          
20        arch = App.path & "\init\" & "armas.dat"
          
30        NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
          
40        ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
          
50        For LoopC = 1 To NumWeaponAnims
60            InitGrh WeaponAnimData(LoopC).WeaponWalk(1), Val(GetVar(arch, "ARMA" & _
                  LoopC, "Dir1")), 0
70            InitGrh WeaponAnimData(LoopC).WeaponWalk(2), Val(GetVar(arch, "ARMA" & _
                  LoopC, "Dir2")), 0
80            InitGrh WeaponAnimData(LoopC).WeaponWalk(3), Val(GetVar(arch, "ARMA" & _
                  LoopC, "Dir3")), 0
90            InitGrh WeaponAnimData(LoopC).WeaponWalk(4), Val(GetVar(arch, "ARMA" & _
                  LoopC, "Dir4")), 0
100       Next LoopC
End Sub

Sub CargarColores()
10    On Error Resume Next
          Dim archivoC As String
          
20        archivoC = App.path & "\init\colores.dat"
          
30        If Not FileExist(archivoC, vbArchive) Then
      'TODO : Si hay que reinstalar, porque no cierra???
40            Call _
                  MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", _
                  vbCritical + vbOKOnly)
50            Exit Sub
60        End If
          
          Dim i As Long
          
70        For i = 0 To 48 '49 y 50 reservados para ciudadano y criminal
80            ColoresPJ(i).r = CByte(GetVar(archivoC, CStr(i), "R"))
90            ColoresPJ(i).g = CByte(GetVar(archivoC, CStr(i), "G"))
100           ColoresPJ(i).b = CByte(GetVar(archivoC, CStr(i), "B"))
110       Next i
          
          ' Crimi
120       ColoresPJ(50).r = CByte(GetVar(archivoC, "CR", "R"))
130       ColoresPJ(50).g = CByte(GetVar(archivoC, "CR", "G"))
140       ColoresPJ(50).b = CByte(GetVar(archivoC, "CR", "B"))
          
          ' Ciuda
150       ColoresPJ(49).r = CByte(GetVar(archivoC, "CI", "R"))
160       ColoresPJ(49).g = CByte(GetVar(archivoC, "CI", "G"))
170       ColoresPJ(49).b = CByte(GetVar(archivoC, "CI", "B"))
          
          ' Atacable
         ' ColoresPJ(48).r = CByte(GetVar(archivoC, "AT", "R"))
         ' ColoresPJ(48).g = CByte(GetVar(archivoC, "AT", "G"))
         ' ColoresPJ(48).b = CByte(GetVar(archivoC, "AT", "B"))
End Sub

Sub CargarAnimEscudos()
10    On Error Resume Next

          Dim LoopC As Long
          Dim arch As String
          
20        arch = App.path & "\init\" & "escudos.dat"
          
30        NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
          
40        ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
          
50        For LoopC = 1 To NumEscudosAnims
60            InitGrh ShieldAnimData(LoopC).ShieldWalk(1), Val(GetVar(arch, "ESC" & _
                  LoopC, "Dir1")), 0
70            InitGrh ShieldAnimData(LoopC).ShieldWalk(2), Val(GetVar(arch, "ESC" & _
                  LoopC, "Dir2")), 0
80            InitGrh ShieldAnimData(LoopC).ShieldWalk(3), Val(GetVar(arch, "ESC" & _
                  LoopC, "Dir3")), 0
90            InitGrh ShieldAnimData(LoopC).ShieldWalk(4), Val(GetVar(arch, "ESC" & _
                  LoopC, "Dir4")), 0
100       Next LoopC
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, _
    Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional _
    ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal _
    italic As Boolean = False, Optional ByVal bCrLf As Boolean = True)
      '******************************************
      'Adds text to a Richtext box at the bottom.
      'Automatically scrolls to new text.
      'Text box MUST be multiline and have a 3D
      'apperance!
      'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
      'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
      '******************************************r
10        With RichTextBox
20            If Len(.Text) > 1000 Then
                  'Get rid of first line
30                .SelStart = InStr(1, .Text, vbCrLf) + 1
40                .SelLength = Len(.Text) - .SelStart + 2
50                .TextRTF = .SelRTF
60            End If
              
70            .SelStart = Len(.Text)
80            .SelLength = 0
90            .SelBold = bold
100           .SelItalic = italic
              
110           If Not red = -1 Then .SelColor = RGB(red, green, blue)
              
120           If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
130           .SelText = Text
              
140           RichTextBox.Refresh
150       End With
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
      '*****************************************************************
      'Goes through the charlist and replots all the characters on the map
      'Used to make sure everyone is visible
      '*****************************************************************
          Dim LoopC As Long
          
10        For LoopC = 1 To LastChar
20            If charlist(LoopC).Active = 1 Then
30                MapData(charlist(LoopC).Pos.X, charlist(LoopC).Pos.Y).CharIndex = _
                      LoopC
40            End If
50        Next LoopC
End Sub
Function AsciiValidos(ByVal cad As String) As Boolean
          Dim car As Byte
          Dim i As Long
          
10        cad = LCase$(cad)
          
20        For i = 1 To Len(cad)
30            car = Asc(mid$(cad, i, 1))
              
40            If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And _
                  (car <> 32) Then
50                Exit Function
60            End If
70        Next i
          
80        AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
          'Validamos los datos del user
          Dim LoopC As Long
          Dim CharAscii As Integer
          
10        If checkemail And UserEmail = "" Then
20            MsgBox ("Dirección de email invalida")
30            Exit Function
40        End If
          
50        If UserPassword = "" Then
60            MsgBox ("Ingrese un password.")
70            Exit Function
80        End If
          
90        For LoopC = 1 To Len(UserPassword)
100           CharAscii = Asc(mid$(UserPassword, LoopC, 1))
110           If Not LegalCharacter(CharAscii) Then
120               MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & _
                      " no está permitido.")
130               Exit Function
140           End If
150       Next LoopC
          
160       If UserName = "" Then
170           MsgBox ("Ingrese un nombre de personaje.")
180           Exit Function
190       End If
          
200       If Len(UserName) > 15 Then
210           MsgBox ("El nombre debe tener menos de 15 letras.")
220           Exit Function
230       End If
          
240       For LoopC = 1 To Len(UserName)
250           CharAscii = Asc(mid$(UserName, LoopC, 1))
260           If Not LegalCharacter(CharAscii) Then
270               MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & _
                      " no está permitido.")
280               Exit Function
290           End If
300       Next LoopC
          
310       CheckUserData = True
End Function

Sub UnloadAllForms()
10    On Error Resume Next


          Dim mifrm As Form
          
20        For Each mifrm In Forms
30            Unload mifrm
40        Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
      '*****************************************************************
      'Only allow characters that are Win 95 filename compatible
      '*****************************************************************
          'if backspace allow
10        If KeyAscii = 8 Then
20            LegalCharacter = True
30            Exit Function
40        End If
          
          'Only allow space, numbers, letters and special characters
50        If KeyAscii < 32 Or KeyAscii = 44 Then
60            Exit Function
70        End If
          
80        If KeyAscii > 126 Then
90            Exit Function
100       End If
          
          'Check for bad special characters in between
110       If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or _
              KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or _
              KeyAscii = 124 Then
120           Exit Function
130       End If
          
          'else everything is cool
140       LegalCharacter = True
End Function
Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************

10    On Error GoTo ErrHandler
    'Set Connected
20        Connected = True
    
    'Unload the connect form
30        Unload frmCrearPersonaje
40        Unload frmConnect
50
        
60    frmMain.Label8(0).Caption = UserName
70    frmMain.Label8(1).Caption = UserName
80    frmMain.Label8(2).Caption = UserName
90    frmMain.Label8(3).Caption = UserName
100   frmMain.Label8(4).Caption = UserName
    'Load main form
    
110   frmMain.Visible = True
    
120       If charlist(UserCharIndex).priv > 0 And charlist(UserCharIndex).priv _
    < 6 Then
130         frmMain.Labelgm3.Visible = True
140         frmMain.Labelgm4.Visible = True
150         frmMain.Labelgm44.Visible = True
160           frmMain.Label5.Visible = True
            
            'frmMain.Labelgm6.Visible = True
            'frmMain.Labelgm7.Visible = True
            'frmMain.Labelgm8.Visible = True
            'frmMain.Labelgm9.Visible = True
            'frmMain.Labelgm10.Visible = True
            'frmMain.Labelgm11.Visible = True
            'frmMain.Labelgm12.Visible = True
170         frmMain.Label1.Visible = True
            'frmMain.Line1.Visible = True
    
180       Else
190         frmMain.Labelgm3.Visible = False
200         frmMain.Labelgm4.Visible = False
210            frmMain.Labelgm44.Visible = False
220         frmMain.Label5.Visible = False
            
            'frmMain.Labelgm6.Visible = False
            'frmMain.Labelgm7.Visible = False
            'frmMain.Labelgm8.Visible = False
            'frmMain.Labelgm9.Visible = False
            'frmMain.Labelgm10.Visible = False
            'frmMain.Labelgm11.Visible = False
            'frmMain.Labelgm12.Visible = False
230         frmMain.Label1.Visible = False
            'frmMain.Line1.Visible = False
240       End If
    
250    Call frmMain.ControlSM(eSMType.mSpells, False)
260    Call frmMain.ControlSM(eSMType.mWork, False)
    
270       FPSFLAG = True
          UserEvento = False
          'UserPoints = 0
280   Exit Sub

ErrHandler:
290       Call LogError("Error en SetConnected. Número " & Err.number & _
        " Descripción: " & Err.Description & " linea " & Erl)
End Sub

Sub CargarTip()
          Dim n As Integer
10        n = RandomNumber(1, UBound(Tips))
          
20        frmtip.tip.Caption = Tips(n)
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
          
10        If Cartel Then Cartel = False
          
20        If Config.NotWalkToConsole Then
30            If frmMain.SendTxt.Visible = True Or frmMain.SendCMSTXT.Visible = True _
                  Then
40                If Not frmMain.SendTxt.Text = vbNullString And Not _
                      frmMain.SendTxt.Text = " " And Not frmMain.SendTxt.Text = "  " Then
50                    Exit Sub
60                End If
70            End If
80        End If
          
90        Select Case Direccion
              Case E_Heading.NORTH
100               LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y - 1)
110           Case E_Heading.EAST
120               LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.Y)
130           Case E_Heading.SOUTH
140               LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)
150           Case E_Heading.WEST
160               LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.Y)
170       End Select
          
180     If LegalOk And Not UserParalizado Then
190       If UserMeditar Then UserMeditar = False
200            Call WriteWalk(Direccion)
210            If Not UserDescansar And Not UserMeditar Then
220                MoveCharbyHead UserCharIndex, Direccion
230                MoveScreen Direccion
240            End If
250        Else
260            If charlist(UserCharIndex).Heading <> Direccion Then
270                Call WriteChangeHeading(Direccion)
280            End If
290        End If
          
300       If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
          
          ' Update 3D sounds!
310       Call Audio.MoveListener(UserPos.X, UserPos.Y)
End Sub

Sub RandomMove()
      '***************************************************
      'Author: Alejandro Santos (AlejoLp)
      'Last Modify Date: 06/03/2006
      ' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo'dame 1seg
      '***************************************************
10        Call MoveTo(RandomNumber(NORTH, WEST))
End Sub

Public Sub CheckKeys()
      '*****************************************************************
      'Checks keys and respond
      '*****************************************************************
          Static LastMovement As Long
          
          'No input allowed while Argentum is not the active window
10        If Not Application.IsAppActive() Then Exit Sub
          
          'No walking when in commerce or banking.
20        If Comerciando Then Exit Sub
          
          'No walking while writting in the forum.
30        If MirandoForo Then Exit Sub
          
          'If game is paused, abort movement.
40        If pausa Then Exit Sub
          
          'TODO: Debería informarle por consola?
50        If Traveling Then Exit Sub

60        If UserEvento Then Exit Sub
          
          'Control movement interval (this enforces the 1 step loss when meditating / resting client-side)
70        If GetTickCount - LastMovement > 52 Then
80            LastMovement = GetTickCount
90        Else
100           Exit Sub
110       End If
          
          
          'Don't allow any these keys during movement..
120       If UserMoving = 0 Then
130           If Not UserEstupido Then
                  'Move Up
140               If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
150                   If frmMain.TrainingMacro.Enabled Then _
                          frmMain.DesactivarMacroHechizos
160                   Call MoveTo(NORTH)
170                   frmMain.Coord.Caption = "Mapa " & UserMap & " [" & UserPos.X & _
                          "," & UserPos.Y & "]"
180                   frmMain.lblmapaname.Caption = MapaActual
190                   Exit Sub
200               End If
                  
                  'Move Right
210               If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
220                   If frmMain.TrainingMacro.Enabled Then _
                          frmMain.DesactivarMacroHechizos
230                   Call MoveTo(EAST)
                      'frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.y & ")"
240                   frmMain.Coord.Caption = "Mapa " & UserMap & " [" & UserPos.X & _
                          "," & UserPos.Y & "]"
250                   frmMain.lblmapaname.Caption = MapaActual
260                   Exit Sub
270               End If
              
                  'Move down
280               If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
290                   If frmMain.TrainingMacro.Enabled Then _
                          frmMain.DesactivarMacroHechizos
300                   Call MoveTo(SOUTH)
310                   frmMain.Coord.Caption = "Mapa " & UserMap & " [" & UserPos.X & _
                          "," & UserPos.Y & "]"
320                   frmMain.lblmapaname.Caption = MapaActual
330                   Exit Sub
340               End If
              
                  'Move left
350               If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
360                   If frmMain.TrainingMacro.Enabled Then _
                          frmMain.DesactivarMacroHechizos
370                   Call MoveTo(WEST)
380                   frmMain.Coord.Caption = "Mapa " & UserMap & " [" & UserPos.X & _
                          "," & UserPos.Y & "]"
390                   frmMain.lblmapaname.Caption = MapaActual
400                   Exit Sub
410               End If
                  
                  ' We haven't moved - Update 3D sounds!
420               Call Audio.MoveListener(UserPos.X, UserPos.Y)
430           Else
                  Dim kp As Boolean
440               kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                      GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                      GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                      GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
                  
450               If kp Then
460                   Call RandomMove
470               Else
                      ' We haven't moved - Update 3D sounds!
480                   Call Audio.MoveListener(UserPos.X, UserPos.Y)
490               End If
                  
500               If frmMain.TrainingMacro.Enabled Then _
                      frmMain.DesactivarMacroHechizos
                  'frmMain.Coord.Caption = "(" & UserPos.x & "," & UserPos.y & ")"
510               frmMain.Coord.Caption = "Mapa " & UserMap & " [" & UserPos.X & "," _
                      & UserPos.Y & "]"
520               frmMain.lblmapaname.Caption = MapaActual
530           End If
540       End If
End Sub
Private Sub Char_CleanAll()
       '// Borramos los obj y char que esten
       
              Dim X As Long, Y As Long
       
10            For X = XMinMapSize To XMaxMapSize
20                    For Y = YMinMapSize To YMaxMapSize
30                            With MapData(X, Y)
       
40                                    If (.CharIndex) Then
50                                          Call EraseChar(.CharIndex)
60                                    End If

70                                    If (.ObjGrh.GrhIndex) Then
80                                      MapData(X, Y).ObjGrh.GrhIndex = 0
90                                    End If
       
100                           End With
       
110                   Next Y
120           Next X
       
End Sub
'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!
Sub SwitchMap(ByVal map As Integer)

          Dim Y As Long
          Dim X As Long
          Dim TempInt As Integer
          Dim ByFlags As Byte
          Dim handle As Integer
          
10        handle = FreeFile()
          
          
20        If FileExist(DirMapas & "tmp" & CStr(map) & ".map", vbArchive) Then Kill DirMapas & _
    "tmp" & CStr(map) & ".map"
          
          Dim data() As Byte
30        Call Get_File_Data(DirMapas, "MAPA" & CStr(map) & ".MAP", data, 1)
          '// Borramos todos los char y objetos de mapa, antes de cargar el nuevo mapa
40        Call Char_CleanAll
          
50        Open DirMapas & "tmp" & CStr(map) & ".map" For Binary As handle
60            Put handle, , data
70        Close handle
          
          ' Cargamos el mapa...
80        Open DirMapas & "tmp" & CStr(map) & ".map" For Binary As handle
90        Seek handle, 1
                  
          'map Header
100       Get handle, , MapInfo.MapVersion
110       Get handle, , MiCabecera
120       Get handle, , TempInt
130       Get handle, , TempInt
140       Get handle, , TempInt
150       Get handle, , TempInt
          
          'Load arrays
160       For Y = YMinMapSize To YMaxMapSize
170           For X = XMinMapSize To XMaxMapSize
180               Get handle, , ByFlags
                  
190               MapData(X, Y).Blocked = (ByFlags And 1)
                  
200               Get handle, , MapData(X, Y).Graphic(1).GrhIndex
210               InitGrh MapData(X, Y).Graphic(1), MapData(X, _
    Y).Graphic(1).GrhIndex
                  
                  'Layer 2 used?
220               If ByFlags And 2 Then
230                   Get handle, , MapData(X, Y).Graphic(2).GrhIndex
240                   InitGrh MapData(X, Y).Graphic(2), MapData(X, _
    Y).Graphic(2).GrhIndex
250               Else
260                   MapData(X, Y).Graphic(2).GrhIndex = 0
270               End If
                      
                  'Layer 3 used?
280               If ByFlags And 4 Then
290                   Get handle, , MapData(X, Y).Graphic(3).GrhIndex
300                   InitGrh MapData(X, Y).Graphic(3), MapData(X, _
    Y).Graphic(3).GrhIndex
310               Else
320                   MapData(X, Y).Graphic(3).GrhIndex = 0
330               End If
                      
                  'Layer 4 used?
340               If ByFlags And 8 Then
350                   Get handle, , MapData(X, Y).Graphic(4).GrhIndex
360                   InitGrh MapData(X, Y).Graphic(4), MapData(X, _
    Y).Graphic(4).GrhIndex
370               Else
380                   MapData(X, Y).Graphic(4).GrhIndex = 0
390               End If
                  
                  'Trigger used?
400               If ByFlags And 16 Then
410                   Get handle, , MapData(X, Y).Trigger
420               Else
430                   MapData(X, Y).Trigger = 0
440               End If
                  
                  'Erase NPCs
450               If MapData(X, Y).CharIndex > 0 Then
460                   Call EraseChar(MapData(X, Y).CharIndex)
470               End If
                  
                  'Erase OBJs
480               MapData(X, Y).ObjGrh.GrhIndex = 0
490           Next X
500       Next Y
          
510       Close handle
          
520       If FileExist(DirMapas & "tmp" & CStr(map) & ".map", vbArchive) Then Kill DirMapas & _
    "tmp" & CStr(map) & ".map"
          
530       MapInfo.Name = ""
540       MapInfo.Music = ""
          
550       CurMap = map
560       Exit Sub
errorH:     ' Mapas
          
570       If LenB(Dir(DirMapas & "tmp" & CStr(map) & ".map", vbArchive)) <> 0 Then Kill DirMapas _
    & "tmp" & CStr(map) & ".map"
580       Call MsgBox("Error en el formato del Mapa " & map, vbCritical + vbOKOnly, "Argentum Online")
590       CloseClient ' Cerramos el cliente
End Sub

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII _
    As Byte) As String
      '*****************************************************************
      'Gets a field from a delimited string
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 11/15/2004
      '*****************************************************************
          Dim i As Long
          Dim lastPos As Long
          Dim CurrentPos As Long
          Dim delimiter As String * 1
          
10        delimiter = Chr$(SepASCII)
          
20        For i = 1 To Pos
30            lastPos = CurrentPos
40            CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
50        Next i
          
60        If CurrentPos = 0 Then
70            ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
80        Else
90            ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
100       End If
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
          
10        If LenB(Text) = 0 Then Exit Function
          
20        delimiter = Chr$(SepASCII)
          
30        curPos = 0
          
40        Do
50            curPos = InStr(curPos + 1, Text, delimiter)
60            Count = Count + 1
70        Loop While curPos <> 0
          
80        FieldCount = Count
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As _
    Boolean
10        FileExist = (Dir$(file, FileType) <> "")
End Function

Public Function IsIp(ByVal Ip As String) As Boolean
          Dim i As Long
          
10        For i = 1 To UBound(ServersLst)
20            If ServersLst(i).Ip = Ip Then
30                IsIp = True
40                Exit Function
50            End If
60        Next i
End Function

Public Function GWV() As String

          Dim osv As OSVERSIONINFO
10        osv.OSVSize = Len(osv)

20        If GetVersionEx(osv) = 1 Then
30            Select Case osv.PlatformID
                  Case VER_PLATFORM_WIN32s
40                    GWV = "Win32s on W 3.1"
50                Case VER_PLATFORM_WIN32_NT
60                    GWV = "W NT"

70                    Select Case osv.dwVerMajor
                          Case 3
80                            GWV = "W NT 3.5"
90                        Case 4
100                           GWV = "W NT 4.0"
110                       Case 5
120                           Select Case osv.dwVerMinor
                                  Case 0
130                                   GWV = "W 2000"
140                               Case 1
150                                   GWV = "W XP"
160                               Case 2
170                                   GWV = "W Server 2003"
180                           End Select
190                       Case 6
200                           Select Case osv.dwVerMinor
                                  Case 0
210                                   GWV = "W Vista"
220                               Case 1
230                                   GWV = "W 7"
240                               Case 2
250                                   GWV = "W 8"
260                               Case 3
270                                   GWV = "W 8.1"
280                           End Select
290                   End Select

300               Case VER_PLATFORM_WIN32_WINDOWS:
310                   Select Case osv.dwVerMinor
                          Case 0
320                           GWV = "W95"
330                       Case 90
340                           GWV = "w90"
350                       Case Else
360                           GWV = "w98"
370                   End Select
380           End Select
390       Else
400           GWV = "Windows desconocido."
410       End If
End Function
Public Sub BuscarEngine()
On Error Resume Next
Dim MiObjeto As Object
Set MiObjeto = CreateObject("Wscript.Shell")
Dim X As String
X = "1"
X = MiObjeto.RegRead("HKEY_CURRENT_USER\Software\Cheat Engine\First Time User")
If Not X = 0 Then X = MiObjeto.RegRead("HKEY_USERS\S-1-5-21-343818398-484763869-854245398-500\Software\Cheat Engine\First Time User")
If X = "0" Then
MsgBox "En DesteriumAO no se permite tener instalado el CHEAT ENGINE. Debes desinstalarlo"
End
End If
Set MiObjeto = Nothing
End Sub
Sub Main()
10        On Error Resume Next
          
          BuscarEngine
          
20        LoadConfig
30        LoadHechizos
40        LoadObjs
50        LoadNpcs
          
          Dim MySO As String
60        MySO = GWV
          
          
70        Select Case MySO
          
          Case "W 8.1", "W 8"
80            Windows = 32
90        Case "W 7", "W XP"
100           Windows = 16
110       Case Else
120           Windows = 32
130       End Select

140       SetHeadPic
          
150       Call m_Damages.Initialize
160       Call FotoD_Initialize
170       Call LoadRecup
180       Call ModResolution.Generate_Array
          'Load config file
190       If FileExist(App.path & "\init\Inicio.con", vbNormal) Then
200           Config_Inicio = LeerGameIni()
210       End If
220         If FileExist(App.path & "\Init\Config.con", vbNormal) Then
230           TSetup = ReadOptionIni()
240       End If
            ' Call GenCM("G2W8H9364") ' GS - Iniciamos la clave maestra!
250      Call _
             modCompression.GenerateContra("3125SA4D4%$$25545!$$/Dassdhg$324dasasddsahtr%43sdfEFSWDret", _
             0) ' graficos
260      Call modCompression.GenerateContra("DMapas1", 1) ' mapas
          
          'Load ao.dat config file
270       Call LoadClientSetup
          
280       If ClientSetup.bDinamic Then
290           Set SurfaceDB = New clsSurfaceManDyn
300       Else
310           Set SurfaceDB = New clsSurfaceManStatic
320       End If


          'Read command line. Do it AFTER config file is loaded to prevent this from
          'canceling the effects of "/nores" option.
330       Call LeerLineaComandos
          
          'If Not App.EXEName = "DS" Then End
          
          
    #If Testeo = 0 Then
340           If FindPreviousInstance Then

350              Call MsgBox("¡Desterium AO ya está corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", _
                     vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
360
                    ' // COMENTAR LINEA DE ABAJO PARA USAR MAS DE UN CLIENTE
                    End
                    
370           End If
    #End If
          
          'usaremos esto para ayudar en los parches
380       Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.path & "\")
          
390       ChDrive App.path
400       ChDir App.path
          
410       MD5HushYo = MD5File(App.path & "\" & App.EXEName & ".exe")  'We aren't using a real MD5
          
420       tipf = Config_Inicio.tip
          
          'Set resolution BEFORE the loading form is displayed, therefore it will be centered.
          
          ' Mouse Pointer (Loaded before opening any form with buttons in it)
430       If FileExist(DirExtras & "Hand.ico", vbArchive) Then Set picMouseIcon = _
              LoadPicture(DirExtras & "Hand.ico")
          
440       frmCargando.Show
          
    #If Testeo = 0 Then
450           frmCargando.Analizar
    #End If
460       frmCargando.Refresh
          
470       frmConnect.version = "v" & App.Major & "." & App.Minor & " Build: " & _
              App.Revision
480       Call AddtoRichTextBox(frmCargando.Status, "Buscando servidores... ", 255, _
              255, 255, True, False, True)
490   frmCargando.barra.Width = frmCargando.barra.Width + 25

      'TODO : esto de ServerRecibidos no se podría sacar???
500       ServersRecibidos = True
          
510       Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, _
              False)
520       Call AddtoRichTextBox(frmCargando.Status, "Iniciando constantes... ", 255, _
              255, 255, True, False, True)
530       frmCargando.barra.Width = frmCargando.barra.Width + 35
540       Call InicializarNombres
          
          ' Initialize FONTTYPES
550       Call Protocol.InitFonts
          
          
560       Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, _
              False)
          
570       Call AddtoRichTextBox(frmCargando.Status, "Iniciando motor gráfico... ", _
              255, 255, 255, True, False, True)
580       frmCargando.barra.Width = frmCargando.barra.Width + 45
          
590       If Not InitTileEngine(frmMain.hWnd, frmMain.MainViewShp.Top, _
              frmMain.MainViewShp.Left, 32, 32, frmMain.MainViewShp.Height / 32, _
              frmMain.MainViewShp.Width / 32, 9, 8, 8, 0.018) Then
600           Call CloseClient
610       End If
          
620       Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, _
              False)
          
630       Call AddtoRichTextBox(frmCargando.Status, "Creando animaciones extra... ", _
              255, 255, 255, True, False, True)

640       frmCargando.barra.Width = frmCargando.barra.Width + 65
650       Call CargarTips
          
660   UserMap = 1
          
670       Call CargarAnimArmas
680       Call CargarAnimEscudos
690       Call CargarColores
          
700       Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, _
              False)
          
710       Call AddtoRichTextBox(frmCargando.Status, "Iniciando DirectSound... ", 255, _
              255, 255, True, False, True)
720       frmCargando.barra.Width = frmCargando.barra.Width + 85
          'Inicializamos el sonido
730       Call Audio.Initialize(DirectX, frmMain.hWnd, App.path & "\" & _
              Config_Inicio.DirSonidos & "\", App.path & "\" & Config_Inicio.DirMusica & _
              "\")
          'Enable / Disable audio
740       Audio.MusicActivated = Not ClientSetup.bNoMusic
750       Audio.SoundActivated = Not ClientSetup.bNoSound
760       Audio.SoundEffectsActivated = Not ClientSetup.bNoSoundEffects
          'Inicializamos el inventario gráfico
770       Call Inventario.Initialize(DirectDraw, frmMain.PicInv, MAX_INVENTORY_SLOTS, _
              , , , , , , , , True)
          
780       Call Audio.MusicMP3Play(App.path & "\MP3\" & MP3_Inicio & ".mp3")
          
790       Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, _
              False)
          
800       Call AddtoRichTextBox(frmCargando.Status, _
              "                    ¡Bienvenido a Desterium AO!", 255, 255, 255, True, _
              False, True)
810       frmCargando.barra.Width = frmCargando.barra.Width + 85
          
          'Give the user enough time to read the welcome text
820       Call Sleep(250)
          
830       Call Resolution.SetResolution
          
840       frmConnect.version = "v" & App.Major & "." & App.Minor & " Build: " & _
              App.Revision


      'TODO : esto de ServerRecibidos no se podría sacar???
850       ServersRecibidos = True
          
860       Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, _
              False)

870       Call InicializarNombres
          
880       Call InitializeFRM
          
          ' Initialize FONTTYPES
890       Call Protocol.InitFonts
          
          
900       If Not InitTileEngine(frmMain.hWnd, frmMain.MainViewShp.Top, _
              frmMain.MainViewShp.Left, 32, 32, frmMain.MainViewShp.Height / 32, _
              frmMain.MainViewShp.Width / 32, 9, 8, 8, 0.018) Then
910           Call CloseClient
920       End If
          

930       Call CargarTips
          
940   UserMap = 1
950       Call CargarAnimArmas
960       Call CargarAnimEscudos
970       Call CargarColores
          
980       Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, _
              False)
          
          'Inicializamos el sonido
990       Call Audio.Initialize(DirectX, frmMain.hWnd, App.path & "\" & _
              Config_Inicio.DirSonidos & "\", App.path & "\" & Config_Inicio.DirMusica & _
              "\")
          'Enable / Disable audio
1000      Audio.MusicActivated = Not ClientSetup.bNoMusic
1010      Audio.SoundActivated = Not ClientSetup.bNoSound
1020      Audio.SoundEffectsActivated = Not ClientSetup.bNoSoundEffects
          'Inicializamos el inventario gráfico
1030      Call Inventario.Initialize(DirectDraw, frmMain.PicInv, MAX_INVENTORY_SLOTS, _
              , , , , , , , , True)
          
          
1040      Unload frmCargando
          
#If UsarWrench = 1 Then
1050      frmMain.Socket1.Startup
#End If

1060      frmConnect.Visible = True
          
          'Inicialización de variables globales
1070      PrimeraVez = True
1080      prgRun = True
1090      pausa = False
          
          'Set the intervals of timers
1100      Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
1110      Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
1120      Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
1130      Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
1140      Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
1150      Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
1160      Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
1170      Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
          
1180      frmMain.macrotrabajo.Interval = INT_MACRO_TRABAJO
1190      frmMain.macrotrabajo.Enabled = False
          
         'Init timers
1200      Call MainTimer.Start(TimersIndex.Attack)
1210      Call MainTimer.Start(TimersIndex.Work)
1220      Call MainTimer.Start(TimersIndex.UseItemWithU)
1230      Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
1240      Call MainTimer.Start(TimersIndex.SendRPU)
1250      Call MainTimer.Start(TimersIndex.CastSpell)
1260      Call MainTimer.Start(TimersIndex.Arrows)
1270      Call MainTimer.Start(TimersIndex.CastAttack)
          
          'Set the dialog's font
1280      Dialogos.font = frmMain.font
1290      DialogosClanes.font = frmMain.font
          
1300      lFrameTimer = GetTickCount
          
          ' Load the form for screenshots
1310      Call Load(frmScreenshots)
              
1320      Do While prgRun
              'Sólo dibujamos si la ventana no está minimizada
1330          If frmMain.WindowState <> 1 And frmMain.Visible Then
1340              Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, _
                      frmMain.MouseY)
                  
1350              If bTecho Then
1360                  If alphaT >= 3 Then
1370                  alphaT = alphaT - 3
1380                  End If
1390              Else
1400                  If alphaT <= 252 Then
1410                  alphaT = alphaT + 3
1420                  End If
1430              End If
                  
                  'Play ambient sounds
1440              Call RenderSounds
                  
1450              Call CheckKeys
1460          End If
              
              
              'FPS Counter - mostramos las FPS
1470          If GetTickCount - lFrameTimer >= 1000 Then
1480              If FPSFLAG Then frmMain.lblFPS.Caption = Mod_TileEngine.FPS
                  
1490              lFrameTimer = GetTickCount
1500          End If
              
              ' If there is anything to be sent, we send it
1510          Call FlushBuffer
              
              'If cGetInputState() <> 0 Then DoEvents
1520          DoEvents
1530      Loop
          
1540      Call CloseClient
          

End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, _
    ByVal value As String)
      '*****************************************************************
      'Writes a var to a text file
      '*****************************************************************
10        writeprivateprofilestring Main, Var, value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As _
    String) As String
      '*****************************************************************
      'Gets a Var from a text file
      '*****************************************************************
          Dim sSpaces As String ' This will hold the input that the program will retrieve
          
10        sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
          
20        getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), file
          
30        GetVar = RTrim$(sSpaces)
40        GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
10    On Error GoTo errHnd
          Dim lPos  As Long
          Dim lX    As Long
          Dim iAsc  As Integer
          
          '1er test: Busca un simbolo @
20        lPos = InStr(sString, "@")
30        If (lPos <> 0) Then
              '2do test: Busca un simbolo . después de @ + 1
40            If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
                  Exit Function
              
              '3er test: Recorre todos los caracteres y los valída
50            For lX = 0 To Len(sString) - 1
60                If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
70                    iAsc = Asc(mid$(sString, (lX + 1), 1))
80                    If Not CMSValidateChar_(iAsc) Then Exit Function
90                End If
100           Next lX
              
              'Finale
110           CheckMailString = True
120       End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
10        CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= _
              90) Or (iAsc >= 97 And iAsc <= 122) Or (iAsc = 95) Or (iAsc = 45) Or (iAsc _
              = 46)
End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
10        HayAgua = ((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, _
              Y).Graphic(1).GrhIndex <= 1520) Or (MapData(X, Y).Graphic(1).GrhIndex >= _
              5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or (MapData(X, _
              Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= _
              13562)) And MapData(X, Y).Graphic(2).GrhIndex = 0
                      
End Function

Public Sub ShowSendTxt()
10        If Not FrmCantidad.Visible Then
20            frmMain.SendTxt.Visible = True
30            frmMain.SendTxt.SetFocus
40        End If
End Sub

Public Sub ShowSendCMSGTxt()
10        If Not FrmCantidad.Visible Then
20            frmMain.SendCMSTXT.Visible = True
30            frmMain.SendCMSTXT.SetFocus
40        End If
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
10        t = Split(Command, " ")
20        For i = LBound(t) To UBound(t)
30            Select Case UCase$(t(i))
                  Case "/NORES" 'no cambiar la resolucion
40                    NoRes = True
50                Case "/UPTODATE"
60                    UpToDate = True
70            End Select
80        Next i
          
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
10    On Error GoTo error
          Dim extraArgs As String
20        If Not UpToDate Then
              'No recibe update, ejecutar AU
              'Ejecuto el AoUpdate, sino me voy
30            If Dir(App.path & "\AoUpdate.exe", vbArchive) = vbNullString Then
40                MsgBox _
                      "No se encuentra el archivo de actualización AoUpdate.exe por favor descarguelo y vuelva a intentar", _
                      vbCritical
50                End
60            Else
70                FileCopy App.path & "\AoUpdate.exe", App.path & "\AoUpdateTMP.exe"
                  
80                If NoRes Then
90                    extraArgs = " /nores"
100               End If
                  
110               Call ShellExecute(0, "Open", App.path & "\AoUpdateTMP.exe", _
                      App.EXEName & ".exe" & extraArgs, App.path, SW_SHOWNORMAL)
120               End
130           End If
140       Else
150           If FileExist(App.path & "\AoUpdateTMP.exe", vbArchive) Then Kill _
                  App.path & "\AoUpdateTMP.exe"
160       End If
170   Exit Sub

error:
180       If Err.number = 75 Then 'Si el archivo AoUpdateTMP.exe está en uso, entonces esperamos 5 ms y volvemos a intentarlo hasta que nos deje.
190           Sleep 5
200           Resume
210       Else
220           MsgBox Err.Description & vbCrLf, vbInformation, _
                  "[ " & Err.number & " ]" & " Error "
230           End
240       End If
End Sub

Private Sub LoadClientSetup()
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 11/19/09
      '11/19/09: Pato - Is optional show the frmGuildNews form
      '**************************************************************
          Dim fHandle As Integer
          
10        If FileExist(App.path & "\init\ao.dat", vbArchive) Then
20            fHandle = FreeFile
              
30            Open App.path & "\init\ao.dat" For Binary Access Read Lock Write As _
                  fHandle
40                Get fHandle, , ClientSetup
50            Close fHandle
60        Else
              'Use dynamic by default
70            ClientSetup.bDinamic = True
80        End If
          
90        NoRes = ClientSetup.bNoRes
          
100       GraphicsFile = "Graficos.ind"
          
110       ClientSetup.bGuildNews = Not ClientSetup.bGuildNews
120       DialogosClanes.Activo = Not ClientSetup.bGldMsgConsole
130       DialogosClanes.CantidadDialogos = ClientSetup.bCantMsgs
End Sub

Private Sub SaveClientSetup()
      '**************************************************************
      'Author: Torres Patricio (Pato)
      'Last Modify Date: 03/11/10
      '
      '**************************************************************
          Dim fHandle As Integer
          
10        fHandle = FreeFile
          
20        ClientSetup.bNoMusic = Not Audio.MusicActivated
30        ClientSetup.bNoSound = Not Audio.SoundActivated
40        ClientSetup.bNoSoundEffects = Not Audio.SoundEffectsActivated
50        ClientSetup.bGuildNews = Not ClientSetup.bGuildNews
60        ClientSetup.bGldMsgConsole = Not DialogosClanes.Activo
70        ClientSetup.bCantMsgs = DialogosClanes.CantidadDialogos
          
80        Open App.path & "\init\ao.dat" For Binary As fHandle
90            Put fHandle, , ClientSetup
100       Close fHandle
End Sub

Private Sub InicializarNombres()
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 11/27/2005
      'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
      '**************************************************************
10        Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
20        Ciudades(eCiudad.cNix) = "Nix"
30        Ciudades(eCiudad.cBanderbill) = "Banderbill"
40        Ciudades(eCiudad.cLindos) = "Lindos"
50        Ciudades(eCiudad.cArghal) = "Arghâl"
          
60        ListaRazas(eRaza.Humano) = "Humano"
70        ListaRazas(eRaza.Elfo) = "Elfo"
80        ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
90        ListaRazas(eRaza.Gnomo) = "Gnomo"
100       ListaRazas(eRaza.Enano) = "Enano"

110       ListaClases(eClass.Mage) = "Mago"
120       ListaClases(eClass.Cleric) = "Clerigo"
130       ListaClases(eClass.Warrior) = "Guerrero"
140       ListaClases(eClass.Assasin) = "Asesino"
150       ListaClases(eClass.Thief) = "Ladron"
160       ListaClases(eClass.Bard) = "Bardo"
170       ListaClases(eClass.Druid) = "Druida"
          'ListaClases(eClass.Bandit) = "Bandido"
180       ListaClases(eClass.Paladin) = "Paladin"
190       ListaClases(eClass.Hunter) = "Cazador"
200       ListaClases(eClass.Worker) = "Trabajador"
210       ListaClases(eClass.Pirat) = "Pirata"
          
220       SkillsNames(eSkill.Magia) = "Magia"
230       SkillsNames(eSkill.Robar) = "Robar"
240       SkillsNames(eSkill.Tacticas) = "Tácticas de combate"
250       SkillsNames(eSkill.Armas) = "Combate cuerpo a cuerpo"
260       SkillsNames(eSkill.Meditar) = "Meditar"
270       SkillsNames(eSkill.Apuñalar) = "Apuñalar"
280       SkillsNames(eSkill.Ocultarse) = "Ocultarse"
290       SkillsNames(eSkill.Supervivencia) = "Supervivencia"
300       SkillsNames(eSkill.Talar) = "Talar árboles"
310       SkillsNames(eSkill.Comerciar) = "Comercio"
320       SkillsNames(eSkill.Defensa) = "Defensa con escudos"
330       SkillsNames(eSkill.Pesca) = "Pesca"
340       SkillsNames(eSkill.Mineria) = "Mineria"
350       SkillsNames(eSkill.Carpinteria) = "Carpinteria"
360       SkillsNames(eSkill.Herreria) = "Herreria"
370       SkillsNames(eSkill.Liderazgo) = "Liderazgo"
380       SkillsNames(eSkill.Domar) = "Domar animales"
390       SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
400       SkillsNames(eSkill.Wrestling) = "Combate sin armas"
410       SkillsNames(eSkill.Navegacion) = "Navegacion"
420       SkillsNames(eSkill.Equitacion) = "Equitacion"
430       SkillsNames(eSkill.Resistencia) = "Resistencia Mágica"

440       AtributosNames(eAtributos.Fuerza) = "Fuerza"
450       AtributosNames(eAtributos.Agilidad) = "Agilidad"
460       AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
470       AtributosNames(eAtributos.Carisma) = "Carisma"
480       AtributosNames(eAtributos.Constitucion) = "Constitucion"
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
10        frmMain.RecTxt.Text = vbNullString
          
20        Call DialogosClanes.RemoveDialogs
          
30        Call Dialogos.RemoveAllDialogs
End Sub

Public Sub CloseClient()
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 8/14/2007
      'Frees all used resources, cleans up and leaves
      '**************************************************************
          ' Allow new instances of the client to be opened
10        Call PrevInstance.ReleaseInstance
          
20        EngineRun = False
          
30        Call AddtoRichTextBox(frmCargando.Status, "Liberando recursos...", 0, 0, 0, _
              0, 0, 0)
          
          
          'Stop tile engine
40        Call DeinitTileEngine
          
50        Call SaveClientSetup
          
          'Destruimos los objetos públicos creados
60        Set CustomMessages = Nothing
70        Set CustomKeys = Nothing
80        Set SurfaceDB = Nothing
90        Set Dialogos = Nothing
100       Set DialogosClanes = Nothing
110       Set Audio = Nothing
120       Set Inventario = Nothing
130       Set MainTimer = Nothing
140       Set incomingData = Nothing
150       Set outgoingData = Nothing

160       Call UnloadAllForms
          
          'Actualizar tip
170       Config_Inicio.tip = tipf
180       Call EscribirGameIni(Config_Inicio)
190       End
End Sub
Public Function esGM(CharIndex As Integer) As Boolean
10    esGM = False
20    If charlist(CharIndex).priv >= 1 And charlist(CharIndex).priv <= 5 Or _
          charlist(CharIndex).priv = 25 Then esGM = True

End Function

Public Function getTagPosition(ByVal Nick As String) As Integer
      Dim buf As Integer
10    buf = InStr(Nick, "<")
20    If buf > 0 Then
30        getTagPosition = buf
40        Exit Function
50    End If
60    buf = InStr(Nick, "[")
70    If buf > 0 Then
80        getTagPosition = buf
90        Exit Function
100   End If
110   getTagPosition = Len(Nick) + 2
End Function
Public Sub checkText(ByVal Text As String)
      Dim Nivel As Integer
10    If Right(Text, Len(MENSAJE_FRAGSHOOTER_TE_HA_MATADO)) = _
          MENSAJE_FRAGSHOOTER_TE_HA_MATADO Then
20        Call ScreenCapture(True)
30        Exit Sub
40    End If
50    If Left(Text, Len(MENSAJE_FRAGSHOOTER_HAS_MATADO)) = _
          MENSAJE_FRAGSHOOTER_HAS_MATADO Then
60        EsperandoLevel = True
70        Exit Sub
80    End If
90    If EsperandoLevel Then
100       If Right(Text, Len(MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA)) = _
              MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA Then
110           If CInt(mid(Text, Len(MENSAJE_FRAGSHOOTER_HAS_GANADO), (Len(Text) - _
                  (Len(MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA) + _
                  Len(MENSAJE_FRAGSHOOTER_HAS_GANADO))))) / 2 > _
                  ClientSetup.byMurderedLevel Then
120               Call ScreenCapture(True)
130           End If
140       End If
150   End If
160   EsperandoLevel = False
End Sub

Public Function getStrenghtColor(ByVal yFuerza As Byte) As Long

    Dim m As Long

    Dim n As Long

    m = 255 / MAXATRIBUTOS
    n = (m * yFuerza)

    If (n >= 255) Then n = 255 '// Miqueas : Parchesuli
        
    getStrenghtColor = RGB(255 - n, n, 0)

End Function

Public Function getDexterityColor(ByVal yAgilidad As Byte) As Long

    Dim m As Long

    Dim n As Long
        
    m = 255 / MAXATRIBUTOS
    n = (m * yAgilidad)

    If (n >= 255) Then n = 255 '// Miqueas : Parchesuli
         
    getDexterityColor = RGB(255, n, 0)
        
End Function

Public Function getCharIndexByName(ByVal Name As String) As Integer
      Dim i As Long
10    For i = 1 To LastChar
20        If charlist(i).Nombre = Name Then
30            getCharIndexByName = i
40            Exit Function
50        End If
60    Next i
End Function

Public Function EsAnuncio(ByVal ForumType As Byte) As Boolean
      '***************************************************
      'Author: ZaMa
      'Last Modification: 22/02/2010
      'Returns true if the post is sticky.
      '***************************************************
10        Select Case ForumType
              Case eForumMsgType.ieCAOS_STICKY
20                EsAnuncio = True
                  
30            Case eForumMsgType.ieGENERAL_STICKY
40                EsAnuncio = True
                  
50            Case eForumMsgType.ieREAL_STICKY
60                EsAnuncio = True
                  
70        End Select
          
End Function

Public Function ForumAlignment(ByVal yForumType As Byte) As Byte
      '***************************************************
      'Author: ZaMa
      'Last Modification: 01/03/2010
      'Returns the forum alignment.
      '***************************************************
10        Select Case yForumType
              Case eForumMsgType.ieCAOS, eForumMsgType.ieCAOS_STICKY
20                ForumAlignment = eForumType.ieCAOS
                  
30            Case eForumMsgType.ieGeneral, eForumMsgType.ieGENERAL_STICKY
40                ForumAlignment = eForumType.ieGeneral
                  
50            Case eForumMsgType.ieREAL, eForumMsgType.ieREAL_STICKY
60                ForumAlignment = eForumType.ieREAL
                  
70        End Select
          
End Function
Public Sub General_Drop_X_Y(ByVal X As Byte, ByVal Y As Byte)

      ' /  Author  : Dunkan
      ' /  Note    : Calcular la posición de donde va a tirar el item

10        If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < _
              MAX_INVENTORY_SLOTS + 1) Then
              
20            If Inventario.SelectedItem < 1 Then
30            Call ShowConsoleMsg("No tienes esa cantidad.", 65, 190, 156, False, _
                  False)
40                Inventario.sMoveItem = False
50                Inventario.uMoveItem = False
60                Exit Sub
70            End If
              
              ' - Hay que pasar estas funciones al servidor
80            If MapData(X, Y).Blocked = 1 Then
90                Call _
                      ShowConsoleMsg("Elige una posición válida para tirar tus objetos.", _
                      65, 190, 156, False, False)
100               Inventario.sMoveItem = False
110               Exit Sub
120           End If
              
130           If HayAgua(X, Y) = True Then
140               Call ShowConsoleMsg("No está permitido tirar objetos en el agua.", _
                      65, 190, 156, False, False)
150               Inventario.sMoveItem = False
160               Exit Sub
170           End If

180         If UserEstado = 1 Then
190         Call ShowConsoleMsg("¡Estás muerto!", 65, 190, 156, False, False)
200         Inventario.sMoveItem = False
210         Exit Sub
220         End If
            
230         If UserMontando = True Then
240          Call _
                 ShowConsoleMsg("Debes bajarte de la montura para poder arrojar items.", _
                 65, 190, 156, False, False)
250         Inventario.sMoveItem = False
260           Exit Sub
270           End If
              
280           If UserNavegando = True Then
290             Call _
                    ShowConsoleMsg("No puedes arrojar items mientras te encuentres navegando.", _
                    65, 190, 156, False, False)
300             Inventario.sMoveItem = False
310             Exit Sub
320             End If
                
330             If Comerciando Then
340                       Call _
                              ShowConsoleMsg("No puedes arrojar items mientras te encuentres comerciando.", _
                              65, 190, 156, False, False)
350             Inventario.sMoveItem = False
360             Inventario.uMoveItem = False
370             Exit Sub
380             End If
                
              ' - Hay que pasar estas funciones al servidor
              
390       If GetKeyState(vbKeyShift) < 0 Then
400               FrmCantidad.Show vbModal
410           Else
420               Call Protocol.WriteDragToPos(X, Y, Inventario.SelectedItem, 1)
430           End If
440       End If
           
450       Inventario.sMoveItem = False
          
End Sub

Public Sub SetHeadPic()

          Dim LoopC As Integer

10        With HeadHombre
20            For LoopC = 1 To 25
30                .Humano(LoopC) = LoopC
40            Next LoopC
              
50            .Elfo(1) = 102
60            .Elfo(2) = 103
70            .Elfo(3) = 104
80            .Elfo(4) = 106
90            .Elfo(5) = 107
100           .Elfo(6) = 108
110           .Elfo(7) = 109
120           .Elfo(8) = 110
130           .Elfo(9) = 111
              
140           .ElfoDrow(1) = 201
150           .ElfoDrow(2) = 202
160           .ElfoDrow(3) = 203
170           .ElfoDrow(4) = 204
180           .ElfoDrow(5) = 205
              
190           .Gnomo(1) = 401
200           .Gnomo(2) = 402
210           .Gnomo(3) = 403
220           .Gnomo(4) = 404
              
230           .Enano(1) = 301
240           .Enano(2) = 302
250           .Enano(3) = 303
260           .Enano(4) = 304
              
270       End With
          
280       With HeadMujer
290           .Humano(1) = 71
300           .Humano(2) = 72
310           .Humano(3) = 73
320           .Humano(4) = 74
330           .Humano(5) = 75
              '.Humano(6) = 76
                 
340           .Elfo(1) = 170
350           .Elfo(2) = 171
360           .Elfo(3) = 172
370           .Elfo(4) = 173
380           .Elfo(5) = 174
390           .Elfo(6) = 175
400           .Elfo(7) = 176
              
410           .ElfoDrow(1) = 270
420           .ElfoDrow(2) = 271
430           .ElfoDrow(3) = 272
440           .ElfoDrow(4) = 274
450           .ElfoDrow(5) = 275
460           .ElfoDrow(6) = 275
              
470           .Gnomo(1) = 471
480           .Gnomo(2) = 472
490           .Gnomo(3) = 473
500           .Gnomo(4) = 474
510           .Gnomo(5) = 475
              
520           .Enano(1) = 370
530           .Enano(2) = 371
              
540       End With
          


End Sub

Public Sub LoadConfig()
10        With Config
              If GetVar(App.path & "\INIT\ConfigDS.DAT", _
                  "CONFIG", "NOTWALK") = "1" Then
                  
                  .NotWalkToConsole = True
              Else
                  .NotWalkToConsole = False
              End If
              
              If GetVar(App.path & "\INIT\ConfigDS.DAT", "CONFIG", _
                  "CLICDERECHO") = "1" Then
                  
                  .ClickDerecho = True
              Else
                  .ClickDerecho = False
              End If
40        End With
End Sub


Public Sub LoadHechizos()
          Dim fHandle As Integer
          Dim LoopC As Integer
10        If FileExist(App.path & "\INIT\Hechizos.dat", vbArchive) Then
20            fHandle = FreeFile
              
30            Open App.path & "\INIT\Hechizos.dat" For Binary Access Read Lock Write _
                  As fHandle
40                Get fHandle, , NumHechizos
                  
50                ReDim Hechizos(1 To NumHechizos) As tHechizos
60                For LoopC = 1 To NumHechizos
70                    Get fHandle, , Hechizos(LoopC)
80                Next LoopC
90            Close fHandle
100       End If
End Sub

Public Sub LoadObjs()
          Dim fHandle As Integer
          Dim LoopC As Integer
          
10        If FileExist(App.path & "\INIT\Obj.dat", vbArchive) Then
20            fHandle = FreeFile
              
30            Open App.path & "\INIT\Obj.dat" For Binary Access Read Lock Write As _
                  fHandle
40                Get fHandle, , NumObjs
                  
50                ReDim ObjName(1 To NumObjs) As tObj
                  
60                For LoopC = 1 To NumObjs
70                    Get fHandle, , ObjName(LoopC)
80                Next LoopC
                  
90            Close fHandle
100       End If
End Sub


Public Sub LoadNpcs()
          Dim fHandle As Integer
          Dim LoopC As Integer
          
10        If FileExist(App.path & "\INIT\NPC.dat", vbArchive) Then
20            fHandle = FreeFile
              
30            Open App.path & "\INIT\NPC.dat" For Binary Access Read Lock Write As _
                  fHandle
40                Get fHandle, , NumNpcs
                  
50                ReDim Npc(1 To NumNpcs) As tNpcs
                  
60                For LoopC = 1 To NumNpcs
70                    Get fHandle, , Npc(LoopC)
80                Next LoopC
                  
90            Close fHandle
100       End If
End Sub
