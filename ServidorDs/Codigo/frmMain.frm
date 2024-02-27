VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desterium AO"
   ClientHeight    =   5700
   ClientLeft      =   1950
   ClientTop       =   1815
   ClientWidth     =   9765
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5700
   ScaleWidth      =   9765
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Caption         =   "Cargas rápidas"
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   5160
      TabIndex        =   12
      Top             =   360
      Width           =   3180
      Begin VB.Label lblUpdateAdmin 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cargar GMS/RANGOS"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   150
         TabIndex        =   13
         Top             =   255
         Width           =   2895
      End
   End
   Begin VB.Timer packetResend 
      Interval        =   10
      Left            =   480
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   1080
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3960
      Top             =   4920
   End
   Begin VB.TextBox txtChat 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1920
      Width           =   4935
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   480
      Top             =   540
   End
   Begin VB.CheckBox SUPERLOG 
      Caption         =   "log"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton CMDDUMP 
      Caption         =   "dump"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   1020
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1440
      Top             =   60
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   480
      Top             =   1080
   End
   Begin VB.Timer npcataca 
      Enabled         =   0   'False
      Interval        =   9000
      Left            =   1920
      Top             =   1020
   End
   Begin VB.Timer KillLog 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1920
      Top             =   60
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   1935
      Top             =   540
   End
   Begin VB.Frame Frame1 
      Caption         =   "BroadCast"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4935
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Broadcast consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Broadcast clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Escuch 
      Caption         =   "Label2"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de usuarios:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1725
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   2520
      TabIndex        =   0
      Top             =   5040
      Width           =   45
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Argentum"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray Servidor"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar Servidor"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONUP = &H205

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, Id As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
          Dim nidTemp As NOTIFYICONDATA

10        nidTemp.cbSize = Len(nidTemp)
20        nidTemp.hWnd = hWnd
30        nidTemp.uID = Id
40        nidTemp.uFlags = flags
50        nidTemp.uCallbackMessage = CallbackMessage
60        nidTemp.hIcon = Icon
70        nidTemp.szTip = Tip & Chr$(0)

80        setNOTIFYICONDATA = nidTemp
End Function

Sub CheckIdleUser()
          Dim iUserIndex As Long
          
10        For iUserIndex = 1 To MaxUsers
20            With UserList(iUserIndex)
                  'Conexion activa? y es un usuario loggeado?
30                If .ConnID <> -1 And .flags.UserLogged Then
                      'Actualiza el contador de inactividad
40                    If .flags.Traveling = 0 Then
50                        .Counters.IdleCount = .Counters.IdleCount + 1
60                    End If
                      
70                    If .Counters.IdleCount >= IdleLimit Then
80                        Call WriteShowMessageBox(iUserIndex, "Demasiado tiempo inactivo. Has sido desconectado.")
                          'mato los comercios seguros
90                        If .ComUsu.DestUsu > 0 Then
100                           If UserList(.ComUsu.DestUsu).flags.UserLogged Then
110                               If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
120                                   Call WriteConsoleMsg(.ComUsu.DestUsu, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
130                                   Call FinComerciarUsu(.ComUsu.DestUsu)
140                                   Call FlushBuffer(.ComUsu.DestUsu) 'flush the buffer to send the message right away
150                               End If
160                           End If
170                           Call FinComerciarUsu(iUserIndex)
180                       End If
190                       Call Cerrar_Usuario(iUserIndex)
200                   End If
210               End If
220           End With
230       Next iUserIndex
End Sub

Private Sub Auditoria_Timer()
10    On Error GoTo errhand
      Static centinelSecs As Byte
      Static Recover As Byte
      
20    centinelSecs = centinelSecs + 1

30    If centinelSecs = 5 Then
          'Every 5 seconds, we try to call the player's attention so it will report the code.
40        Call modCentinela.CallUserAttention
          
50        centinelSecs = 0
60    End If

70    Call PasarSegundo 'sistema de desconexion de 10 segs

      If Recover = 30 Then
            Call CheckRecoverPasswd
            Recover = 0
      End If
      
      'Call ActualizaEstadisticasWeb

80    Exit Sub

errhand:

90    Call LogError("Error en Timer Auditoria. Err: " & Err.Description & " - " & Err.Number)
100   Resume Next

End Sub

Private Sub AutoSave_Timer()

10    On Error GoTo Errhandler
      'fired every minute
      Static Minutos As Long
      Static MinutosLatsClean As Long
      Static MinsPjesSave As Long
      Static MinutosGranPoder As Long
      
      Dim i As Integer
      Dim Num As Long

      
20    Minutos = Minutos + 1
30    MinutosGranPoder = MinutosGranPoder + 1
40    MinsPjesSave = MinsPjesSave + 1

50    If MinutosGranPoder >= 15 Then
60        LoopDias
70        MinutosGranPoder = 0
80    End If

      '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
90    Call ModAreas.AreasOptimizacion
      '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

      'Actualizamos el centinela
100   Call modCentinela.PasarMinutoCentinela

110   If Minutos = MinutosWs - 1 Then
120       Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Worldsave y limpieza del mundo en 1 minuto ...", FontTypeNames.FONTTYPE_VENENO))
130   End If

140   If Minutos >= MinutosWs Then
150       Call ES.DoBackUp
160       Call aClon.VaciarColeccion
170       Minutos = 0
180   End If

190   If MinsPjesSave = MinutosGuardarUsuarios - 1 Then
         ' Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("CharSave en 1 minuto ...", FontTypeNames.FONTTYPE_VENENO))
200   End If
210   If MinsPjesSave >= MinutosGuardarUsuarios Then
          'Call mdParty.ActualizaExperiencias
220       Call BackupUsers
230       MinsPjesSave = 0
240   End If

250   If MinutosLatsClean >= 15 Then
260       MinutosLatsClean = 0
270       Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
280   Else
290       MinutosLatsClean = MinutosLatsClean + 1
300   End If

310   Call PurgarPenas
320   Call CheckIdleUser

      '<<<<<-------- Log the number of users online ------>>>
      Dim N As Integer
330   N = FreeFile()
340   Open App.Path & "\logs\numusers.log" For Output Shared As N
350   Print #N, NumUsers
360   Close #N
      '<<<<<-------- Log the number of users online ------>>>

370   Exit Sub
Errhandler:
380       Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.Description)
390       Resume Next
End Sub

Private Sub CMDDUMP_Click()
10    On Error Resume Next

      Dim i As Integer
20    For i = 1 To MaxUsers
30        Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).Name & " UserLogged: " & UserList(i).flags.UserLogged)
40    Next i

50    Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

Private Sub Command1_Click()
10    Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(BroadMsg.Text))
      ''''''''''''''''SOLO PARA EL TESTEO'''''''
      ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
20    txtChat.Text = txtChat.Text & vbNewLine & "Servidor> " & BroadMsg.Text
End Sub

Public Sub InitMain(ByVal f As Byte)

10    If f = 1 Then
20        Call mnuSystray_Click
30    Else
40        frmMain.Show
50    End If

End Sub

Private Sub Command2_Click()
10    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & BroadMsg.Text, FontTypeNames.FONTTYPE_SERVER))
      ''''''''''''''''SOLO PARA EL TESTEO'''''''
      ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
20    txtChat.Text = txtChat.Text & vbNewLine & "Servidor> " & BroadMsg.Text
End Sub

Private Sub Command4_Click()

        'Exit Sub
          'Dim UserIndex As Integer
          
10        'UserIndex = NameIndex("LAUTARO")
          
20        'With UserList(UserIndex)
30            'ChangeUserChar NameIndex("LAUTARO"), txtCara.Text, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, False
40        'End With
End Sub

Private Sub Command3_Click()
    Call Protocol.WriteInfoPj(NameIndex("ARTIC"), "LAUTARO")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
10    On Error Resume Next
         
20       If Not Visible Then
30            Select Case X \ Screen.TwipsPerPixelX
                      
                  Case WM_LBUTTONDBLCLK
40                    WindowState = vbNormal
50                    Visible = True
                      Dim hProcess As Long
60                    GetWindowThreadProcessId hWnd, hProcess
70                    AppActivate hProcess
80                Case WM_RBUTTONUP
90                    hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
100                   PopupMenu mnuPopUp
110                   If hHook Then UnhookWindowsHookEx hHook: hHook = 0
120           End Select
130      End If
         
End Sub

Private Sub QuitarIconoSystray()
10    On Error Resume Next

      'Borramos el icono del systray
      Dim i As Integer
      Dim nid As NOTIFYICONDATA

20    nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

30    i = Shell_NotifyIconA(NIM_DELETE, nid)
          

End Sub

Private Sub Form_Unload(Cancel As Integer)
10    On Error Resume Next

      'Save stats!!!
      'Call Statistics.DumpStatistics

20    Call QuitarIconoSystray

#If UsarQueSocket = 1 Then
30    Call LimpiaWsApi
#ElseIf UsarQueSocket = 0 Then
40    Socket1.Cleanup
#ElseIf UsarQueSocket = 2 Then
50    Serv.Detener
#End If

      Dim LoopC As Integer

60    For LoopC = 1 To MaxUsers
70        If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
80    Next

      'Log
      Dim N As Integer
90    N = FreeFile
100   Open App.Path & "\logs\Main.log" For Append Shared As #N
110   Print #N, Date & " " & time & " server cerrado."
120   Close #N

130   End

140   Set SonidosMapas = Nothing

End Sub

Private Sub FX_Timer()
10    On Error GoTo hayerror

20    Call SonidosMapas.ReproducirSonidosDeMapas

30    Exit Sub
hayerror:

End Sub

Private Sub GameTimer_Timer()
      '********************************************************
      'Author: Unknown
      'Last Modify Date: -
      '********************************************************
          Dim iUserIndex As Long
          Dim bEnviarStats As Boolean
          Dim bEnviarAyS As Boolean
          
10    On Error GoTo hayerror
          
          '<<<<<< Procesa eventos de los usuarios >>>>>>
20        For iUserIndex = 1 To LastUser 'LastUser
30            With UserList(iUserIndex)
                 mIntervalos.LoopInterval iUserIndex
                 
                 'Conexion activa?
40               If .ConnID <> -1 Then
                      '¿User valido?
                      
50                    If .ConnIDValida And .flags.UserLogged Then
                          
                          '[Alejo-18-5]
60                        bEnviarStats = False
70                        bEnviarAyS = False
                          
                          
80                        If .flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
90                        If .flags.Ceguera = 1 Or .flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
                          
100                       If .Counters.TimePotFull > 0 Then
110                           .Counters.TimePotFull = .Counters.TimePotFull - 1
                              
                              ' El conteo llego a 0 y el usuario no siguió poteando.
120                           If .Counters.TimePotFull = 0 Then
130                               .FailedPot = .FailedPot + 1
                                  
140                               If .FailedPot = 5 Then
150                                   LogAntiCheat "El personaje " & .Name & " puede tener DLL."
160                                   .FailedPot = 0
170                                   .PotFull = False
180                               End If
                                  
190                           End If
200                       End If
                          
210                       Call Mod_AntiCheat.RestoTiempo(iUserIndex)
                          
220                       If .flags.Muerto = 0 Then
                              
230                            If (.flags.Privilegios And PlayerType.User) And (.Pos.map = 169 Or .Pos.map = 170 Or .Pos.map = 171) And Not (.Char.body = 171 Or .Char.body = 172 Or .Char.body = 173) Then Call General.elefectofrio(iUserIndex)
                              
                              '[Consejeros]
                              'If (.flags.Privilegios And PlayerType.User) Then Call EfectoLava(iUserIndex)
                              
240                           If .flags.Desnudo <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoFrio(iUserIndex)
                              
250                           If .flags.Meditando Then Call DoMeditar(iUserIndex)
                              
260                           If .flags.Envenenado <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoVeneno(iUserIndex)
                              
270                           If .flags.AdminInvisible <> 1 Then
280                               If .flags.invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
290                               If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)
300                           End If
                              
310                           If .flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)
                              
320                           If .flags.AtacablePor > 0 Then Call EfectoEstadoAtacable(iUserIndex)
                              
330                           Call DuracionPociones(iUserIndex)
                              
340                           Call HambreYSed(iUserIndex, bEnviarAyS)
                              
350                           If .flags.Hambre = 0 And .flags.Sed = 0 Then
360                               If Lloviendo Then
370                                   If Not Intemperie(iUserIndex) Then
380                                       If Not .flags.Descansar Then
                                          'No esta descansando
390                                           Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
400                                           If bEnviarStats Then
410                                               Call WriteUpdateHP(iUserIndex)
420                                               Call WriteUpdateFollow(iUserIndex)
430                                               bEnviarStats = False
440                                           End If
450                                           Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
460                                           If bEnviarStats Then
470                                               Call WriteUpdateSta(iUserIndex)
480                                               bEnviarStats = False
490                                           End If
500                                       Else
                                          'esta descansando
510                                           Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
520                                           If bEnviarStats Then
530                                               Call WriteUpdateHP(iUserIndex)
540                                               Call WriteUpdateFollow(iUserIndex)
550                                               bEnviarStats = False
560                                           End If
570                                           Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
580                                           If bEnviarStats Then
590                                               Call WriteUpdateSta(iUserIndex)
600                                               bEnviarStats = False
610                                           End If
                                              'termina de descansar automaticamente
620                                           If .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
630                                               Call WriteRestOK(iUserIndex)
640                                               Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
650                                               .flags.Descansar = False
660                                           End If
                                              
670                                       End If
680                                   End If
690                               Else
700                                   If Not .flags.Descansar Then
                                      'No esta descansando
                                          
710                                       Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
720                                       If bEnviarStats Then
730                                           Call WriteUpdateHP(iUserIndex)
740                                           Call WriteUpdateFollow(iUserIndex)
750                                           bEnviarStats = False
760                                       End If
770                                       Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
780                                       If bEnviarStats Then
790                                           Call WriteUpdateSta(iUserIndex)
800                                           bEnviarStats = False
810                                       End If
                                          
820                                   Else
                                      'esta descansando
                                          
830                                       Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
840                                       If bEnviarStats Then
850                                           Call WriteUpdateHP(iUserIndex)
860                                           Call WriteUpdateFollow(iUserIndex)
870                                           bEnviarStats = False
880                                       End If
890                                       Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
900                                       If bEnviarStats Then
910                                           Call WriteUpdateSta(iUserIndex)
920                                           bEnviarStats = False
930                                       End If
                                          'termina de descansar automaticamente
940                                       If .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
950                                           Call WriteRestOK(iUserIndex)
960                                           Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
970                                           .flags.Descansar = False
980                                       End If
                                          
990                                   End If
1000                              End If
1010                          End If
                              
1020                          If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)
                              
1030                          If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex)
1040                      End If 'Muerto
1050                  Else 'no esta logeado?
                          'Inactive players will be removed!
1060                      .Counters.IdleCount = .Counters.IdleCount + 1
1070                      If .Counters.IdleCount > IntervaloParaConexion Then
1080                          .Counters.IdleCount = 0
1090                          Call CloseSocket(iUserIndex)
1100                      End If
1110                  End If 'UserLogged
                      
                      'If there is anything to be sent, we send it
1120                  Call FlushBuffer(iUserIndex)
1130              End If
1140          End With
1150      Next iUserIndex
1160  Exit Sub

hayerror:
1170      LogError ("Error en GameTimer: " & Err.Description & " UserIndex = " & iUserIndex)
End Sub



Private Sub lblUpdateAdmin_Click()
    loadAdministrativeUsers
End Sub

Private Sub mnuCerrar_Click()


10    If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
          Dim f
20        For Each f In Forms
30            Unload f
40        Next
50    End If

End Sub

Private Sub mnusalir_Click()
10        Call mnuCerrar_Click
End Sub

Public Sub mnuMostrar_Click()
10    On Error Resume Next
20        WindowState = vbNormal
30        Form_MouseMove 0, 0, 7725, 0
End Sub

Private Sub KillLog_Timer()
10    On Error Resume Next
20    If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
30    If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
40    If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
50    If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
60    If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
70    If Not FileExist(App.Path & "\logs\nokillwsapi.txt") Then
80        If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"
90    End If

End Sub

Private Sub mnuServidor_Click()
10    frmServidor.Visible = True
End Sub

Private Sub mnuSystray_Click()

      Dim i As Integer
      Dim S As String
      Dim nid As NOTIFYICONDATA

10    S = "ARGENTUM-ONLINE"
20    nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
30    i = Shell_NotifyIconA(NIM_ADD, nid)
          
40    If WindowState <> vbMinimized Then WindowState = vbMinimized
50    Visible = False

End Sub

Private Sub npcataca_Timer()

10    On Error Resume Next
      Dim Npc As Long

20    For Npc = 1 To LastNPC
30        Npclist(Npc).CanAttack = 1
40    Next Npc

End Sub

Private Sub packetResend_Timer()
      '***************************************************
      'Autor: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 04/01/07
      'Attempts to resend to the user all data that may be enqueued.
      '***************************************************
On Error GoTo Errhandler:
          Dim i As Long
          
10        For i = 1 To MaxUsers
20            If UserList(i).ConnIDValida Then
30                If UserList(i).outgoingData.length > 0 Then
40                    Call EnviarDatosASlot(i, UserList(i).outgoingData.ReadASCIIStringFixed(UserList(i).outgoingData.length))
50                End If
60            End If
70        Next i

80    Exit Sub

Errhandler:
90        LogError ("Error en packetResend - Error: " & Err.Number & " - Desc: " & Err.Description)
100       Resume Next
End Sub

Private Sub TIMER_AI_Timer()

10    On Error GoTo ErrorHandler
      Dim NpcIndex As Long
      Dim X As Integer
      Dim Y As Integer
      Dim UseAI As Integer
      Dim Mapa As Integer
      Dim e_p As Integer

      'Barrin 29/9/03
20    If Not haciendoBK And Not EnPausa Then
          'Update NPCs
30        For NpcIndex = 1 To LastNPC
              
40            With Npclist(NpcIndex)
50                If .flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
                  
                      ' Chequea si contiua teniendo dueño
60                    If .Owner > 0 Then Call ValidarPermanenciaNpc(NpcIndex)
                  
70                    If .flags.Paralizado = 1 Then
80                        Call EfectoParalisisNpc(NpcIndex)
90                    Else
100                       e_p = esPretoriano(NpcIndex)
110                       If e_p > 0 Then
120                           Select Case e_p
                                  Case 1  ''clerigo
130                                   Call PRCLER_AI(NpcIndex)
140                               Case 2  ''mago
150                                   Call PRMAGO_AI(NpcIndex)
160                               Case 3  ''cazador
170                                   Call PRCAZA_AI(NpcIndex)
180                               Case 4  ''rey
190                                   Call PRREY_AI(NpcIndex)
200                               Case 5  ''guerre
210                                   Call PRGUER_AI(NpcIndex)
220                           End Select
230                       Else
                              'Usamos AI si hay algun user en el mapa
240                           If .flags.Inmovilizado = 1 Then
250                              Call EfectoParalisisNpc(NpcIndex)
260                           End If
                              
270                           Mapa = .Pos.map
                              
280                           If Mapa > 0 Then
290                               If MapInfo(Mapa).NumUsers > 0 Then
300                                   If .Movement <> TipoAI.ESTATICO Then
310                                       Call NPCAI(NpcIndex)
320                                   End If
330                               End If
340                           End If
350                       End If
360                   End If
370               End If
380           End With
390       Next NpcIndex
400   End If

410   Exit Sub

ErrorHandler:
420       Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).Pos.map)
430       Call MuereNpc(NpcIndex, 0)
End Sub

Private Sub tPiqueteC_Timer()
          Dim NuevaA As Boolean
         ' Dim NuevoL As Boolean
          Dim GI As Integer
          
          Dim i As Long
          
10    On Error GoTo Errhandler
20        For i = 1 To LastUser
30            With UserList(i)
40                If .flags.UserLogged And .ConnID >= 0 And .ConnIDValida Then
50                    If InMapBounds(.Pos.map, .Pos.X, .Pos.Y) Then
60                        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ANTIPIQUETE Then
70                            .Counters.PiqueteC = .Counters.PiqueteC + 1
80                            Call WriteConsoleMsg(i, "¡¡¡Estás obstruyendo la vía pública, muévete o serás encarcelado!!!", FontTypeNames.FONTTYPE_INFO)
                              
90                            If .Counters.PiqueteC > 23 Then
100                               .Counters.PiqueteC = 0
110                               Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
120                           End If
130                       Else
140                           .Counters.PiqueteC = 0
150                       End If
160                   End If
                      
                      'ustedes se preguntaran que hace esto aca?
                      'bueno la respuesta es simple: el codigo de AO es una mierda y encontrar
                      'todos los puntos en los cuales la alineacion puede cambiar es un dolor de
                      'huevos, asi que lo controlo aca, cada 6 segundos, lo cual es razonable
              
170                   GI = .GuildIndex
180                   If GI > 0 Then
190                       NuevaA = False
                         ' NuevoL = False
200                       If Not modGuilds.m_ValidarPermanencia(i, True, NuevaA) Then
210                           Call WriteConsoleMsg(i, "Has sido expulsado del clan. ¡El clan ha sumado un punto de antifacción!", FontTypeNames.FONTTYPE_GUILD)
220                       End If
230                       If NuevaA Then
240                           Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg("¡El clan ha pasado a tener alineación " & GuildAlignment(GI) & "!", FontTypeNames.FONTTYPE_GUILD))
250                           Call LogClanes("¡El clan cambio de alineación!")
260                       End If
      '                    If NuevoL Then
      '                        Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg("¡El clan tiene un nuevo líder!", FontTypeNames.FONTTYPE_GUILD))
      '                        Call LogClanes("¡El clan tiene nuevo lider!")
      '                    End If
270                   End If
                      
280                   Call FlushBuffer(i)
290               End If
300           End With
310       Next i
320   Exit Sub

Errhandler:
330       Call LogError("Error en tPiqueteC_Timer " & Err.Number & ": " & Err.Description)
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''USO DEL CONTROL TCPSERV'''''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#If UsarQueSocket = 3 Then

Private Sub TCPServ_Eror(ByVal Numero As Long, ByVal Descripcion As String)
10        Call LogError("TCPSERVER SOCKET ERROR: " & Numero & "/" & Descripcion)
End Sub

Private Sub TCPServ_NuevaConn(ByVal Id As Long)
10    On Error GoTo errorHandlerNC

20        ESCUCHADAS = ESCUCHADAS + 1
30        Escuch.Caption = ESCUCHADAS
          
          Dim i As Integer
          
          Dim NewIndex As Integer
40        NewIndex = NextOpenUser
          
50        If NewIndex <= MaxUsers Then
              'call logindex(NewIndex, "******> Accept. ConnId: " & ID)
              
60            TCPServ.SetDato Id, NewIndex
              
70            If aDos.MaxConexiones(TCPServ.GetIP(Id)) Then
80                Call aDos.RestarConexion(TCPServ.GetIP(Id))
90                Call ResetUserSlot(NewIndex)
100               Exit Sub
110           End If

120   If aDos.MaxConexiones(UserList(NewIndex).ip) Then
130           UserList(NewIndex).ConnID = -1
140           Call aDos.RestarConexion(UserList(NewIndex).ip)
150           Call apiclosesocket(NuevoSock)
160       End If

170           If NewIndex > LastUser Then LastUser = NewIndex

180           UserList(NewIndex).ConnID = Id
190           UserList(NewIndex).ip = TCPServ.GetIP(Id)
200           UserList(NewIndex).ConnIDValida = True
210           Set UserList(NewIndex).CommandsBuffer = New CColaArray
              
220           For i = 1 To BanIps.Count
230               If BanIps.Item(i) = TCPServ.GetIP(Id) Then
240                   Call ResetUserSlot(NewIndex)
250                   Exit Sub
260               End If
270           Next i

280       Else
290           Call CloseSocket(NewIndex, True)
300           LogCriticEvent ("NEWINDEX > MAXUSERS. IMPOSIBLE ALOCATEAR SOCKETS")
310       End If

320   Exit Sub

errorHandlerNC:
330   Call LogError("TCPServer::NuevaConexion " & Err.Description)
End Sub

Private Sub TCPServ_Close(ByVal Id As Long, ByVal MiDato As Long)
10        On Error GoTo eh
          '' No cierro yo el socket. El on_close lo cierra por mi.
          'call logindex(MiDato, "******> Remote Close. ConnId: " & ID & " Midato: " & MiDato)
20        Call CloseSocket(MiDato, False)
30    Exit Sub
eh:
40        Call LogError("Ocurrio un error en el evento TCPServ_Close. ID/miDato:" & Id & "/" & MiDato)
End Sub

Private Sub TCPServ_Read(ByVal Id As Long, Datos As Variant, ByVal Cantidad As Long, ByVal MiDato As Long)
10    On Error GoTo errorh

20    With UserList(MiDato)
30        Datos = StrConv(StrConv(Datos, vbUnicode), vbFromUnicode)
          
40        Call .incomingData.WriteASCIIStringFixed(Datos)
          
50        If .ConnID <> -1 Then
              
60            While UserList(Index).incomingData.length And HandleIncomingData(MiDato)
70            Wend
80        Else
90            Exit Sub
100       End If
110   End With

120   Exit Sub

errorh:
130   Call LogError("Error socket read: " & MiDato & " dato:" & RD & " userlogged: " & UserList(MiDato).flags.UserLogged & " connid:" & UserList(MiDato).ConnID & " ID Parametro" & Id & " error:" & Err.Description)

End Sub

#End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FIN  USO DEL CONTROL TCPSERV'''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub txtChat_Change()

End Sub
