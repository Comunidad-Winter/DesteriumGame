VERSION 5.00
Begin VB.Form frmServidor 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Servidor"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   469
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   311
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command29 
      Caption         =   "Invasiones"
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
      Left            =   1125
      TabIndex        =   33
      Top             =   6225
      Width           =   1080
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Eventos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3270
      TabIndex        =   32
      Top             =   6180
      Width           =   975
   End
   Begin VB.CommandButton cmbRetos 
      Caption         =   "Retos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3270
      TabIndex        =   31
      Top             =   6405
      Width           =   975
   End
   Begin VB.CommandButton cmbUpdateCanje 
      Caption         =   "Canjes"
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
      Left            =   240
      TabIndex        =   30
      Top             =   6240
      Width           =   810
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Reload Ranking"
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
      Left            =   240
      TabIndex        =   29
      Top             =   0
      Width           =   4095
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Reset Listen"
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
      Left            =   2940
      TabIndex        =   26
      Top             =   6675
      Width           =   1455
   End
   Begin VB.PictureBox picFuera 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4350
      Left            =   120
      ScaleHeight     =   288
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   6
      Top             =   120
      Width           =   4590
      Begin VB.VScrollBar VS1 
         Height          =   4335
         LargeChange     =   50
         Left            =   4320
         SmallChange     =   17
         TabIndex        =   24
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picCont 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   4815
         Left            =   0
         ScaleHeight     =   321
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   289
         TabIndex        =   7
         Top             =   0
         Width           =   4334
         Begin VB.CommandButton Command27 
            Caption         =   "Debug UserList"
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
            TabIndex        =   27
            Top             =   4440
            Width           =   4095
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Administración"
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
            TabIndex        =   8
            Top             =   4200
            Width           =   4095
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Pausar el servidor"
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
            TabIndex        =   9
            Top             =   3960
            Width           =   4095
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Actualizar npcs.dat"
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
            TabIndex        =   10
            Top             =   3720
            Width           =   4095
         End
         Begin VB.CommandButton Command25 
            Caption         =   "Reload MD5s"
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
            TabIndex        =   25
            Top             =   3480
            Width           =   4095
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Reload Server.ini"
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
            TabIndex        =   11
            Top             =   3240
            Width           =   4095
         End
         Begin VB.CommandButton Command28 
            Caption         =   "Reload Balance.dat"
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
            TabIndex        =   28
            Top             =   3000
            Width           =   4095
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Update MOTD"
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
            TabIndex        =   12
            Top             =   2760
            Width           =   4095
         End
         Begin VB.CommandButton Command19 
            Caption         =   "Unban All IPs (PELIGRO!)"
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
            TabIndex        =   13
            Top             =   2520
            Width           =   4095
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Unban All (PELIGRO!)"
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
            TabIndex        =   14
            Top             =   2280
            Width           =   4095
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Debug Npcs"
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
            TabIndex        =   15
            Top             =   2040
            Width           =   4095
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Stats de los slots"
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
            TabIndex        =   16
            Top             =   1800
            Width           =   4095
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Trafico"
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
            TabIndex        =   17
            Top             =   1560
            Width           =   4095
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Reload Lista Nombres Prohibidos"
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
            TabIndex        =   18
            Top             =   1320
            Width           =   4095
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Actualizar hechizos"
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
            TabIndex        =   19
            Top             =   1080
            Width           =   4095
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Configurar intervalos"
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
            TabIndex        =   20
            Top             =   840
            Width           =   4095
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Reiniciar"
            Enabled         =   0   'False
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
            TabIndex        =   21
            Top             =   600
            Width           =   4095
         End
         Begin VB.CommandButton Command6 
            Caption         =   "ReSpawn Guardias en posiciones originales"
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
            TabIndex        =   22
            Top             =   360
            Width           =   4095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Actualizar objetos.dat"
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
            TabIndex        =   23
            Top             =   120
            Width           =   4095
         End
      End
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Boton Magico para apagar server"
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
      Left            =   240
      TabIndex        =   5
      Top             =   5520
      Width           =   4095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cargar BackUp del mundo"
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
      Left            =   240
      TabIndex        =   1
      Top             =   5160
      Width           =   4095
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Guardar todos los personajes"
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
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hacer un Backup del mundo"
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
      Left            =   240
      TabIndex        =   2
      Top             =   4680
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   1860
      TabIndex        =   0
      Top             =   6675
      Width           =   945
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Reset sockets"
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
      Left            =   135
      TabIndex        =   4
      Top             =   6675
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Reload"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   600
      Index           =   0
      Left            =   135
      TabIndex        =   34
      Top             =   6015
      Width           =   2925
      Begin VB.CommandButton Command30 
         Caption         =   "Cofres"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2205
         TabIndex        =   36
         Top             =   255
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Desactivar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Index           =   1
      Left            =   3150
      TabIndex        =   35
      Top             =   6015
      Width           =   1320
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   120
      Top             =   4560
      Width           =   4335
   End
End
Attribute VB_Name = "frmServidor"
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

Private Sub cmbRetos_Click()
10        If RetosActivos Then
20            RetosActivos = False
30            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Retos» Desactivados.", FontTypeNames.FONTTYPE_VENENO)
40        Else
50            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Retos» Activados.", FontTypeNames.FONTTYPE_VENENO)
60            RetosActivos = True
70        End If
End Sub

Private Sub cmbUpdateCanje_Click()
10        LoadCanjes
20        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> Archivo #Canjes.dat ha sido actualizado.", FontTypeNames.FONTTYPE_SERVER))
          
End Sub

Private Sub cmdRecargarAdministradores_Click()

End Sub

Private Sub Command1_Click()
10    Call ResetForums
20    Call LoadOBJData

End Sub

Private Sub Command10_Click()
10    frmTrafic.Show
End Sub

Private Sub Command11_Click()
10    frmConID.Show
End Sub

Private Sub Command12_Click()
10    frmDebugNpc.Show
End Sub






Private Sub Command15_Click()
10    On Error Resume Next

      Dim Fn As String
      Dim cad$
      Dim N As Integer, k As Integer

      Dim sENtrada As String

20    sENtrada = InputBox("Escribe ""estoy DE acuerdo"" entre comillas y con distinción de mayúsculas minúsculas para desbanear a todos los personajes.", "UnBan", "hola")
30    If sENtrada = "estoy DE acuerdo" Then

40        Fn = App.Path & "\logs\GenteBanned.log"
          
50        If FileExist(Fn, vbNormal) Then
60            N = FreeFile
70            Open Fn For Input Shared As #N
80            Do While Not EOF(N)
90                k = k + 1
100               Input #N, cad$
110               Call UnBan(cad$)
                  
120           Loop
130           Close #N
140           MsgBox "Se han habilitado " & k & " personajes."
150           Kill Fn
160       End If
170   End If

End Sub

Private Sub Command16_Click()
10    Call LoadSini
End Sub

Private Sub Command17_Click()
10        Call CargaNpcsDat
End Sub

Private Sub Command18_Click()
10    Me.MousePointer = 11
20    Call mGroup.DistributeExpAndGldGroups
30    Call GuardarUsuarios
40    Me.MousePointer = 0
50    MsgBox "Grabado de personajes OK!"
End Sub

Private Sub Command19_Click()
      Dim i As Long, N As Long

      Dim sENtrada As String

10    sENtrada = InputBox("Escribe ""estoy DE acuerdo"" sin comillas y con distinción de mayúsculas minúsculas para desbanear a todos los personajes", "UnBan", "hola")
20    If sENtrada = "estoy DE acuerdo" Then
          
30        N = BanIps.Count
40        For i = 1 To BanIps.Count
50            BanIps.Remove 1
60        Next i
          
70        MsgBox "Se han habilitado " & N & " ipes"
80    End If

End Sub

Private Sub Command2_Click()
10    frmServidor.Visible = False
End Sub

Private Sub Command20_Click()
#If UsarQueSocket = 1 Then

10    If MsgBox("¿Está seguro que desea reiniciar los sockets? Se cerrarán todas las conexiones activas.", vbYesNo, "Reiniciar Sockets") = vbYes Then
20        Call WSApiReiniciarSockets
30    End If

#ElseIf UsarQueSocket = 2 Then

      Dim LoopC As Integer

40    If MsgBox("¿Está seguro que desea reiniciar los sockets? Se cerrarán todas las conexiones activas.", vbYesNo, "Reiniciar Sockets") = vbYes Then
50        For LoopC = 1 To MaxUsers
60            If UserList(LoopC).ConnID <> -1 And UserList(LoopC).ConnIDValida Then
70                Call CloseSocket(LoopC)
80            End If
90        Next LoopC
          
100       Call frmMain.Serv.Detener
110       Call frmMain.Serv.Iniciar(Puerto)
120   End If

#End If
End Sub

'Barrin 29/9/03
Private Sub Command21_Click()

10    If EnPausa = False Then
20        EnPausa = True
30        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
40        Command21.Caption = "Reanudar el servidor"
50    Else
60        EnPausa = False
70        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
80        Command21.Caption = "Pausar el servidor"
90    End If

End Sub

Private Sub Command22_Click()
10        Me.Visible = False
20        frmAdmin.Show
End Sub

Private Sub Command23_Click()
10    If MsgBox("¿Está seguro que desea hacer WorldSave, guardar pjs y cerrar?", vbYesNo, "Apagar Magicamente") = vbYes Then
20        Me.MousePointer = 11
          
30        FrmStat.Show
         
          'WorldSave
40        Call ES.DoBackUp

          'commit experiencia
50        Call mGroup.DistributeExpAndGldGroups

          'Guardar Pjs
60        Call GuardarUsuarios
          
          'Chauuu
70        Unload frmMain
80    End If
End Sub

Private Sub Command24_Click()
    If EventosActivos = False Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» Activados.", FontTypeNames.FONTTYPE_SERVER))
        EventosActivos = True
    Else
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Eventos» Desactivados.", FontTypeNames.FONTTYPE_SERVER))
        EventosActivos = False
    End If
    
End Sub

Private Sub Command25_Click()
10    Call MD5sCarga

End Sub

Private Sub Command26_Click()
#If UsarQueSocket = 1 Then
          'Cierra el socket de escucha
10        If SockListen >= 0 Then Call apiclosesocket(SockListen)
          
          'Inicia el socket de escucha
20        SockListen = ListenForConnect(Puerto, hWndMsg, "")
#End If
End Sub

Private Sub Command27_Click()
10    frmUserList.Show

End Sub

Private Sub Command28_Click()
10        Call LoadBalance
End Sub

Private Sub Command29_Click()
10    LoadInvasiones
20    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> Archivo #Invasiones.dat ha sido actualizado.", FontTypeNames.FONTTYPE_SERVER))
    
End Sub

Private Sub Command3_Click()
10    If MsgBox("¡¡Atencion!! Si reinicia el servidor puede provocar la pérdida de datos de los usarios. ¿Desea reiniciar el servidor de todas maneras?", vbYesNo) = vbYes Then
20        Me.Visible = False
30        Call General.Restart
40    End If
End Sub

Private Sub Command30_Click()
10     Call LoadCofres
20     Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> Archivo #Cofres.dat ha sido actualizado.", FontTypeNames.FONTTYPE_SERVER))

End Sub

Private Sub Command4_Click()
10    On Error GoTo eh
20        Me.MousePointer = 11
30        FrmStat.Show
40        Call ES.DoBackUp
50        Me.MousePointer = 0
60        MsgBox "WORLDSAVE OK!!"
70    Exit Sub
eh:
80    Call LogError("Error en WORLDSAVE")
End Sub

Private Sub Command5_Click()

      'Se asegura de que los sockets estan cerrados e ignora cualquier err
10    On Error Resume Next

20    If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

30    FrmStat.Show

40    If FileExist(App.Path & "\logs\errores.log", vbNormal) Then Kill App.Path & "\logs\errores.log"
50    If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\Connect.log"
60    If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
70    If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
80    If FileExist(App.Path & "\logs\Resurrecciones.log", vbNormal) Then Kill App.Path & "\logs\Resurrecciones.log"
90    If FileExist(App.Path & "\logs\Teleports.Log", vbNormal) Then Kill App.Path & "\logs\Teleports.Log"


#If UsarQueSocket = 1 Then
100   Call apiclosesocket(SockListen)
#ElseIf UsarQueSocket = 0 Then
110   frmMain.Socket1.Cleanup
120   frmMain.Socket2(0).Cleanup
#ElseIf UsarQueSocket = 2 Then
130   frmMain.Serv.Detener
#End If

      Dim LoopC As Integer

140   For LoopC = 1 To MaxUsers
150       Call CloseSocket(LoopC)
160   Next
        

170   LastUser = 0
180   NumUsers = 0

190   Call FreeNPCs
200   Call FreeCharIndexes

210   Call LoadSini
220   Call CargarBackUp
230   Call LoadOBJData

#If UsarQueSocket = 1 Then
240   SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 0 Then
250   frmMain.Socket1.AddressFamily = AF_INET
260   frmMain.Socket1.Protocol = IPPROTO_IP
270   frmMain.Socket1.SocketType = SOCK_STREAM
280   frmMain.Socket1.Binary = False
290   frmMain.Socket1.Blocking = False
300   frmMain.Socket1.BufferSize = 1024

310   frmMain.Socket2(0).AddressFamily = AF_INET
320   frmMain.Socket2(0).Protocol = IPPROTO_IP
330   frmMain.Socket2(0).SocketType = SOCK_STREAM
340   frmMain.Socket2(0).Blocking = False
350   frmMain.Socket2(0).BufferSize = 2048

      'Escucha
360   frmMain.Socket1.LocalPort = Puerto
370   frmMain.Socket1.listen
#End If

380   If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

End Sub

Private Sub Command6_Click()
10    Call ReSpawnOrigPosNpcs
End Sub

Private Sub Command7_Click()
10    FrmInterv.Show
End Sub

Private Sub Command8_Click()
10    Call CargarHechizos
End Sub

Private Sub Command9_Click()
10    Call CargarForbidenWords
End Sub

Private Sub Form_Deactivate()
10    frmServidor.Visible = False
End Sub

Private Sub Form_Load()
#If UsarQueSocket = 1 Then
10    Command20.Visible = True
20    Command26.Visible = True
#ElseIf UsarQueSocket = 0 Then
30    Command20.Visible = False
40    Command26.Visible = False
#ElseIf UsarQueSocket = 2 Then
50    Command20.Visible = True
60    Command26.Visible = False
#End If

70    VS1.min = 0
80    If picCont.Height > picFuera.ScaleHeight Then
90        VS1.max = picCont.Height - picFuera.ScaleHeight
100   Else
110       VS1.max = 0
120   End If
130   picCont.Top = -VS1.Value

End Sub

Private Sub VS1_Change()
10    picCont.Top = -VS1.Value
End Sub

Private Sub VS1_Scroll()
10    picCont.Top = -VS1.Value
End Sub


