VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Desterium AO"
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox PicRecuperar 
      BorderStyle     =   0  'None
      Height          =   3915
      Left            =   7200
      Picture         =   "frmConnect.frx":43B32
      ScaleHeight     =   3915
      ScaleWidth      =   6000
      TabIndex        =   10
      Top             =   3600
      Visible         =   0   'False
      Width           =   6000
      Begin VB.TextBox txtPin 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   285
         Left            =   2730
         TabIndex        =   13
         Top             =   2650
         Width           =   2550
      End
      Begin VB.TextBox txtMail 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   285
         Left            =   2730
         TabIndex        =   12
         Top             =   2100
         Width           =   2550
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   285
         Left            =   2730
         TabIndex        =   11
         Top             =   1450
         Width           =   2550
      End
      Begin VB.Image Image5 
         Height          =   405
         Left            =   3315
         Top             =   3315
         Width           =   2355
      End
      Begin VB.Image Image1 
         Height          =   405
         Left            =   390
         Top             =   3315
         Width           =   2355
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   11115
      Top             =   195
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox checkrecu 
      BackColor       =   &H8000000D&
      Caption         =   "Recordar Password"
      Height          =   195
      Left            =   4290
      TabIndex        =   6
      Top             =   6000
      Width           =   195
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   4485
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4800
      Width           =   3090
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4485
      TabIndex        =   0
      Top             =   4080
      Width           =   3090
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4890
      TabIndex        =   2
      Text            =   "7666"
      Top             =   -4800
      Width           =   825
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5760
      TabIndex        =   4
      Text            =   "localhost"
      Top             =   -4800
      Width           =   1575
   End
   Begin SHDocVwCtl.WebBrowser WebAuxiliar 
      Height          =   360
      Left            =   960
      TabIndex        =   5
      Top             =   -840
      Visible         =   0   'False
      Width           =   330
      ExtentX         =   582
      ExtentY         =   635
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label lblSaveData 
      BackStyle       =   0  'Transparent
      Caption         =   "Guardar datos de la cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   5940
      Width           =   3495
   End
   Begin VB.Image imgClose 
      Height          =   600
      Left            =   9945
      Top             =   8190
      Width           =   1770
   End
   Begin VB.Image imgConectar 
      Height          =   600
      Left            =   6240
      Top             =   5265
      Width           =   1380
   End
   Begin VB.Image Image4 
      Height          =   525
      Left            =   4290
      Top             =   5340
      Width           =   1770
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   6435
      Top             =   8190
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   4095
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label OnOff 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OFFLINE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   210
      Left            =   4485
      TabIndex        =   9
      Top             =   7800
      Width           =   3330
   End
   Begin VB.Image Recuperar 
      Height          =   600
      Left            =   975
      Top             =   8190
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "crear PEJOTAAAAAAAAAAAAAAAA"
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   -7560
      Width           =   4215
   End
   Begin VB.Label imgSalir 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   6480
      TabIndex        =   7
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Image imgServArgentina 
      Height          =   795
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   9240
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   -240
      Width           =   555
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Desterium AO 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software;you can redistribute it and/or modify
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
'
'Matías Fernando Pequeño
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Código Postal 1405

Option Explicit

Private cBotonCrearPj As clsGraphicalButton
Private cBotonRecuperarPass As clsGraphicalButton
Private cBotonManual As clsGraphicalButton
Private cBotonReglamento As clsGraphicalButton
Private cBotonCodigoFuente As clsGraphicalButton
Private cBotonBorrarPj As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton
Private cBotonLeerMas As clsGraphicalButton
Private cBotonForo As clsGraphicalButton
Private cBotonConectarse As clsGraphicalButton
Private cBotonTeclas As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub cmdNueva_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Activate()
      'On Error Resume Next


10    If ServersRecibidos Then
20        If CurServer <> 0 Then
30            IPTxt = ServersLst(1).Ip
40            PortTxt = ServersLst(1).Puerto
50        Else
60            IPTxt = IPdelServidor
70            PortTxt = PuertoDelServidor
80        End If
90    End If

100   If Winsock1.State <> sckClosed Then
110     Winsock1.Close
120   End If

130   Winsock1.Connect CurServerIp, CurServerPort
      'LOCOIP
      '127.0.0.1
TieneClan = False
    LoadDataAccount
    
    If DataAccount.State = 1 Then
        Me.checkrecu.value = 1
        Me.txtNombre.Text = DataAccount.Name
        Me.txtPasswd.Text = DataAccount.Passwd
    Else
        Me.checkrecu.value = 0
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
10        If KeyCode = 27 Then
20            prgRun = False
30        End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

      'Make Server IP and Port box visible
10    If KeyCode = vbKeyI And Shift = vbCtrlMask Then
          
          'Port
20        PortTxt.Visible = True
          'Label4.Visible = True
          
          'Server IP
30        PortTxt.Text = "7666"
40        IPTxt.Text = "localhost"
50        IPTxt.Visible = True
          'Label5.Visible = True
          
60        KeyCode = 0
70        Exit Sub
80    End If

End Sub


Private Sub Image1_Click()
    PicRecuperar.Visible = False
End Sub

Private Sub Image2_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    ShellExecute Me.hWnd, "open", "https://www.desteriumao.com/", "", "", 1
End Sub

Private Sub Image3_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    ShellExecute Me.hWnd, "open", "https://www.desteriumao.com/donaciones.html", "", "", 1
End Sub

Private Sub Image4_Click()
    Call Audio.PlayWave(SND_CLICK)
    
    FrmNuevaCuenta.Show
    
    Unload Me
End Sub

Private Sub Image5_Click()
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
          
    With Cuenta
        .Account = txtName.Text
        .PIN = txtPin.Text
        .Email = txtMail.Text
    End With
    
    'If CheckDataAccount = False Then Exit Sub

    EstadoLogin = e_RecoverAccount

    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
    
    Call Login

PicRecuperar.Visible = False
End Sub

Private Sub imgClose_Click()
    CloseClient
    
End Sub

Private Sub imgConectar_Click()
    
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    
    With Cuenta
        .Account = txtNombre.Text
        .Passwd = txtPasswd.Text
    End With
    
    'If CheckDataAccount = False Then Exit Sub

    EstadoLogin = e_ConnectAccount

    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
    
    Call Login
    
              
    If checkrecu.value = 1 Then
        SaveDataAccount Cuenta.Account, Cuenta.Passwd, 1
    
    Else
        SaveDataAccount 0, 0, 0
    
    End If
End Sub

Private Sub imgSalir_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    If MsgBox("¿Desea cerrar Desterium AO?", vbYesNo + vbQuestion, "Desterium AO") _
          = vbYes Then
30      prgRun = False
40    Else
50                Exit Sub
60            End If
End Sub

Private Sub imgServArgentina_Click()
10        Call Audio.PlayWave(SND_CLICK)
20        IPTxt.Text = IPdelServidor
30        PortTxt.Text = PuertoDelServidor
End Sub

Private Sub lblSaveData_Click()
    If checkrecu.value Then
        checkrecu.value = 0
    Else
        checkrecu.value = 1
    End If
End Sub

Private Sub RECUPERAR_Click()
    Call Audio.PlayWave(SND_CLICK)

    PicRecuperar.Visible = True
       
End Sub


Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
10        If KeyAscii = vbKeyReturn Then imgConectar_Click
End Sub

Private Sub WebAuxiliar_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, _
    flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As _
    Variant, Cancel As Boolean)
          
10        If InStr(1, URL, "alkon") <> 0 Then
20            Call ShellExecute(hWnd, "open", URL, vbNullString, vbNullString, _
                  SW_SHOWNORMAL)
30            Cancel = True
40        End If
          
End Sub

Private Sub webNoticias_NewWindow2(ppDisp As Object, Cancel As Boolean)
          
10        WebAuxiliar.RegisterAsBrowser = True
20        Set ppDisp = WebAuxiliar.Object
          
End Sub

Private Sub Winsock1_Connect()
10    OnOff.ForeColor = vbGreen
20    OnOff.Caption = "Online"
End Sub
 
Private Sub Winsock1_Error(ByVal number As Integer, Description As String, _
    ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal _
    HelpContext As Long, CancelDisplay As Boolean)
10    OnOff.ForeColor = vbRed
20    OnOff.Caption = "Offline"
End Sub
