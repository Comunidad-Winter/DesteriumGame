VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpciones 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOpciones.frx":0152
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox imgChkMusica 
      Caption         =   "Check1"
      Height          =   195
      Left            =   4155
      TabIndex        =   15
      Top             =   840
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox imgChkMostrarNews 
      Caption         =   "Check1"
      Height          =   195
      Left            =   4155
      TabIndex        =   14
      Top             =   840
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H80000007&
      Caption         =   "AB"
      Height          =   195
      Left            =   4155
      MaskColor       =   &H00E0E0E0&
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   1320
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox imgChkSonidos 
      Caption         =   "Check1"
      Height          =   195
      Left            =   4155
      TabIndex        =   11
      Top             =   1320
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox imgChkEfectosSonido 
      Caption         =   "Check1"
      Height          =   195
      Left            =   4155
      TabIndex        =   10
      Top             =   1800
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.ComboBox lstlenguajes 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmOpciones.frx":A5E7
      Left            =   5400
      List            =   "frmOpciones.frx":A5F1
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "optConsola"
      Height          =   195
      Left            =   4155
      TabIndex        =   7
      Top             =   1800
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   195
      Left            =   4155
      TabIndex        =   6
      Top             =   2280
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check3"
      Height          =   195
      Left            =   4155
      TabIndex        =   5
      Top             =   2760
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox imgChkPantalla 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   9360
      Width           =   3615
   End
   Begin VB.CheckBox imgChkNoMostrarNews 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   9000
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox imgChkConsola 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   8160
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox txtCantMensajes 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8400
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "5"
      Top             =   8280
      Width           =   255
   End
   Begin VB.TextBox txtLevel 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8400
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "40"
      Top             =   8670
      Width           =   255
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   1455
      Index           =   0
      Left            =   5520
      TabIndex        =   30
      Top             =   2160
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   2566
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   4
      Min             =   30
      Max             =   100
      SelStart        =   40
      TickStyle       =   2
      TickFrequency   =   4
      Value           =   40
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   1575
      Index           =   1
      Left            =   4680
      TabIndex        =   31
      Top             =   2040
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   2778
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   4
      Min             =   30
      Max             =   100
      SelStart        =   40
      TickStyle       =   2
      TickFrequency   =   4
      Value           =   40
      TextPosition    =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cambiar Clave Pin"
      Height          =   375
      Left            =   4080
      TabIndex        =   18
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calcular Promedio"
      Height          =   375
      Left            =   4080
      TabIndex        =   17
      Top             =   3240
      Width           =   2775
   End
   Begin VB.CommandButton imgConfigTeclas 
      Caption         =   "Configurar Teclas"
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton imgCambiarPasswd 
      Caption         =   "Cambiar Contraseña"
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton imgMsgPersonalizado 
      Caption         =   "Mensajes Personalizados"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   840
      Width           =   2775
   End
   Begin VB.Image imgSalir 
      Height          =   525
      Left            =   4320
      Top             =   4200
      Width           =   2490
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Musica"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5850
      TabIndex        =   29
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Efectos"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   28
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Efectos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   27
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Efectos de Sonido 3D"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   26
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Música"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   840
      Width           =   1110
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   600
      Top             =   3360
      Width           =   2490
   End
   Begin VB.Image Image3 
      Height          =   405
      Left            =   600
      Top             =   1200
      Width           =   2490
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Idioma:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   24
      Top             =   3675
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mostrar Noticias"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4425
      TabIndex        =   23
      Top             =   840
      Width           =   1710
   End
   Begin VB.Image Image4 
      Height          =   525
      Left            =   600
      Top             =   2280
      Width           =   2490
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Alphableding"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   22
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Consola Flotante"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      Top             =   1800
      Width           =   1710
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Limitar Fps"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Efecto de Combate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   19
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   4200
      Top             =   10080
      Width           =   210
   End
   Begin VB.Image imgChkDesactivarFragShooter 
      Height          =   225
      Left            =   5355
      Top             =   9300
      Width           =   210
   End
   Begin VB.Image imgChkAlMorir 
      Height          =   225
      Left            =   5355
      Top             =   9000
      Width           =   210
   End
   Begin VB.Image imgChkRequiredLvl 
      Height          =   225
      Left            =   5355
      Top             =   8640
      Width           =   210
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Tierras Nórdicas 0.11.6
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
'Tierras Nórdicas is based on Baronsoft's VB6 Online RPG
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

Private clsFormulario As clsFormMovementManager

Private cBotonConfigTeclas As clsGraphicalButton
Private cBotonMsgPersonalizado As clsGraphicalButton
Private cBotonMapa As clsGraphicalButton
Private cBotonCambiarPasswd As clsGraphicalButton
Private cBotonManual As clsGraphicalButton
Private cBotonRadio As clsGraphicalButton
Private cBotonSoporte As clsGraphicalButton
Private cBotonTutorial As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private picCheckBox As Picture

Private bMusicActivated As Boolean
Private bSoundActivated As Boolean
Private bSoundEffectsActivated As Boolean

Private loading As Boolean

Private Sub Check1_Click()
DialogosClanes.Activo = False
End Sub

Private Sub Check2_Click()
If ConAlfaB = 1 And Check2.value = vbUnchecked Then
ConAlfaB = 0
Else
ConAlfaB = 1
End If
End Sub

Private Sub Check3_Click()

If TSetup.bFPS = 0 And Check3.value = vbUnchecked Then
TSetup.bFPS = 1
Else
TSetup.bFPS = 0
End If
End Sub

Private Sub Check4_Click()
If TSetup.bGameCombat = 0 And Check4.value = vbUnchecked Then
TSetup.bGameCombat = 1
Else
TSetup.bGameCombat = 0
End If
End Sub

Private Sub Command1_Click()
  frmNewpin.Show vbModal, frmOpciones
'FrmSopORT.Show , frmMain
End Sub

Private Sub Command2_Click()
Call Audio.PlayWave(SND_CLICK)
    Call ShellExecute(0, "Open", "http://dsao.ucoz.es/Calculadora.html", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub Image2_Click()
'Video
imgChkMostrarNews.Visible = False
Label5.Visible = False
Check2.Visible = False
Label6.Visible = False
Check1.Visible = False
Label7.Visible = False
Check3.Visible = False
    Label8.Visible = False
    Check4.Visible = False
    Label9.Visible = False
    Command2.Visible = True
'General
imgMsgPersonalizado.Visible = True
imgCambiarPasswd.Visible = True
imgConfigTeclas.Visible = True
Label4.Visible = True
lstlenguajes.Visible = True
Command1.Visible = True
'Audio
imgChkMusica.Visible = False
Label3.Visible = False
imgChkSonidos.Visible = False
Label1.Visible = False
imgChkEfectosSonido.Visible = False
Label2.Visible = False
Label14.Visible = False
Label13.Visible = False
Slider1(0).Visible = False
Slider1(1).Visible = False
End Sub

Private Sub Image3_Click()
'Video
imgChkMostrarNews.Visible = False
Label5.Visible = False
Check2.Visible = False
Label6.Visible = False
Check1.Visible = False
Command2.Visible = False
Label7.Visible = False
Check3.Visible = False
    Label8.Visible = False
    Check4.Visible = False
    Label9.Visible = False
'General
imgMsgPersonalizado.Visible = False
imgCambiarPasswd.Visible = False
imgConfigTeclas.Visible = False
Label4.Visible = False
lstlenguajes.Visible = False
Command1.Visible = False
'Audio
imgChkMusica.Visible = True
Label3.Visible = True
imgChkSonidos.Visible = True
Label1.Visible = True
imgChkEfectosSonido.Visible = True
Label2.Visible = True
Label14.Visible = True
Label13.Visible = True
Slider1(0).Visible = True
Slider1(1).Visible = True
End Sub

Private Sub Image4_Click()
'Video
imgChkMostrarNews.Visible = True
Label5.Visible = True
Check2.Visible = True
Label6.Visible = True
Check1.Visible = True
Label7.Visible = True
Command2.Visible = False
Check3.Visible = True
Label8.Visible = True
Check4.Visible = True
    Label9.Visible = True
'General
imgMsgPersonalizado.Visible = False
imgCambiarPasswd.Visible = False
imgConfigTeclas.Visible = False
Label4.Visible = False
lstlenguajes.Visible = False
Command1.Visible = False

'Audio
imgChkMusica.Visible = False
Label3.Visible = False
imgChkSonidos.Visible = False
Label1.Visible = False
imgChkEfectosSonido.Visible = False
Label2.Visible = False
Label14.Visible = False
Label13.Visible = False
Slider1(0).Visible = False
Slider1(1).Visible = False
End Sub

Private Sub imgCambiarPasswd_Click()
    Call frmNewPassword.Show(vbModal, Me)
End Sub

Private Sub imgChkAlMorir_Click()
    ClientSetup.bDie = Not ClientSetup.bDie
    
    If ClientSetup.bDie Then
        imgChkAlMorir.Picture = picCheckBox
    Else
        Set imgChkAlMorir.Picture = Nothing
    End If
End Sub

Private Sub imgChkDesactivarFragShooter_Click()
    ClientSetup.bActive = Not ClientSetup.bActive
    
    If ClientSetup.bActive Then
        Set imgChkDesactivarFragShooter.Picture = Nothing
    Else
        imgChkDesactivarFragShooter.Picture = picCheckBox
    End If
End Sub

Private Sub imgChkRequiredLvl_Click()
    ClientSetup.bKill = Not ClientSetup.bKill
    
    If ClientSetup.bKill Then
        imgChkRequiredLvl.Picture = picCheckBox
    Else
        Set imgChkRequiredLvl.Picture = Nothing
    End If
End Sub

Private Sub txtCantMensajes_Change()
    txtCantMensajes.Text = Val(txtCantMensajes.Text)
    
    If txtCantMensajes.Text > 0 Then
        DialogosClanes.CantidadDialogos = txtCantMensajes.Text
    Else
        txtCantMensajes.Text = 5
    End If
End Sub

Private Sub txtLevel_Change()
    If Not IsNumeric(txtLevel) Then txtLevel = 0
    txtLevel = Trim$(txtLevel)
    ClientSetup.byMurderedLevel = CByte(txtLevel)
End Sub

Private Sub imgChkConsola_Click()
    DialogosClanes.Activo = False
    
    imgChkConsola.Picture = picCheckBox
    Set imgChkPantalla.Picture = Nothing
End Sub

Private Sub imgChkEfectosSonido_Click()

    If loading Then Exit Sub
    
    Call Audio.PlayWave(SND_CLICK)
        
    bSoundEffectsActivated = Not bSoundEffectsActivated
    
    Audio.SoundEffectsActivated = bSoundEffectsActivated
    
    If bSoundEffectsActivated Then
        imgChkEfectosSonido.Picture = picCheckBox
    Else
        Set imgChkEfectosSonido.Picture = Nothing
    End If
            
End Sub

Private Sub imgChkMostrarNews_Click()
    ClientSetup.bGuildNews = True
    
    imgChkMostrarNews.Picture = picCheckBox
    Set imgChkNoMostrarNews.Picture = Nothing
End Sub

Private Sub imgChkMusica_Click()

    If loading Then Exit Sub
    
    Call Audio.PlayWave(SND_CLICK)
    
    bMusicActivated = Not bMusicActivated
            
    If Not bMusicActivated Then
        Audio.MusicActivated = False
        Slider1(0).Enabled = False
        Set imgChkMusica.Picture = Nothing
    Else
        If Not Audio.MusicActivated Then  'Prevent the music from reloading
            Audio.MusicActivated = True
            Slider1(0).Enabled = True
            Slider1(0).value = Audio.MusicVolume
        End If
        
        imgChkMusica.Picture = picCheckBox
    End If

End Sub

Private Sub imgChkNoMostrarNews_Click()
    ClientSetup.bGuildNews = False
    
    imgChkNoMostrarNews.Picture = picCheckBox
    Set imgChkMostrarNews.Picture = Nothing
End Sub

Private Sub imgChkPantalla_Click()
    DialogosClanes.Activo = True
    
    imgChkPantalla.Picture = picCheckBox
    Set imgChkConsola.Picture = Nothing
End Sub

Private Sub imgChkSonidos_Click()

    If loading Then Exit Sub
    
    Call Audio.PlayWave(SND_CLICK)
    
    bSoundActivated = Not bSoundActivated
    
    If Not bSoundActivated Then
        Audio.SoundActivated = False
        RainBufferIndex = 0
        frmMain.IsPlaying = PlayLoop.plNone
        Slider1(1).Enabled = False
        
        Set imgChkSonidos.Picture = Nothing
    Else
        Audio.SoundActivated = True
        Slider1(1).Enabled = True
        Slider1(1).value = Audio.SoundVolume
        
        imgChkSonidos.Picture = picCheckBox
    End If
End Sub

Private Sub imgConfigTeclas_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Call frmCustomKeys.Show(vbModal, Me)
End Sub

Private Sub imgManual_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Call ShellExecute(0, "Open", "http://ao.alkon.com.ar/manual/", "", App.path, SW_SHOWNORMAL)
End Sub
Private Sub imgMsgPersonalizado_Click()
Call Audio.PlayWave(SND_CLICK)
    Call frmMessageTxt.Show(vbModeless, Me)
End Sub


Private Sub imgSalir_Click()
    Unload Me
    Call Audio.PlayWave(SND_CLICK)
    frmMain.SetFocus
End Sub
Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    'Me.Picture = LoadPicture(App.path & "\graficos\VentanaOpciones.jpg")
    LoadButtons
    
    loading = True      'Prevent sounds when setting check's values
    LoadUserConfig
    loading = False     'Enable sounds when setting check's values
    'General
    imgMsgPersonalizado.Visible = False
    imgMsgPersonalizado.Visible = False
    imgCambiarPasswd.Visible = False
    imgConfigTeclas.Visible = False
    Label4.Visible = False
    Command2.Visible = False
    
    lstlenguajes.Visible = False
    Command1.Visible = False
    'Video
    imgChkMostrarNews.Visible = False
    Label5.Visible = False
    Check2.Visible = False
    Label6.Visible = False
    Check2.Visible = False
    Label6.Visible = False
    Check1.Visible = False
   Label7.Visible = False
   Check3.Visible = False
    Label8.Visible = False
    Check4.Visible = False
    Label9.Visible = False
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    
    GrhPath = DirGraficos

    Set cBotonConfigTeclas = New clsGraphicalButton
    Set cBotonMsgPersonalizado = New clsGraphicalButton
    Set cBotonMapa = New clsGraphicalButton
    Set cBotonCambiarPasswd = New clsGraphicalButton
    Set cBotonManual = New clsGraphicalButton
    Set cBotonRadio = New clsGraphicalButton
    Set cBotonSoporte = New clsGraphicalButton
    Set cBotonTutorial = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton

End Sub

Private Sub LoadUserConfig()

    ' Load music config
    bMusicActivated = Audio.MusicActivated
    Slider1(0).Enabled = bMusicActivated
    
    If bMusicActivated Then
        imgChkMusica.Picture = picCheckBox
        
        Slider1(0).value = Audio.MusicVolume
    End If
    
    
    ' Load Sound config
    bSoundActivated = Audio.SoundActivated
    Slider1(1).Enabled = bSoundActivated
    
    If bSoundActivated Then
        imgChkSonidos.Picture = picCheckBox
        
        Slider1(1).value = Audio.SoundVolume
    End If
    
    
    ' Load Sound Effects config
    bSoundEffectsActivated = Audio.SoundEffectsActivated
    If bSoundEffectsActivated Then imgChkEfectosSonido.Picture = picCheckBox
    
    txtCantMensajes.Text = CStr(DialogosClanes.CantidadDialogos)
    
    If DialogosClanes.Activo Then
        imgChkPantalla.Picture = picCheckBox
    Else
        imgChkConsola.Picture = picCheckBox
    End If
    
    If ClientSetup.bGuildNews Then
        imgChkMostrarNews.Picture = picCheckBox
    Else
        imgChkNoMostrarNews.Picture = picCheckBox
    End If
        
    If ClientSetup.bKill Then imgChkRequiredLvl.Picture = picCheckBox
    If ClientSetup.bDie Then imgChkAlMorir.Picture = picCheckBox
    If Not ClientSetup.bActive Then imgChkDesactivarFragShooter.Picture = picCheckBox
    
    txtLevel = ClientSetup.byMurderedLevel
End Sub

Private Sub Slider1_Change(Index As Integer)
    Select Case Index
        Case 0
            Audio.MusicVolume = Slider1(0).value
        Case 1
            Audio.SoundVolume = Slider1(1).value
    End Select
End Sub

Private Sub Slider1_Scroll(Index As Integer)
    Select Case Index
        Case 0
            Audio.MusicVolume = Slider1(0).value
        Case 1
            Audio.SoundVolume = Slider1(1).value
    End Select
End Sub
