VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmOpciones 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   ClientHeight    =   4785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
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
   ScaleHeight     =   319
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkDerecho 
      Caption         =   "Check3"
      Height          =   195
      Left            =   4140
      TabIndex        =   32
      Top             =   3510
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox chkWalk 
      Caption         =   "Check3"
      Height          =   195
      Left            =   4170
      TabIndex        =   30
      Top             =   3120
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calcular Promedio"
      Height          =   375
      Left            =   4080
      TabIndex        =   29
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check3"
      Height          =   195
      Left            =   4155
      TabIndex        =   27
      Top             =   2760
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   195
      Left            =   4155
      TabIndex        =   25
      Top             =   2280
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Check1 
      Caption         =   "optConsola"
      Height          =   195
      Left            =   4155
      TabIndex        =   23
      Top             =   1800
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox imgChkPantalla 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   9360
      Width           =   3615
   End
   Begin VB.CheckBox imgChkNoMostrarNews 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   9000
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox imgChkConsola 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   8160
      Value           =   1  'Checked
      Width           =   255
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
      ItemData        =   "frmOpciones.frx":1A94D
      Left            =   5400
      List            =   "frmOpciones.frx":1A957
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton imgMsgPersonalizado 
      Caption         =   "Mensajes Personalizados"
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   840
      Width           =   2775
   End
   Begin VB.CheckBox imgChkEfectosSonido 
      Caption         =   "Check1"
      Height          =   195
      Left            =   4155
      TabIndex        =   8
      Top             =   1800
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox imgChkSonidos 
      Caption         =   "Check1"
      Height          =   195
      Left            =   4155
      TabIndex        =   6
      Top             =   1320
      Value           =   1  'Checked
      Width           =   195
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
      Height          =   1695
      Index           =   0
      Left            =   5760
      TabIndex        =   3
      Top             =   2040
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   2990
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   4
      Max             =   100
      SelStart        =   40
      TickStyle       =   2
      TickFrequency   =   4
      Value           =   40
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   1695
      Index           =   1
      Left            =   4320
      TabIndex        =   2
      Top             =   2040
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   2990
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   4
      Max             =   100
      SelStart        =   90
      TickStyle       =   2
      TickFrequency   =   4
      Value           =   90
      TextPosition    =   1
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H80000007&
      Caption         =   "AB"
      Height          =   195
      Left            =   4155
      MaskColor       =   &H00E0E0E0&
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   1320
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CommandButton imgConfigTeclas 
      Caption         =   "Configurar Teclas"
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CheckBox imgChkMostrarNews 
      Caption         =   "Check1"
      Height          =   195
      Left            =   4155
      TabIndex        =   19
      Top             =   840
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox imgChkMusica 
      Caption         =   "Check1"
      Height          =   195
      Left            =   4155
      TabIndex        =   10
      Top             =   840
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.Label lblDerecho 
      BackStyle       =   0  'Transparent
      Caption         =   "Ver perfiles"
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
      Left            =   4485
      TabIndex        =   33
      Top             =   3510
      Width           =   2100
   End
   Begin VB.Label lblWalk 
      BackStyle       =   0  'Transparent
      Caption         =   "No moverse al hablar"
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
      Left            =   4485
      TabIndex        =   31
      Top             =   3120
      Width           =   2100
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
      TabIndex        =   28
      Top             =   2760
      Width           =   2055
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
      TabIndex        =   26
      Top             =   2280
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
      TabIndex        =   24
      Top             =   1800
      Width           =   1710
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
   Begin VB.Image Image4 
      Height          =   525
      Left            =   600
      Top             =   2280
      Width           =   2490
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
      TabIndex        =   20
      Top             =   840
      Width           =   1710
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
      TabIndex        =   15
      Top             =   3675
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   405
      Left            =   600
      Top             =   1200
      Width           =   2490
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   600
      Top             =   3360
      Width           =   2490
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
      TabIndex        =   11
      Top             =   840
      Width           =   1110
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   4200
      Top             =   10080
      Width           =   210
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
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
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
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Efectos"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Musica"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5850
      TabIndex        =   4
      Top             =   3720
      Width           =   735
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
   Begin VB.Image imgSalir 
      Height          =   525
      Left            =   4320
      Top             =   4200
      Width           =   2490
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

Private Sub chkDerecho_Click()
10        Config.ClickDerecho = CBool(chkDerecho.value)
          
          If Config.ClickDerecho Then
            Call WriteVar(App.path & "\INIT\CONFIGDS.DAT", "CONFIG", "CLICDERECHO", "1")
          Else
            Call WriteVar(App.path & "\INIT\CONFIGDS.DAT", "CONFIG", "CLICDERECHO", "0")
          End If
20
End Sub

Private Sub chkWalk_Click()
          
10        Config.NotWalkToConsole = CBool(chkWalk.value)

         If Config.NotWalkToConsole Then
20          Call WriteVar(App.path & "\INIT\CONFIGDS.DAT", "CONFIG", "NOTWALK", "1")
         Else
            Call WriteVar(App.path & "\INIT\CONFIGDS.DAT", "CONFIG", "NOTWALK", "0")
         End If

End Sub

Private Sub Command2_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    ShellExecute Me.hWnd, "open", "https://www.desteriumao.com/calc_vida.html", "", "", 1
End Sub

Private Sub Check1_Click()
10    DialogosClanes.Activo = False
End Sub

Private Sub Check2_Click()
10    If ConAlfaB = 1 And Check2.value = vbUnchecked Then
20    ConAlfaB = 0
30    Else
40    ConAlfaB = 1
50    End If
End Sub

Private Sub Check3_Click()
10    If TSetup.bFPS = 0 And Check3.value = vbUnchecked Then
20    TSetup.bFPS = 1
30    Else
40    TSetup.bFPS = 0
50    End If
End Sub

Private Sub Check4_Click()
10    If TSetup.bGameCombat = 1 And Check4.value = vbUnchecked Then
20    TSetup.bGameCombat = 1
30    Else
40    TSetup.bGameCombat = 1
50    End If
End Sub

Private Sub Command1_Click()
10    Call Audio.PlayWave(SND_CLICK)
20        Call frmNewpin.Show(vbModal, Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
10        LastPressed.ToggleToNormal
End Sub

Private Sub Image2_Click()
10    Call Audio.PlayWave(SND_CLICK)
      'Video
20    imgChkMostrarNews.Visible = False
30    Label5.Visible = False
40    Check2.Visible = False
50    Label6.Visible = False
60    Check1.Visible = False
70    Label7.Visible = False
80    Check3.Visible = False
90        Label8.Visible = False
100       Check4.Visible = False
110       Label9.Visible = False
120           chkDerecho.Visible = False
130       lblDerecho.Visible = False
140       chkWalk.Visible = False
150       lblWalk.Visible = False
      'General
160   imgMsgPersonalizado.Visible = True
170   'imgCambiarPasswd.Visible = True
180   imgConfigTeclas.Visible = True
190   Label4.Visible = True
200   lstlenguajes.Visible = True
210   'Command1.Visible = True
220   Command2.Visible = True
      'Audio
230   imgChkMusica.Visible = False
240   Label3.Visible = False
250   imgChkSonidos.Visible = False
260   Label1.Visible = False
270   imgChkEfectosSonido.Visible = False
280   Label2.Visible = False
290   Label14.Visible = False
300   Label13.Visible = False
310   Slider1(0).Visible = False
320   Slider1(1).Visible = False
End Sub

Private Sub Image3_Click()
10    Call Audio.PlayWave(SND_CLICK)
      'Video
20    imgChkMostrarNews.Visible = False
30    Label5.Visible = False
40    Check2.Visible = False
50    Label6.Visible = False
60    Check1.Visible = False
70    Label7.Visible = False
80    Check3.Visible = False
90        Label8.Visible = False
100       Check4.Visible = False
110       Label9.Visible = False
120       chkWalk.Visible = False
130       lblWalk.Visible = False
      'General
140   imgMsgPersonalizado.Visible = False
150   'imgCambiarPasswd.Visible = False
160   imgConfigTeclas.Visible = False
170   Label4.Visible = False
180   lstlenguajes.Visible = False
190   'Command1.Visible = False
200   Command2.Visible = False
210   chkWalk.Visible = False
220   lblWalk.Visible = False
230       chkDerecho.Visible = False
240       lblDerecho.Visible = False
      'Audio

250   imgChkMusica.Visible = True
260   Label3.Visible = True
270   imgChkSonidos.Visible = True
280   Label1.Visible = True
290   imgChkEfectosSonido.Visible = True
300   Label2.Visible = True
310   Label14.Visible = True
320   Label13.Visible = True
330   Slider1(0).Visible = True
340   Slider1(1).Visible = True
End Sub

Private Sub Image4_Click()
10    Call Audio.PlayWave(SND_CLICK)
      'Video
20    imgChkMostrarNews.Visible = True
30    Label5.Visible = True
40    Check2.Visible = True
50    Label6.Visible = True
60    Check1.Visible = True
70    Label7.Visible = True
80    Check3.Visible = True
90    Label8.Visible = True
100   Check4.Visible = True
110       Label9.Visible = True
120       chkWalk.Visible = True
130       lblWalk.Visible = True
140           chkDerecho.Visible = True
150       lblDerecho.Visible = True
      'General
160   imgMsgPersonalizado.Visible = False
170   'imgCambiarPasswd.Visible = False
180   imgConfigTeclas.Visible = False
190   Label4.Visible = False
200   lstlenguajes.Visible = False
210   'Command1.Visible = False
220   Command2.Visible = False
      'Audio
230   imgChkMusica.Visible = False
240   Label3.Visible = False
250   imgChkSonidos.Visible = False
260   Label1.Visible = False
270   imgChkEfectosSonido.Visible = False
280   Label2.Visible = False
290   Label14.Visible = False
300   Label13.Visible = False
310   Slider1(0).Visible = False
320   Slider1(1).Visible = False
End Sub

Private Sub imgCambiarPasswd_Click()
10    Call Audio.PlayWave(SND_CLICK)
20        Call frmNewPassword.Show(vbModal, Me)
End Sub

Private Sub imgChkAlMorir_Click()
10        ClientSetup.bDie = Not ClientSetup.bDie
          
20        If ClientSetup.bDie Then
30            imgChkAlMorir.Picture = picCheckBox
40        Else
50            Set imgChkAlMorir.Picture = Nothing
60        End If
End Sub

Private Sub imgChkDesactivarFragShooter_Click()
10        ClientSetup.bActive = Not ClientSetup.bActive
          
20        If ClientSetup.bActive Then
30            Set imgChkDesactivarFragShooter.Picture = Nothing
40        Else
50            imgChkDesactivarFragShooter.Picture = picCheckBox
60        End If
End Sub

Private Sub imgChkRequiredLvl_Click()
10        ClientSetup.bKill = Not ClientSetup.bKill
          
20        If ClientSetup.bKill Then
30            imgChkRequiredLvl.Picture = picCheckBox
40        Else
50            Set imgChkRequiredLvl.Picture = Nothing
60        End If
End Sub

Private Sub txtCantMensajes_Change()
10        txtCantMensajes.Text = Val(txtCantMensajes.Text)
          
20        If txtCantMensajes.Text > 0 Then
30            DialogosClanes.CantidadDialogos = txtCantMensajes.Text
40        Else
50            txtCantMensajes.Text = 5
60        End If
End Sub

Private Sub txtLevel_Change()
10        If Not IsNumeric(txtLevel) Then txtLevel = 0
20        txtLevel = Trim$(txtLevel)
30        ClientSetup.byMurderedLevel = CByte(txtLevel)
End Sub

Private Sub imgChkConsola_Click()
10        DialogosClanes.Activo = False
          
20        imgChkConsola.Picture = picCheckBox
30        Set imgChkPantalla.Picture = Nothing
End Sub

Private Sub imgChkEfectosSonido_Click()

10        If loading Then Exit Sub
          
20        Call Audio.PlayWave(SND_CLICK)
              
30        bSoundEffectsActivated = Not bSoundEffectsActivated
          
40        Audio.SoundEffectsActivated = bSoundEffectsActivated
          
50        If bSoundEffectsActivated Then
60            imgChkEfectosSonido.Picture = picCheckBox
70        Else
80            Set imgChkEfectosSonido.Picture = Nothing
90        End If
                  
End Sub

Private Sub imgChkMostrarNews_Click()
10        ClientSetup.bGuildNews = True
          
20        imgChkMostrarNews.Picture = picCheckBox
30        Set imgChkNoMostrarNews.Picture = Nothing
End Sub

Private Sub imgChkMusica_Click()

10        If loading Then Exit Sub
          
20        Call Audio.PlayWave(SND_CLICK)
          
30        bMusicActivated = Not bMusicActivated
                  
40        If Not bMusicActivated Then
50            Audio.MusicActivated = False
60            Slider1(0).Enabled = False
70            Set imgChkMusica.Picture = Nothing
80        Else
90            If Not Audio.MusicActivated Then  'Prevent the music from reloading
100               Audio.MusicActivated = True
110               Slider1(0).Enabled = True
120               Slider1(0).value = Audio.MusicVolume
130           End If
              
140           imgChkMusica.Picture = picCheckBox
150       End If

End Sub

Private Sub imgChkNoMostrarNews_Click()
10        ClientSetup.bGuildNews = False
          
20        imgChkNoMostrarNews.Picture = picCheckBox
30        Set imgChkMostrarNews.Picture = Nothing
End Sub

Private Sub imgChkPantalla_Click()
10        DialogosClanes.Activo = True
          
20        imgChkPantalla.Picture = picCheckBox
30        Set imgChkConsola.Picture = Nothing
End Sub

Private Sub imgChkSonidos_Click()

10        If loading Then Exit Sub
          
20        Call Audio.PlayWave(SND_CLICK)
          
30        bSoundActivated = Not bSoundActivated
          
40        If Not bSoundActivated Then
50            Audio.SoundActivated = False
60            RainBufferIndex = 0
70            frmMain.IsPlaying = PlayLoop.plNone
80            Slider1(1).Enabled = False
              
90            Set imgChkSonidos.Picture = Nothing
100       Else
110           Audio.SoundActivated = True
120           Slider1(1).Enabled = True
130           Slider1(1).value = Audio.SoundVolume
              
140           imgChkSonidos.Picture = picCheckBox
150       End If
End Sub

Private Sub imgConfigTeclas_Click()
10        If Not loading Then Call Audio.PlayWave(SND_CLICK)
20        Call frmCustomKeys.Show(vbModal, Me)
End Sub

Private Sub imgManual_Click()
10        If Not loading Then Call Audio.PlayWave(SND_CLICK)
20        Call ShellExecute(0, "Open", "http://ao.alkon.com.ar/manual/", "", App.path, _
              SW_SHOWNORMAL)
End Sub
Private Sub imgMsgPersonalizado_Click()
10    Call Audio.PlayWave(SND_CLICK)
20        Call frmMessageTxt.Show(vbModeless, Me)
End Sub


Private Sub imgSalir_Click()
10    Call Audio.PlayWave(SND_CLICK)
20        Unload Me
30        frmMain.SetFocus
End Sub
Private Sub Form_Load()
          ' Handles Form movement (drag and drop).
10        Set clsFormulario = New clsFormMovementManager
20        clsFormulario.Initialize Me
          
          'Me.Picture = LoadPicture(App.path & "\Recursos\VentanaOpciones.jpg")
30        LoadButtons
          
40        loading = True      'Prevent sounds when setting check's values
50        LoadUserConfig
60        loading = False     'Enable sounds when setting check's values
          'General
70        imgMsgPersonalizado.Visible = False
80        imgMsgPersonalizado.Visible = False
90        'imgCambiarPasswd.Visible = False
100       imgConfigTeclas.Visible = False
110       Label4.Visible = False
120       lstlenguajes.Visible = False
130       'Command1.Visible = False
140       Command2.Visible = False

          'Video
150       imgChkMostrarNews.Visible = False
160       Label5.Visible = False
170       Check2.Visible = False
180       Label6.Visible = False
190       Check2.Visible = False
200       Label6.Visible = False
210       Check1.Visible = False
220      Label7.Visible = False
230      Check3.Visible = False
240       Label8.Visible = False
250       Check4.Visible = False
260       Label9.Visible = False
270       chkWalk.Visible = False
280       lblWalk.Visible = False
290       chkDerecho.Visible = False
300       lblDerecho.Visible = False
          
310       If Config.NotWalkToConsole Then
320           chkWalk.value = 1
330       Else
340           chkWalk.value = 0
350       End If
          
360       If Config.ClickDerecho Then
370           chkDerecho.value = 1
380       Else
390           chkDerecho.value = 0
400       End If
End Sub

Private Sub LoadButtons()
          Dim GrhPath As String
          
10        GrhPath = DirGraficos

20        Set cBotonConfigTeclas = New clsGraphicalButton
30        Set cBotonMsgPersonalizado = New clsGraphicalButton
40        Set cBotonMapa = New clsGraphicalButton
50        Set cBotonCambiarPasswd = New clsGraphicalButton
60        Set cBotonManual = New clsGraphicalButton
70        Set cBotonRadio = New clsGraphicalButton
80        Set cBotonSoporte = New clsGraphicalButton
90        Set cBotonTutorial = New clsGraphicalButton
100       Set cBotonSalir = New clsGraphicalButton
          
110       Set LastPressed = New clsGraphicalButton

End Sub

Private Sub LoadUserConfig()

          ' Load music config
10        bMusicActivated = Audio.MusicActivated
20        Slider1(0).Enabled = bMusicActivated
          
30        If bMusicActivated Then
40            imgChkMusica.Picture = picCheckBox
              
50            Slider1(0).value = Audio.MusicVolume
60        End If
          
          
          ' Load Sound config
70        bSoundActivated = Audio.SoundActivated
80        Slider1(1).Enabled = bSoundActivated
          
90        If bSoundActivated Then
100           imgChkSonidos.Picture = picCheckBox
              
110           Slider1(1).value = Audio.SoundVolume
120       End If
          
          
          ' Load Sound Effects config
130       bSoundEffectsActivated = Audio.SoundEffectsActivated
140       If bSoundEffectsActivated Then imgChkEfectosSonido.Picture = picCheckBox
          
150       txtCantMensajes.Text = CStr(DialogosClanes.CantidadDialogos)
          
160       If DialogosClanes.Activo Then
170           imgChkPantalla.Picture = picCheckBox
180       Else
190           imgChkConsola.Picture = picCheckBox
200       End If
          
210       If ClientSetup.bGuildNews Then
220           imgChkMostrarNews.Picture = picCheckBox
230       Else
240           imgChkNoMostrarNews.Picture = picCheckBox
250       End If
              
260       If ClientSetup.bKill Then imgChkRequiredLvl.Picture = picCheckBox
270       If ClientSetup.bDie Then imgChkAlMorir.Picture = picCheckBox
280       If Not ClientSetup.bActive Then imgChkDesactivarFragShooter.Picture = _
              picCheckBox
          
290       txtLevel = ClientSetup.byMurderedLevel
End Sub

Private Sub Slider1_Change(Index As Integer)
10        Select Case Index
              Case 0
20                Audio.MusicVolume = Slider1(0).value
30            Case 1
40                Audio.SoundVolume = Slider1(1).value + 90
50        End Select
End Sub

Private Sub Slider1_Scroll(Index As Integer)
10        Select Case Index
              Case 0
20                Audio.MusicVolume = Slider1(0).value
30            Case 1
40                Audio.SoundVolume = Slider1(1).value
50        End Select
End Sub
