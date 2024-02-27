VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Desterium AO"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearPersonaje.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3195
      Left            =   14235
      ScaleHeight     =   3195
      ScaleWidth      =   2535
      TabIndex        =   21
      Top             =   2535
      Width           =   2535
   End
   Begin VB.TextBox DESCRIPCIONCLASE 
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
      ForeColor       =   &H00C0E0FF&
      Height          =   1380
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   6885
      Width           =   6450
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":5BB4B
      Left            =   1755
      List            =   "frmCrearPersonaje.frx":5BB4D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4290
      Width           =   2265
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":5BB4F
      Left            =   1755
      List            =   "frmCrearPersonaje.frx":5BB59
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5850
      Width           =   2265
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":5BB6C
      Left            =   1755
      List            =   "frmCrearPersonaje.frx":5BB6E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   5070
      Width           =   2265
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1755
      MaxLength       =   30
      TabIndex        =   0
      Top             =   3510
      Width           =   2355
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   5
      Left            =   8805
      TabIndex        =   19
      Top             =   5775
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   4
      Left            =   8805
      TabIndex        =   18
      Top             =   5325
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   3
      Left            =   8805
      TabIndex        =   17
      Top             =   4830
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   2
      Left            =   8805
      TabIndex        =   16
      Top             =   4350
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   1
      Left            =   8805
      TabIndex        =   15
      Top             =   3900
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   5
      Left            =   9525
      TabIndex        =   14
      Top             =   5775
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   4
      Left            =   9525
      TabIndex        =   13
      Top             =   5325
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   3
      Left            =   9525
      TabIndex        =   12
      Top             =   4830
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   2
      Left            =   9525
      TabIndex        =   11
      Top             =   4350
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   1
      Left            =   9525
      TabIndex        =   10
      Top             =   3870
      Width           =   225
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   3525
      Left            =   4680
      TabIndex        =   9
      Top             =   3120
      Width           =   2250
   End
   Begin VB.Image imgVolver 
      Height          =   450
      Left            =   1560
      Top             =   7215
      Width           =   2460
   End
   Begin VB.Image imgCrear 
      Height          =   435
      Left            =   1560
      Top             =   6630
      Width           =   2460
   End
   Begin VB.Image imgGenero 
      Height          =   240
      Left            =   1170
      Top             =   5460
      Width           =   3240
   End
   Begin VB.Image imgClase 
      Height          =   360
      Left            =   1170
      Top             =   3900
      Width           =   3285
   End
   Begin VB.Image imgRaza 
      Height          =   255
      Left            =   1365
      Top             =   4680
      Width           =   3105
   End
   Begin VB.Image imgConstitucion 
      Height          =   255
      Left            =   7215
      Top             =   5850
      Width           =   1080
   End
   Begin VB.Image imgCarisma 
      Height          =   360
      Left            =   7410
      Top             =   5265
      Width           =   765
   End
   Begin VB.Image imgInteligencia 
      Height          =   345
      Left            =   7215
      Top             =   4725
      Width           =   1245
   End
   Begin VB.Image imgAgilidad 
      Height          =   360
      Left            =   7215
      Top             =   4290
      Width           =   1215
   End
   Begin VB.Image imgFuerza 
      Height          =   360
      Left            =   7410
      Top             =   3705
      Width           =   795
   End
   Begin VB.Image imgF 
      Height          =   270
      Left            =   10335
      Top             =   3510
      Width           =   270
   End
   Begin VB.Image imgM 
      Height          =   270
      Left            =   9555
      Top             =   3510
      Width           =   270
   End
   Begin VB.Image imgD 
      Height          =   270
      Left            =   8775
      Top             =   3510
      Width           =   270
   End
   Begin VB.Image imgNombre 
      Height          =   360
      Left            =   1170
      Top             =   3120
      Width           =   3330
   End
   Begin VB.Image imgTirarDados 
      Height          =   960
      Left            =   10140
      Top             =   2535
      Width           =   210
   End
   Begin VB.Image Image1 
      Height          =   3120
      Left            =   12285
      Stretch         =   -1  'True
      Top             =   1755
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Image imgDados 
      Height          =   975
      Left            =   10140
      MouseIcon       =   "frmCrearPersonaje.frx":5BB70
      MousePointer    =   99  'Custom
      Top             =   2535
      Width           =   1380
   End
   Begin VB.Image imgHogar 
      Height          =   2850
      Left            =   14235
      Picture         =   "frmCrearPersonaje.frx":5BCC2
      Top             =   6240
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   4
      Left            =   10305
      TabIndex        =   8
      Top             =   5310
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   3
      Left            =   10305
      TabIndex        =   7
      Top             =   4830
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   5
      Left            =   10305
      TabIndex        =   6
      Top             =   5790
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   2
      Left            =   10305
      TabIndex        =   5
      Top             =   4350
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   1
      Left            =   10305
      TabIndex        =   4
      Top             =   3885
      Width           =   225
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
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

Private cBotonPasswd As clsGraphicalButton
Private cBotonTirarDados As clsGraphicalButton
Private cBotonMail As clsGraphicalButton
Private cBotonNombre As clsGraphicalButton
Private cBotonConfirmPasswd As clsGraphicalButton
Private cBotonAtributos As clsGraphicalButton
Private cBotonD As clsGraphicalButton
Private cBotonM As clsGraphicalButton
Private cBotonF As clsGraphicalButton
Private cBotonFuerza As clsGraphicalButton
Private cBotonAgilidad As clsGraphicalButton
Private cBotonInteligencia As clsGraphicalButton
Private cBotonCarisma As clsGraphicalButton
Private cBotonConstitucion As clsGraphicalButton
Private cBotonEvasion As clsGraphicalButton
Private cBotonMagia As clsGraphicalButton
Private cBotonVida As clsGraphicalButton
Private cBotonEscudos As clsGraphicalButton
Private cBotonArmas As clsGraphicalButton
Private cBotonArcos As clsGraphicalButton
Private cBotonEspecialidad As clsGraphicalButton
Private cBotonPuebloOrigen As clsGraphicalButton
Private cBotonRaza As clsGraphicalButton
Private cBotonClase As clsGraphicalButton
Private cBotonGenero As clsGraphicalButton
Private cBotonAlineacion As clsGraphicalButton
Private cBotonVolver As clsGraphicalButton
Private cBotonCrear As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private picFullStar As Picture
Private picHalfStar As Picture
Private picGlowStar As Picture

Private Enum eHelp
    iePasswd
    ieTirarDados
    ieMail
    ieNombre
    ieConfirmPasswd
    ieAtributos
    ieD
    ieM
    ieF
    ieFuerza
    ieAgilidad
    ieInteligencia
    ieCarisma
    ieConstitucion
    ieEvasion
    ieMagia
    ieVida
    ieEscudos
    ieArmas
    ieArcos
    ieEspecialidad
    iePuebloOrigen
    ieRaza
    ieClase
    ieGenero
    ieAlineacion
End Enum

Private vHelp(25) As String
Private vEspecialidades() As String

Private Type tModRaza
    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Carisma As Single
    Constitucion As Single
End Type

Private Type tModClase
    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    DañoArmas As Double
    DañoProyectiles As Double
    Escudo As Double
    Magia As Double
    Vida As Double
    Hit As Double
End Type

Private ModRaza() As tModRaza
Private ModClase() As tModClase

Private NroRazas As Integer
Private NroClases As Integer

Private Cargando As Boolean

Private currentGrh As Long
Private Dir As E_Heading

Private Sub Form_Load()
          'Me.Picture = LoadPicture(DirGraficos & "VentanaCrearPersonaje.jpg")
          
10        Cargando = True
20        Call LoadCharInfo
30        Call CargarEspecialidades
          
40        Call IniciarGraficos
50        Call CargarCombos
          
60        Call LoadHelp
          
          
70        Call TirarDados
          
80        Cargando = False
          
          'UserClase = 0
90        UserSexo = 0
100       UserRaza = 0
110       UserHogar = 0
120       UserEmail = ""
130       UserHead = 0

End Sub

Private Sub CargarEspecialidades()

10        ReDim vEspecialidades(1 To NroClases)
          
20        vEspecialidades(eClass.Hunter) = "Ocultarse"
30        vEspecialidades(eClass.Thief) = "Robar y Ocultarse"
40        vEspecialidades(eClass.Assasin) = "Apuñalar"
          'vEspecialidades(eClass.Bandit) = "Combate Sin Armas"
50        vEspecialidades(eClass.Druid) = "Domar"
60        vEspecialidades(eClass.Pirat) = "Navegar"
End Sub

Private Sub IniciarGraficos()

          Dim GrhPath As String
10        GrhPath = DirGraficos
          
20        Set cBotonPasswd = New clsGraphicalButton
30        Set cBotonTirarDados = New clsGraphicalButton
40        Set cBotonMail = New clsGraphicalButton
50        Set cBotonNombre = New clsGraphicalButton
60        Set cBotonConfirmPasswd = New clsGraphicalButton
70        Set cBotonAtributos = New clsGraphicalButton
80        Set cBotonD = New clsGraphicalButton
90        Set cBotonM = New clsGraphicalButton
100       Set cBotonF = New clsGraphicalButton
110       Set cBotonFuerza = New clsGraphicalButton
120       Set cBotonAgilidad = New clsGraphicalButton
130       Set cBotonInteligencia = New clsGraphicalButton
140       Set cBotonCarisma = New clsGraphicalButton
150       Set cBotonConstitucion = New clsGraphicalButton
160       Set cBotonEvasion = New clsGraphicalButton
170       Set cBotonMagia = New clsGraphicalButton
180       Set cBotonVida = New clsGraphicalButton
190       Set cBotonEscudos = New clsGraphicalButton
200       Set cBotonArmas = New clsGraphicalButton
210       Set cBotonArcos = New clsGraphicalButton
220       Set cBotonEspecialidad = New clsGraphicalButton
230       Set cBotonPuebloOrigen = New clsGraphicalButton
240       Set cBotonRaza = New clsGraphicalButton
250       Set cBotonClase = New clsGraphicalButton
260       Set cBotonGenero = New clsGraphicalButton
270       Set cBotonAlineacion = New clsGraphicalButton
280       Set cBotonVolver = New clsGraphicalButton
290       Set cBotonCrear = New clsGraphicalButton
          
300       Set LastPressed = New clsGraphicalButton
          
          
          'Call cBotonPasswd.Initialize(imgPasswd, "", GrhPath & "BotonContraseña.jpg", GrhPath & "BotonContraseña.jpg", Me, , , False, False)
                                          
          'Call cBotonTirarDados.Initialize(imgTirarDados, "", GrhPath & "BotonTirarDados.jpg", GrhPath & "BotonTirarDados.jpg", Me, , , False, False)
                                          
          'Call cBotonMail.Initialize(imgMail, "", GrhPath & "BotonMailPj.jpg", GrhPath & "BotonMailPj.jpg", Me, , , False, False)
                                          
          'Call cBotonNombre.Initialize(imgNombre, "", GrhPath & "BotonNombrePJ.jpg", GrhPath & "BotonNombrePJ.jpg", Me, , , False, False)
                                          
          'Call cBotonConfirmPasswd.Initialize(imgConfirmPasswd, "", GrhPath & "BotonRepetirContraseña.jpg", GrhPath & "BotonRepetirContraseña.jpg", Me, , , False, False)
                                          
          'Call cBotonAtributos.Initialize(imgAtributos, "", GrhPath & "BotonAtributos.jpg", GrhPath & "BotonAtributos.jpg", Me, , , False, False)
                                          
          'Call cBotonD.Initialize(imgD, "", GrhPath & "BotonD.jpg", GrhPath & "BotonD.jpg", Me, , , False, False)
                                          
          'Call cBotonM.Initialize(imgM, "", GrhPath & "BotonM.jpg", GrhPath & "BotonM.jpg", Me, , , False, False)
                                          
          'Call cBotonF.Initialize(imgF, "", GrhPath & "BotonF.jpg", GrhPath & "BotonF.jpg", Me, , , False, False)
                                          
          'Call cBotonFuerza.Initialize(imgFuerza, "", GrhPath & "BotonFuerza.jpg", GrhPath & "BotonFuerza.jpg", Me, , , False, False)
                                          
         ' Call cBotonAgilidad.Initialize(imgAgilidad, "", GrhPath & "BotonAgilidad.jpg", GrhPath & "BotonAgilidad.jpg", Me, , , False, False)
                                          
          'Call cBotonInteligencia.Initialize(imgInteligencia, "", GrhPath & "BotonInteligencia.jpg", GrhPath & "BotonInteligencia.jpg", Me, , , False, False)
                                          
          'Call cBotonCarisma.Initialize(imgCarisma, "", GrhPath & "BotonCarisma.jpg", GrhPath & "BotonCarisma.jpg", Me, , , False, False)
                                          
          'Call cBotonConstitucion.Initialize(imgConstitucion, "", GrhPath & "BotonConstitucion.jpg", GrhPath & "BotonConstitucion.jpg", Me, , , False, False)
                                          
         ' Call cBotonEvasion.Initialize(imgEvasion, "", GrhPath & "BotonEvasion.jpg", GrhPath & "BotonEvasion.jpg", Me, , , False, False)
                                          
         ' Call cBotonMagia.Initialize(imgMagia, "", GrhPath & "BotonMagia.jpg", GrhPath & "BotonMagia.jpg", Me, , , False, False)
                                          
          'Call cBotonVida.Initialize(imgVida, "", GrhPath & "BotonVida.jpg", GrhPath & "BotonVida.jpg", Me, , , False, False)
                                          
          'Call cBotonEscudos.Initialize(imgEscudos, "", GrhPath & "BotonEscudos.jpg", GrhPath & "BotonEscudos.jpg", Me, , , False, False)
                                          
         ' Call cBotonArmas.Initialize(imgArmas, "", GrhPath & "BotonArmas.jpg", GrhPath & "BotonArmas.jpg", Me, , , False, False)
                                          
         ' Call cBotonArcos.Initialize(imgArcos, "", GrhPath & "BotonArcos.jpg", GrhPath & "BotonArcos.jpg", Me, , , False, False)
                                          
         ' Call cBotonEspecialidad.Initialize(imgEspecialidad, "", GrhPath & "BotonEspecialidad.jpg", GrhPath & "BotonEspecialidad.jpg", Me, , , False, False)
                                          
          'Call cBotonPuebloOrigen.Initialize(imgPuebloOrigen, "", GrhPath & "BotonPuebloOrigen.jpg", GrhPath & "BotonPuebloOrigen.jpg", Me, , , False, False)
                                          
          'Call cBotonRaza.Initialize(imgRaza, "", GrhPath & "BotonRaza.jpg", GrhPath & "BotonRaza.jpg", Me, , , False, False)
                                          
          'Call cBotonClase.Initialize(imgClase, "", GrhPath & "BotonClase.jpg", GrhPath & "BotonClase.jpg", Me, , , False, False)
                                          
          'Call cBotonGenero.Initialize(imgGenero, "", GrhPath & "BotonGenero.jpg", GrhPath & "BotonGenero.jpg", Me, , , False, False)
                                          
         ' Call cBotonAlineacion.Initialize(imgalineacion, "", GrhPath & "BotonAlineacion.jpg", GrhPath & "BotonAlineacion.jpg", Me, , , False, False)
                                          
         ' Call cBotonVolver.Initialize(imgVolver, "", GrhPath & "BotonVolverRollover.jpg", GrhPath & "BotonVolverClick.jpg", Me)
                                          
          'Call cBotonCrear.Initialize(imgCrear, "", GrhPath & "BotonCrearPersonajeRollover.jpg", GrhPath & "BotonCrearPersonajeClick.jpg", Me)

           
         ' Set picFullStar = LoadPicture(GrhPath & "EstrellaSimple.jpg")
          'Set picHalfStar = LoadPicture(GrhPath & "EstrellaMitad.jpg")
          'Set picGlowStar = LoadPicture(GrhPath & "EstrellaBrillante.jpg")

End Sub

Private Sub CargarCombos()
          Dim i As Integer
          
10        lstProfesion.Clear
          
20        For i = LBound(ListaClases) To NroClases
30            lstProfesion.AddItem ListaClases(i)
40        Next i
          
50        'lstHogar.Clear
          
60       ' For i = LBound(Ciudades()) To UBound(Ciudades())
70        '    lstHogar.AddItem Ciudades(i)
80        'Next i
          
90        lstRaza.Clear
          
100       For i = LBound(ListaRazas()) To NroRazas
110           lstRaza.AddItem ListaRazas(i)
120       Next i
          
130       lstProfesion.ListIndex = 1
End Sub

Function CheckData() As Boolean
170       If UserRaza = 0 Then
180           MsgBox "Seleccione la raza del personaje."
190           Exit Function
200       End If
          
210       If UserSexo = 0 Then
220           MsgBox "Seleccione el sexo del personaje."
230           Exit Function
240       End If
          
250       If UserClase = 0 Then
260           MsgBox "Seleccione la clase del personaje."
270           Exit Function
280       End If

          Dim i As Integer
330       For i = 1 To NUMATRIBUTOS
340           If UserAtributos(i) = 0 Then
350               MsgBox "Los atributos del personaje son invalidos."
360               Exit Function
370           End If
380       Next i
          
390       If Len(UserName) > 30 Then
400           MsgBox ("El nombre debe tener menos de 30 letras.")
410           Exit Function
420       End If
          
430       CheckData = True
End Function

Private Sub TirarDados()
10        Call WriteThrowDices
20        Call FlushBuffer
End Sub

Private Sub DirPJ_Click(Index As Integer)
10        Select Case Index
              Case 0
20                Dir = CheckDir(Dir + 1)
30            Case 1
40                Dir = CheckDir(Dir - 1)
50        End Select
          
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
10        ClearLabel
End Sub


Private Sub imgCrear_Click()

          Dim i As Integer
          Dim CharAscii As Byte
          
10        UserName = txtNombre.Text
                  
20    If Len(txtNombre.Text) < 3 Then
30        MsgBox "¡¡El nombre debe de tener mas de 3 carácteres!!"
40        Exit Sub
50    End If
       
60    If Len(txtNombre.Text) >= 12 Then
70        MsgBox "¡¡El nombre debe de tener menos de 12 caracteres!!"
80        Exit Sub
90    End If
                  
                  
      Dim AllCr As Long
      Dim CantidadEsp As Byte
      Dim thiscr As String
       
180   Do
190       AllCr = AllCr + 1
200       If AllCr > Len(UserName) Then Exit Do
210       thiscr = mid(UserName, AllCr, 1)
220       If InStr(1, " ", UCase(thiscr)) = 1 Then
230              CantidadEsp = CantidadEsp + 1
240       End If
250   Loop
260   If CantidadEsp > 1 Then
270        MsgBox "Nick inválido. El nombre no puede tener mas de un espacio."
280        Exit Sub
290   End If
300
          UserName = Trim$(UserName)
310       UserRaza = lstRaza.ListIndex + 1
320       UserSexo = lstGenero.ListIndex + 1
330       UserClase = lstProfesion.ListIndex + 1
          
340       For i = 1 To NUMATRIBUTES
350           UserAtributos(i) = Val(lblAtributos(i).Caption)
360       Next i
          
370       'UserHogar = lstHogar.ListIndex + 1
          
380       If Not CheckData Then Exit Sub

560       EstadoLogin = E_MODO.e_CreateCharAccount
620       Call Login
          
640       bShowTutorial = True
End Sub

Private Sub imgDados_Click()
10        Call Audio.PlayWave(SND_DICE)
20                Call TirarDados
End Sub

Private Sub imgEspecialidad_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieEspecialidad)
End Sub

Private Sub imgNombre_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieNombre)
End Sub

Private Sub imgPasswd_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.iePasswd)
End Sub

Private Sub imgConfirmPasswd_MouseMove(Button As Integer, Shift As Integer, X _
    As Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieConfirmPasswd)
End Sub

Private Sub imgAtributos_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieAtributos)
End Sub

Private Sub imgD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
10        lblHelp.Caption = vHelp(eHelp.ieD)
End Sub

Private Sub imgM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
10        lblHelp.Caption = vHelp(eHelp.ieM)
End Sub

Private Sub imgF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
10        lblHelp.Caption = vHelp(eHelp.ieF)
End Sub

Private Sub imgFuerza_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieFuerza)
End Sub

Private Sub imgAgilidad_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieAgilidad)
End Sub

Private Sub imgInteligencia_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieInteligencia)
End Sub

Private Sub imgCarisma_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieCarisma)
End Sub

Private Sub imgConstitucion_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieConstitucion)
End Sub

Private Sub imgArcos_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieArcos)
End Sub

Private Sub imgArmas_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieArmas)
End Sub

Private Sub imgEscudos_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieEscudos)
End Sub

Private Sub imgEvasion_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieEvasion)
End Sub

Private Sub imgMagia_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieMagia)
End Sub

Private Sub imgMail_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieMail)
End Sub

Private Sub imgVida_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieVida)
End Sub

Private Sub imgTirarDados_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieTirarDados)
End Sub

Private Sub imgPuebloOrigen_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.iePuebloOrigen)
End Sub

Private Sub imgRaza_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieRaza)
End Sub

Private Sub imgClase_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieClase)
End Sub

Private Sub imgGenero_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieGenero)
End Sub

Private Sub imgalineacion_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieAlineacion)
End Sub

Private Sub imgVolver_Click()
        
        
        FrmCuenta.Visible = True
          
10        bShowTutorial = False
          
20        Unload Me
End Sub

Private Sub lstGenero_Click()
10        UserSexo = lstGenero.ListIndex + 1
End Sub

Private Sub lstProfesion_Click()
10    On Error Resume Next
      '    Image1.Picture = LoadPicture(App.path & "\Recursos\" & lstProfesion.Text & ".jpg")
      '
20        UserClase = lstProfesion.ListIndex + 1
          
30        Select Case lstProfesion.Text
       
      Case Is = "Mago"
40           DESCRIPCIONCLASE.Text = _
                 "Mago : Es la clase por excelencia preferida por muchos de los jugadores. Si te interesan mucho los hechizos, encantamientos e invocaciones, esta clase es la indicada, ya que posee mucha maná. Pero la desventaja de esta clase es el hecho de tener poca vida como también poca evasión, lo que significa que las clases de cuerpo a cuerpo le van a acertar mucho más facilmente los golpes. Ventajas: Mucha maná y la mejor resistencia mágica. Desventajas: Poca evasión y pocos puntos de vida."
50            Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\Mago.jpg")
60    Case Is = "Clerigo"
70           DESCRIPCIONCLASE.Text = _
                 "Clérigo : Sabios y fuertes, es una clase que se especializa tanto en las artes mágicas como la lucha de cuerpo a cuerpo. Aunque sus golpes no sean tan fuertes puede combinar sus dos especialidades. Así como el druida, el clérigo tiene mucho maná, pero no tanto como el Mago. Ventajas: Buena defensa física y Mágica. Desventajas: Tiene un daño equitativo."
80            Picture1.Picture = LoadPicture(App.path & _
                  "\Recursos\Clases\Clerigo.jpg")
90    Case Is = "Guerrero"
100          DESCRIPCIONCLASE.Text = _
                 "Guerrero : Una clase que no usa magias sino que directamente usa la fuerza y combate cuerpo a cuerpo, sus golpes resultan ser increíblemente devastadores cuando éste se encuentra en su punto más elevado, posee una gran vida pero su desventaja es que no puede usar magias, esto hace que sea una presa fácil si no está acompañado. Aunque siempre puede contar con su velocidad con el arco y la flecha. Ventajas: Muchos puntos de vida, mucha fuerza, mucha defensa física. Desventajas: No tiene mana."
110           Picture1.Picture = LoadPicture(App.path & _
                  "\Recursos\Clases\Guerrero.jpg")
120   Case Is = "Asesino"
130          DESCRIPCIONCLASE.Text = _
                 "Asesino : Simplemente un ser sanguinario, la característica especial de esta clase es apuñalar mortalmente, esto significa que de un golpe podrías dejar a tu enemigo prácticamente muerto, su evasión es la mejor y no conoce el miedo contra quienes intenten pegarle. Ventajas: Mucha evasión y golpe mortal. Desventajas: Poco maná."
140           Picture1.Picture = LoadPicture(App.path & _
                  "\Recursos\Clases\Asesino.jpg")
150   Case Is = "Ladron"
160          DESCRIPCIONCLASE.Text = _
                 "Ladrón: Sigiloso delincuente. Quizás no sean poderosamente buenos con la espada, quizás no usen magias, no domen bestias, pero si te cruzas con ellos corre o ataca rápido dado que su gran habilidad es el hurto, con esta gran habilidad puede dejar a grandes guerreros totalmente indefensos, despojándolos de todas sus pertenecías mas preciadas. Son deshonrosos y delincuentes pero esto es solo una porción de como verdaderamente describirlos. Ventajas: Pueden ocultarse durante demasiado tiempo y robarte pertenencias u oro. Desventajas: Pocos puntos de vida, poca resistencia mágica y no tiene maná."
170           Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\Ladron.jpg")
      'Case Is = "Bandido"
             'DESCRIPCIONCLASE.Text = "Bandido: Sigiloso guerrero que puede ser letal si no lo ves a tiempo. Tiene casi tantos puntos de vida como un paladín pero su maná es mas limitada. El mismo cuenta con una habilidad para ocultarse que no la tienen ninguna otra clase y si te descuidas puedes recibir un golpe crítico y mortal de su parte. Ventajas: Puede ocultarse durante mucho tiempo y dar golpes críticos. Desventajas: Poca resistencia mágica y poco maná."
             ' Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\bandido.jpg")
180   Case Is = "Bardo"
190          DESCRIPCIONCLASE.Text = _
                 "Bardo : Una clase muy práctica frente a las clases de cuerpo a cuerpo ya que posee una gran evasión lo que resulta difícil para enemigo cuando intente acertarle un golpe.  Sus ataques mágicos son muy poderosos aún más cuando usa un ítem especial que le da bonificación de poder mágico. Ventajas: mucho poder mágico y mucha evasión contra golpes. Desventajas: Poca resistencia mágica."
200           Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\Bardo.jpg")
210   Case Is = "Druida"
220          DESCRIPCIONCLASE.Text = _
                 "Druida : Amantes de la naturaleza, una clase que se especializa en domar criaturas y usarlas como mascotas, también usa conjuros para invocar otros tipos de criaturas que acuden a su ayuda, lo que permite que en su entrenamiento nunca esté solo. Tiene un hechizo especial que consiste en tomar la imagen de alguien o algo y transformarse, esto sirve para confundir a los enemigos, ideal para despistar en una batalla. Aunque no tiene tanta maná como el Mago, está entre las clases que más poseen. Ventajas: Mucho poder mágico y muy buena defensa mágica. Desventajas: Poca evasión."
230           Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\Druida.jpg")
240   Case Is = "Paladin"
250          DESCRIPCIONCLASE.Text = _
                 "Paladín : El paladín es una clase con mucha vida y una gran fuerza, ideal para encuentro contra otra clase que lucha cuerpo a cuerpo ya que sus golpes son algo más débiles que los del guerrero. Una de las desventajas es que su maná que es muy limitada y eso dificulta sus peleas contra clases mágicas. Ventajas: Buena Fuerza, buena defensa cuerpo a cuerpo y muchos puntos de vida. Desventajas: Poca resistencia mágica, poco mana."
260           Picture1.Picture = LoadPicture(App.path & _
                  "\Recursos\Clases\Paladin.jpg")
270   Case Is = "Cazador"
280          DESCRIPCIONCLASE.Text = _
                 "Cazador : Clase que no usa magia, pero que es muy hábil usando armas a distancias, tiene la habilidad de poder ocultarse entre las sombras cuando éste usa su armadura de cazador, tiene una bonificación de daño crítico para mejorar el rendimiento de su entrenamiento y esto hace que resulte fácil de entrenar. Sus fuertes ataques a distancia, hacen que cualquiera evite acercarse a él, aunque si está solo, el no tener magia le trae muchos problemas. Ventajas: Ataque a distancia de gran poder, habilidad de ocultarse. Desventajas: No tiene mana."
290           Picture1.Picture = LoadPicture(App.path & _
                  "\Recursos\Clases\Cazador.jpg")
300   Case Is = "Trabajador"
310          DESCRIPCIONCLASE.Text = _
                 "Trabajador : Tienen gran conocimiento en la construcción o extracción de bienes materiales para la subsistencia de las ciudades. Poseen gran cantidad de puntos de vida ya que los mismos se arriesgan a tareas en lugares peligrosos. Ventajas: Pueden construir o extraer materiales para la construcción o para la subsistencia. Desventajas: No tienen maná ni dominio en artes mágicas o de combate."
320           Picture1.Picture = LoadPicture(App.path & _
                  "\Recursos\Clases\trabajador.jpg")
330   Case Is = "Pirata"
340          DESCRIPCIONCLASE.Text = _
                 "Pirata : Explorador sin miedo. Se dice que esta fabulosa clase dedico su vida a al conocimiento del mar por ende antiguos escritos dicen que mas de la mitad de las tierras de Desterium fueron descubiertas por ellos aunque son muy aptos para la navegación también son comerciantes despreciable y grandes estafadores si te cruzas con ellos mas vale corre pues nunca andan solos. Ventajas: Amplio espacio en el inventario, al igual que vida y habilidades de navegación. Desventajas: No tiene mana."
350           Picture1.Picture = LoadPicture(App.path & "\Recursos\Clases\Pirata.jpg")
360   End Select
          
370       Call UpdateStats
End Sub

Private Sub lstRaza_Click()
10        UserRaza = lstRaza.ListIndex + 1
          
20        Call UpdateStats
End Sub

Private Sub picHead_Click(Index As Integer)
          ' No se mueve si clickea al medio
10        If Index = 2 Then Exit Sub
          
          Dim Counter As Integer
          Dim Head As Integer
          
20        Head = UserHead
          
          
30        UserHead = Head
          
          
End Sub

Private Sub PIN_CLICK()
10    MsgBox _
          "Recuerda colocar datos que sólo tu sepas para evitar robos o pérdidas. También para poder acceder a funciones como borrar o recuperar el personaje, intercambiar tu personaje dentro del juego."
End Sub


Private Sub txtConfirmPasswd_MouseMove(Button As Integer, Shift As Integer, X _
    As Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieConfirmPasswd)
End Sub

Private Sub txtMail_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieMail)
End Sub


Private Sub txtNombre_Change()
    
10        txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
10        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Function CheckDir(ByRef Dir As E_Heading) As E_Heading

10        If Dir > E_Heading.WEST Then Dir = E_Heading.NORTH
20        If Dir < E_Heading.NORTH Then Dir = E_Heading.WEST
          
30        CheckDir = Dir
          
40      '  currentGrh = BodyData(UserBody).Walk(Dir).GrhIndex
50      '  If currentGrh > 0 Then tAnimacion.Interval = _
              Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)

End Function

Private Sub LoadHelp()
10        vHelp(eHelp.iePasswd) = _
              "La contraseña que utilizarás para conectar tu personaje al juego."
20        vHelp(eHelp.ieTirarDados) = _
              "Presionando sobre los dados, se modificarán al azar los atributos de tu personaje, de esta manera puedes elegir los que más te parezcan para definir a tu personaje."
30        vHelp(eHelp.ieMail) = _
              "Es sumamente importante que ingreses una dirección de correo electrónico válida, ya que en el caso de perder la contraseña de tu personaje, se te enviará cuando lo requieras, a esa dirección."
40        vHelp(eHelp.ieNombre) = _
              "Sé cuidadoso al seleccionar el nombre de tu personaje. Argentum es un juego de rol, un mundo mágico y fantástico, y si seleccionás un nombre obsceno o con connotación política, los administradores borrarán tu personaje y no habrá ninguna posibilidad de recuperarlo."
50        vHelp(eHelp.ieConfirmPasswd) = _
              "La contraseña que utilizarás para conectar tu personaje al juego."
60        vHelp(eHelp.ieAtributos) = _
              "Son las cualidades que definen tu personaje. Generalmente se los llama ""Dados"". (Ver Tirar Dados)"
70        vHelp(eHelp.ieD) = _
              "Son los atributos que obtuviste al azar. Presioná la esfera roja para volver a tirarlos."
80        vHelp(eHelp.ieM) = _
              "Son los modificadores por raza que influyen en los atributos de tu personaje."
90        vHelp(eHelp.ieF) = _
              "Los atributos finales de tu personaje, de acuerdo a la raza que elegiste."
100       vHelp(eHelp.ieFuerza) = _
              "De ella dependerá qué tan potentes serán tus golpes, tanto con armas de cuerpo a cuerpo, a distancia o sin armas."
110       vHelp(eHelp.ieAgilidad) = _
              "Este atributo intervendrá en qué tan bueno seas, tanto evadiendo como acertando golpes, respecto de otros personajes como de las criaturas a las q te enfrentes."
120       vHelp(eHelp.ieInteligencia) = _
              "Influirá de manera directa en cuánto maná ganarás por nivel."
130       vHelp(eHelp.ieCarisma) = _
              "Será necesario tanto para la relación con otros personajes (entrenamiento en parties) como con las criaturas (domar animales)."
140       vHelp(eHelp.ieConstitucion) = _
              "Afectará a la cantidad de vida que podrás ganar por nivel."
150       vHelp(eHelp.ieEvasion) = "Evalúa la habilidad esquivando ataques físicos."
160       vHelp(eHelp.ieMagia) = "Puntúa la cantidad de maná que se tendrá."
170       vHelp(eHelp.ieVida) = _
              "Valora la cantidad de salud que se podrá llegar a tener."
180       vHelp(eHelp.ieEscudos) = _
              "Estima la habilidad para rechazar golpes con escudos."
190       vHelp(eHelp.ieArmas) = _
              "Evalúa la habilidad en el combate cuerpo a cuerpo con armas."
200       vHelp(eHelp.ieArcos) = _
              "Evalúa la habilidad en el combate a distancia con arcos. "
210       vHelp(eHelp.ieEspecialidad) = ""
220       vHelp(eHelp.iePuebloOrigen) = _
              "Define el hogar de tu personaje. Sin embargo, el personaje nacerá en Nemahuak, la ciudad de los novatos."
230       vHelp(eHelp.ieRaza) = _
              "De la raza que elijas dependerá cómo se modifiquen los dados que saques. Podés cambiar de raza para poder visualizar cómo se modifican los distintos atributos."
240       vHelp(eHelp.ieClase) = _
              "La clase influirá en las características principales que tenga tu personaje, asi como en las magias e items que podrá utilizar. Las estrellas que ves abajo te mostrarán en qué habilidades se destaca la misma."
250       vHelp(eHelp.ieGenero) = _
              "Indica si el personaje será masculino o femenino. Esto influye en los items que podrá equipar."
260       vHelp(eHelp.ieAlineacion) = _
              "Indica si el personaje seguirá la senda del mal o del bien. (Actualmente deshabilitado)"
End Sub

Private Sub ClearLabel()
10        LastPressed.ToggleToNormal
20        lblHelp = ""
End Sub

Private Sub txtNombre_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.ieNombre)
End Sub

Private Sub txtPasswd_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblHelp.Caption = vHelp(eHelp.iePasswd)
End Sub

Public Sub UpdateStats()
          
10        Call UpdateRazaMod
20        Call UpdateStars
End Sub

Private Sub UpdateRazaMod()
          Dim SelRaza As Integer
          Dim i As Integer
          
          
10        If lstRaza.ListIndex > -1 Then
          
20            SelRaza = lstRaza.ListIndex + 1
              
30            With ModRaza(SelRaza)
40                lblModRaza(eAtributos.Fuerza).Caption = IIf(.Fuerza >= 0, "+", "") _
                      & .Fuerza
50                lblModRaza(eAtributos.Agilidad).Caption = IIf(.Agilidad >= 0, "+", _
                      "") & .Agilidad
60                lblModRaza(eAtributos.Inteligencia).Caption = IIf(.Inteligencia >= _
                      0, "+", "") & .Inteligencia
70                lblModRaza(eAtributos.Carisma).Caption = IIf(.Carisma >= 0, "+", _
                      "") & .Carisma
80                lblModRaza(eAtributos.Constitucion).Caption = IIf(.Constitucion >= _
                      0, "+", "") & .Constitucion
90            End With
100       End If
          
          ' Atributo total
110       For i = 1 To NUMATRIBUTES
120           lblAtributoFinal(i).Caption = Val(lblAtributos(i).Caption) + _
                  Val(lblModRaza(i))
130       Next i
          
End Sub

Private Sub UpdateStars()
    
End Sub

Private Sub SetStars(ByRef ImgContainer As Object, ByVal NumStars As Integer)
          Dim FullStars As Integer
          Dim HasHalfStar As Boolean
          Dim Index As Integer
          Dim Counter As Integer

10        If NumStars > 0 Then
              
20            If NumStars > 10 Then NumStars = 10
              
30            FullStars = Int(NumStars / 2)
              
              ' Tienen brillo extra si estan todas
40            If FullStars = 5 Then
50                For Index = 1 To FullStars
60                  '  ImgContainer(Index).Picture = picGlowStar
70                Next Index
80            Else
                  ' Numero impar? Entonces hay que poner "media estrella"
90                If (NumStars Mod 2) > 0 Then HasHalfStar = True
                  
                  ' Muestro las estrellas enteras
100               If FullStars > 0 Then
110                   For Index = 1 To FullStars
120                       'ImgContainer(Index).Picture = picFullStar
130                   Next Index
                      
140                   Counter = FullStars
150               End If
                  
                  ' Muestro la mitad de la estrella (si tiene)
160               If HasHalfStar Then
170                   Counter = Counter + 1
                      
180                  ' ImgContainer(Counter).Picture = picHalfStar
190               End If
                  
                  ' Si estan completos los espacios, no borro nada
200               If Counter <> 5 Then
                      ' Limpio las que queden vacias
210                   For Index = Counter + 1 To 5
220                       'Set ImgContainer(Index).Picture = Nothing
230                   Next Index
240               End If
                  
250           End If
260       Else
              ' Limpio todo
270           For Index = 1 To 5
280               'Set ImgContainer(Index).Picture = Nothing
290           Next Index
300       End If

End Sub

Private Sub LoadCharInfo()
          Dim SearchVar As String
          Dim i As Integer
          
10        NroRazas = UBound(ListaRazas())
20        NroClases = UBound(ListaClases())

30        ReDim ModRaza(1 To NroRazas)
40        ReDim ModClase(1 To NroClases)
          
          'Modificadores de Clase
50        For i = 1 To NroClases
60            With ModClase(i)
70                SearchVar = ListaClases(i)
                  
80                .Evasion = Val(GetVar(IniPath & "CharInfo.dat", "MODEVASION", _
                      SearchVar))
90                .AtaqueArmas = Val(GetVar(IniPath & "CharInfo.dat", _
                      "MODATAQUEARMAS", SearchVar))
100               .AtaqueProyectiles = Val(GetVar(IniPath & "CharInfo.dat", _
                      "MODATAQUEPROYECTILES", SearchVar))
110               .DañoArmas = Val(GetVar(IniPath & "CharInfo.dat", "MODDAÑOARMAS", _
                      SearchVar))
120               .DañoProyectiles = Val(GetVar(IniPath & "CharInfo.dat", _
                      "MODDAÑOPROYECTILES", SearchVar))
130               .Escudo = Val(GetVar(IniPath & "CharInfo.dat", "MODESCUDO", _
                      SearchVar))
140               .Hit = Val(GetVar(IniPath & "CharInfo.dat", "HIT", SearchVar))
150               .Magia = Val(GetVar(IniPath & "CharInfo.dat", "MODMAGIA", _
                      SearchVar))
160               .Vida = Val(GetVar(IniPath & "CharInfo.dat", "MODVIDA", SearchVar))
170           End With
180       Next i
          
          'Modificadores de Raza
190       For i = 1 To NroRazas
200           With ModRaza(i)
210               SearchVar = Replace(ListaRazas(i), " ", "")
              
220               .Fuerza = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar _
                      + "Fuerza"))
230               .Agilidad = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", _
                      SearchVar + "Agilidad"))
240               .Inteligencia = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", _
                      SearchVar + "Inteligencia"))
250               .Carisma = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", _
                      SearchVar + "Carisma"))
260               .Constitucion = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", _
                      SearchVar + "Constitucion"))
270           End With
280       Next i

End Sub
