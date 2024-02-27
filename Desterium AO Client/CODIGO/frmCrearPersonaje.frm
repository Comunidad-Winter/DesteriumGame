VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearPersonaje.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox PIN 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   5520
      MaxLength       =   30
      TabIndex        =   19
      Top             =   7680
      Width           =   3375
   End
   Begin VB.ComboBox lstAlienacion 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
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
      ItemData        =   "frmCrearPersonaje.frx":1D5AE
      Left            =   -2520
      List            =   "frmCrearPersonaje.frx":1D5B8
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.TextBox txtMail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
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
      Height          =   225
      Left            =   5520
      MaxLength       =   100
      TabIndex        =   3
      Top             =   7200
      Width           =   3375
   End
   Begin VB.TextBox txtConfirmPasswd 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   5520
      MaxLength       =   40
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   6600
      Width           =   3375
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   5520
      MaxLength       =   40
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   6120
      Width           =   3375
   End
   Begin VB.Timer tAnimacion 
      Enabled         =   0   'False
      Left            =   840
      Top             =   1080
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
      ItemData        =   "frmCrearPersonaje.frx":1D5CB
      Left            =   1560
      List            =   "frmCrearPersonaje.frx":1D5CD
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3525
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
      ItemData        =   "frmCrearPersonaje.frx":1D5CF
      Left            =   1560
      List            =   "frmCrearPersonaje.frx":1D5D9
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4440
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
      ItemData        =   "frmCrearPersonaje.frx":1D5EC
      Left            =   1560
      List            =   "frmCrearPersonaje.frx":1D5EE
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2520
      Width           =   2265
   End
   Begin VB.ComboBox lstHogar 
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
      ItemData        =   "frmCrearPersonaje.frx":1D5F0
      Left            =   1560
      List            =   "frmCrearPersonaje.frx":1D5F2
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1680
      Width           =   2265
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000012&
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
      Height          =   345
      Left            =   5520
      MaxLength       =   15
      TabIndex        =   0
      Top             =   5640
      Width           =   3375
   End
   Begin VB.PictureBox picPJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   -2520
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   9
      Top             =   6480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   -2520
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   10
      Top             =   6480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   1
      Left            =   -405
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   13
      Top             =   6120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   2
      Left            =   -480
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   14
      Top             =   6120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   3
      Left            =   -2475
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   15
      Top             =   6000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   4
      Left            =   -2520
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   16
      Top             =   6000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   -810
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   12
      Top             =   6120
      Visible         =   0   'False
      Width           =   360
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
      Left            =   8640
      TabIndex        =   34
      Top             =   1080
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
      Left            =   8640
      TabIndex        =   33
      Top             =   1395
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
      Left            =   8640
      TabIndex        =   32
      Top             =   2400
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
      Left            =   8640
      TabIndex        =   31
      Top             =   1740
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
      Index           =   4
      Left            =   8640
      TabIndex        =   30
      Top             =   2070
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
      Index           =   5
      Left            =   5595
      TabIndex        =   29
      Top             =   2250
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
      Left            =   5595
      TabIndex        =   28
      Top             =   1080
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
      Left            =   5595
      TabIndex        =   27
      Top             =   1365
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
      Left            =   5595
      TabIndex        =   26
      Top             =   1665
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
      Left            =   5595
      TabIndex        =   25
      Top             =   1950
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
      Left            =   4800
      TabIndex        =   24
      Top             =   1080
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
      Left            =   4800
      TabIndex        =   23
      Top             =   1365
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
      Left            =   4800
      TabIndex        =   22
      Top             =   1665
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
      Left            =   4800
      TabIndex        =   21
      Top             =   1950
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
      Left            =   4800
      TabIndex        =   20
      Top             =   2250
      Width           =   225
   End
   Begin VB.Image PClase 
      Height          =   3150
      Left            =   1320
      Picture         =   "frmCrearPersonaje.frx":1D5F4
      Top             =   5160
      Width           =   2505
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2175
      Left            =   6690
      TabIndex        =   18
      Top             =   2865
      Width           =   3855
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   5
      Left            =   1680
      Top             =   7020
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   4
      Left            =   1455
      Top             =   7020
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   3
      Left            =   1230
      Top             =   7020
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   2
      Left            =   -1515
      Top             =   6900
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   1
      Left            =   -1740
      Top             =   6900
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   5
      Left            =   1680
      Top             =   6720
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   4
      Left            =   1455
      Top             =   6720
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   3
      Left            =   1230
      Top             =   6720
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   2
      Left            =   -1515
      Top             =   6600
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   5
      Left            =   1080
      Top             =   6435
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   4
      Left            =   -1665
      Top             =   6315
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   3
      Left            =   -1890
      Top             =   6315
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   2
      Left            =   -2115
      Top             =   6315
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   5
      Left            =   1080
      Top             =   6150
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   4
      Left            =   -2400
      Top             =   6000
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   3
      Left            =   -1890
      Top             =   6030
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   2
      Left            =   -2115
      Top             =   6030
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   5
      Left            =   1080
      Top             =   5865
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   4
      Left            =   -1665
      Top             =   5745
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   3
      Left            =   -1890
      Top             =   5745
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   2
      Left            =   -2115
      Top             =   5745
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   1
      Left            =   -1740
      Top             =   6600
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   1
      Left            =   -2340
      Top             =   6315
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   1
      Left            =   -2340
      Top             =   6030
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   1
      Left            =   -2340
      Top             =   5745
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   5
      Left            =   1080
      Top             =   5580
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   4
      Left            =   -1665
      Top             =   5460
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   3
      Left            =   -1890
      Top             =   5460
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   2
      Left            =   -2115
      Top             =   5460
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   1
      Left            =   -2340
      Top             =   5460
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblEspecialidad 
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   -1800
      TabIndex        =   17
      Top             =   7155
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   3
      Visible         =   0   'False
      X1              =   -33
      X2              =   -8
      Y1              =   432
      Y2              =   432
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   2
      Visible         =   0   'False
      X1              =   -33
      X2              =   -8
      Y1              =   407
      Y2              =   407
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   1
      Visible         =   0   'False
      X1              =   -8
      X2              =   -8
      Y1              =   408
      Y2              =   432
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   0
      Visible         =   0   'False
      X1              =   -161
      X2              =   -161
      Y1              =   400
      Y2              =   424
   End
   Begin VB.Image imgAtributos 
      Height          =   270
      Left            =   7440
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   4230
      TabIndex        =   11
      Top             =   2835
      Width           =   2055
   End
   Begin VB.Image imgVolver 
      Height          =   645
      Left            =   -240
      Top             =   8400
      Width           =   2010
   End
   Begin VB.Image imgCrear 
      Height          =   525
      Left            =   10320
      Top             =   8520
      Width           =   1935
   End
   Begin VB.Image imgalineacion 
      Height          =   240
      Left            =   1440
      Top             =   5040
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image imgGenero 
      Height          =   240
      Left            =   2280
      Top             =   4080
      Width           =   675
   End
   Begin VB.Image imgClase 
      Height          =   240
      Left            =   2400
      Top             =   3120
      Width           =   555
   End
   Begin VB.Image imgRaza 
      Height          =   255
      Left            =   2400
      Top             =   2160
      Width           =   450
   End
   Begin VB.Image imgPuebloOrigen 
      Height          =   225
      Left            =   1800
      Top             =   1320
      Width           =   1875
   End
   Begin VB.Image imgEspecialidad 
      Height          =   240
      Left            =   -2910
      Top             =   7170
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image imgArcos 
      Height          =   225
      Left            =   -2895
      Top             =   6900
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image imgArmas 
      Height          =   240
      Left            =   -2910
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgEscudos 
      Height          =   255
      Left            =   -1005
      Top             =   6420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgVida 
      Height          =   225
      Left            =   -2910
      Top             =   6030
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image imgMagia 
      Height          =   255
      Left            =   -2955
      Top             =   5715
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgEvasion 
      Height          =   255
      Left            =   -1035
      Top             =   5550
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgConstitucion 
      Height          =   255
      Left            =   6960
      Top             =   2400
      Width           =   960
   End
   Begin VB.Image imgCarisma 
      Height          =   240
      Left            =   6840
      Top             =   2040
      Width           =   765
   End
   Begin VB.Image imgInteligencia 
      Height          =   240
      Left            =   6840
      Top             =   1710
      Width           =   1005
   End
   Begin VB.Image imgAgilidad 
      Height          =   240
      Left            =   6720
      Top             =   1350
      Width           =   735
   End
   Begin VB.Image imgFuerza 
      Height          =   240
      Left            =   6840
      Top             =   1080
      Width           =   675
   End
   Begin VB.Image imgF 
      Height          =   270
      Left            =   5520
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgM 
      Height          =   270
      Left            =   4800
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgD 
      Height          =   270
      Left            =   5280
      Top             =   1800
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgConfirmPasswd 
      Height          =   375
      Left            =   4080
      Top             =   6600
      Width           =   1320
   End
   Begin VB.Image imgPasswd 
      Height          =   375
      Left            =   4080
      Top             =   6120
      Width           =   1245
   End
   Begin VB.Image imgNombre 
      Height          =   360
      Left            =   4080
      Top             =   5640
      Width           =   1275
   End
   Begin VB.Image imgMail 
      Height          =   360
      Left            =   3960
      Top             =   7080
      Width           =   1515
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   1
      Left            =   -2160
      Picture         =   "frmCrearPersonaje.frx":21DF9
      Top             =   7440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   0
      Left            =   -2520
      Picture         =   "frmCrearPersonaje.frx":2210B
      Top             =   7440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image HeadPJ 
      Height          =   225
      Index           =   1
      Left            =   -2400
      Top             =   5640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image HeadPJ 
      Height          =   225
      Index           =   0
      Left            =   -1125
      Picture         =   "frmCrearPersonaje.frx":2241D
      Top             =   6165
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   3120
      Left            =   -3840
      Stretch         =   -1  'True
      Top             =   4560
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Image imgDados 
      Height          =   1485
      Left            =   9240
      MouseIcon       =   "frmCrearPersonaje.frx":2272F
      MousePointer    =   99  'Custom
      Top             =   960
      Width           =   1500
   End
   Begin VB.Image imgHogar 
      Height          =   2850
      Left            =   5640
      Picture         =   "frmCrearPersonaje.frx":22881
      Top             =   9120
      Visible         =   0   'False
      Width           =   2985
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Developed AO 13.0
'FrmCrearPersonaje

Option Explicit

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

    PClase = LoadPicture(App.path & "\Recursos\Clases\Clerigo.jpg")
    UserClase = eClass.Cleric
    Cargando = True
    Call LoadCharInfo
    Call CargarEspecialidades

    Call IniciarGraficos
    Call CargarCombos

    Call LoadHelp

    Call DrawImageInPicture(picPJ, Me.Picture, 0, 0, , , picPJ.Left, picPJ.Top)
    Dir = SOUTH

    Call TirarDados

    Cargando = False

    'UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    UserHead = 0

End Sub

Private Sub CargarEspecialidades()

    ReDim vEspecialidades(1 To NroClases)

    vEspecialidades(eClass.Hunter) = "Ocultarse"
    vEspecialidades(eClass.Thief) = "Robar y Ocultarse"
    vEspecialidades(eClass.Assasin) = "Apuñalar"
    vEspecialidades(eClass.Bandit) = "Combate Sin Armas"
    vEspecialidades(eClass.Druid) = "Domar"
    vEspecialidades(eClass.Pirat) = "Navegar"
    vEspecialidades(eClass.Worker) = "Extracción y Construcción"
End Sub

Private Sub IniciarGraficos()

    Dim GrhPath As String
    GrhPath = DirGraficos

    Set LastPressed = New clsGraphicalButton

    Set picFullStar = LoadPicture(GrhPath & "EstrellaSimple.jpg")
    Set picHalfStar = LoadPicture(GrhPath & "EstrellaMitad.jpg")
    Set picGlowStar = LoadPicture(GrhPath & "EstrellaBrillante.jpg")

End Sub

Private Sub CargarCombos()
    Dim i As Integer

    lstProfesion.Clear

    For i = LBound(ListaClases) To NroClases
        lstProfesion.AddItem ListaClases(i)
    Next i

    lstProfesion.ListIndex = 1
    lstHogar.Clear

    lstHogar.AddItem Ciudades(1)

    lstRaza.Clear

    For i = LBound(ListaRazas()) To NroRazas
        lstRaza.AddItem ListaRazas(i)
    Next i

    'lstProfesion.ListIndex = 1
End Sub

Function CheckData() As Boolean

    If txtPasswd.Text <> txtConfirmPasswd.Text Then
        MsgBox "Los passwords que tipeo no coinciden, por favor vuelva a ingresarlos."
        Exit Function
    End If

    If Not CheckMailString(txtMail.Text) Then
        MsgBox "Direccion de mail invalida."
        Exit Function
    End If

    If UserRaza = 0 Then
        MsgBox "Seleccione la raza del personaje."
        Exit Function
    End If

    If UserSexo = 0 Then
        MsgBox "Seleccione el sexo del personaje."
        Exit Function
    End If

    If UserClase = 0 Then
        MsgBox "Seleccione la clase del personaje."
        Exit Function
    End If

    If UserHogar = 0 Then
        MsgBox "Seleccione el hogar del personaje."
        Exit Function
    End If

    Dim i As Long

    For i = 1 To NUMATRIBUTOS
        If UserAtributos(i) = 0 Then
            MsgBox "Los atributos del personaje son invalidos."
            Exit Function
        End If
    Next i

    If Len(txtPasswd.Text) < 4 Then
        MsgBox ("La contraseña debe de tener mas de 4 caracteres!!")
        Exit Function
    End If

    If Len(PIN.Text) < 4 Then
        MsgBox ("El pin debe de tener mas de 4 caracteres!!")
        Exit Function
    End If

    If PIN.Text = txtPasswd.Text Then
        MsgBox "Tu Pin no puede ser igual a tu password"
        Exit Function
    End If

    CheckData = True

End Function

Private Sub TirarDados()
    Call WriteThrowDices
    Call FlushBuffer
End Sub

Private Sub DirPJ_Click(Index As Integer)
    Select Case Index
        Case 0
            Dir = CheckDir(Dir + 1)
        Case 1
            Dir = CheckDir(Dir - 1)
    End Select

    Call UpdateHeadSelection
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ClearLabel
End Sub

Private Sub HeadPJ_Click(Index As Integer)
    Select Case Index
        Case 0
            UserHead = CheckCabeza(UserHead + 1)
        Case 1
            UserHead = CheckCabeza(UserHead - 1)
    End Select

    Call UpdateHeadSelection

End Sub

Private Sub UpdateHeadSelection()
    Dim Head As Integer

    Head = UserHead
    Call DrawHead(Head, 2)

    Head = Head + 1
    Call DrawHead(CheckCabeza(Head), 3)

    Head = Head + 1
    Call DrawHead(CheckCabeza(Head), 4)

    Head = UserHead

    Head = Head - 1
    Call DrawHead(CheckCabeza(Head), 1)

    Head = Head - 1
    Call DrawHead(CheckCabeza(Head), 0)
End Sub

Private Sub imgCrear_Click()

    Dim i As Integer
    Dim CharAscii As Byte

    UserName = txtNombre.Text
 
    UserName = Trim$(UserName)
    
    UserRaza = lstRaza.ListIndex + 1
    UserSexo = lstGenero.ListIndex + 1
    UserClase = lstProfesion.ListIndex + 1

    For i = 1 To NUMATRIBUTES
        UserAtributos(i) = Val(lblAtributos(i).Caption)
    Next i

    UserHogar = lstHogar.ListIndex + 1

    If Not CheckData Then Exit Sub

    UserPassword = txtPasswd.Text

    For i = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, i, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Sub
        End If
    Next i

    UserEmail = txtMail.Text
    UserPin = PIN

    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort

    EstadoLogin = E_MODO.CrearNuevoPj

    If Not frmMain.Socket1.Connected Then
        MsgBox "ERROR" & vbCrLf & "Se ha perdido la conexión con el servidor.", vbCritical
        frmConnect.Visible = True
        Unload Me
    Else
        Call Login
    End If

End Sub

Private Sub imgDados_Click()
    Call Audio.PlayWave(SND_DICE)
    Call TirarDados
End Sub

Private Sub imgEspecialidad_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEspecialidad)
End Sub

Private Sub imgNombre_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieNombre)
End Sub

Private Sub imgPasswd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.iePasswd)
End Sub

Private Sub imgConfirmPasswd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieConfirmPasswd)
End Sub

Private Sub imgAtributos_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieAtributos)
End Sub

Private Sub imgD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieD)
End Sub

Private Sub imgM_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieM)
End Sub

Private Sub imgF_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieF)
End Sub

Private Sub imgFuerza_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieFuerza)
End Sub

Private Sub imgAgilidad_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieAgilidad)
End Sub

Private Sub imgInteligencia_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieInteligencia)
End Sub

Private Sub imgCarisma_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieCarisma)
End Sub

Private Sub imgConstitucion_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieConstitucion)
End Sub

Private Sub imgArcos_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieArcos)
End Sub

Private Sub imgArmas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieArmas)
End Sub

Private Sub imgEscudos_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEscudos)
End Sub

Private Sub imgEvasion_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEvasion)
End Sub

Private Sub imgMagia_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieMagia)
End Sub

Private Sub imgMail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieMail)
End Sub

Private Sub imgVida_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieVida)
End Sub

Private Sub imgPuebloOrigen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.iePuebloOrigen)
End Sub

Private Sub imgRaza_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieRaza)
End Sub

Private Sub imgClase_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieClase)
End Sub

Private Sub imgGenero_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieGenero)
End Sub

Private Sub imgalineacion_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieAlineacion)
End Sub

Private Sub imgVolver_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call Audio.PlayMIDI("2.mid")
    frmConnect.Visible = True
    Unload Me
End Sub

Private Sub lstGenero_Click()
    UserSexo = lstGenero.ListIndex + 1
  '  Call DarCuerpoYCabeza
End Sub

Private Sub lstProfesion_Click()

    On Error Resume Next
    '    Image1.Picture = LoadPicture(App.path & "\graficos\" & lstProfesion.Text & ".jpg")
    '
    UserClase = lstProfesion.ListIndex + 1

    Call UpdateStats
    Call UpdateEspecialidad(UserClase)

    Select Case UserClase
        Case 1
            PClase = LoadPicture(App.path & "\Recursos\Clases\Mago.jpg")
            Label1 = "Mago es la clase por excelencia preferida por muchos de los jugadores. Si te interesan mucho los hechizos, encantamientos e invocaciones, esta clase es la indicada, ya que posee mucha maná."
        Case 2
            PClase = LoadPicture(App.path & "\Recursos\Clases\Clerigo.jpg")
            Label1 = "Clerigo es una clase que se especializa tanto en las artes mágicas como la lucha de cuerpo a cuerpo. Aunque sus golpes no sean tan fuertes puede combinar sus dos especialidades."
        Case 3
            PClase = LoadPicture(App.path & "\Recursos\Clases\Guerrero.jpg")
            Label1 = "Guerrero es una clase que no usa magias sino que directamente usa la fuerza y combate cuerpo a cuerpo, sus golpes resultan ser increíblemente devastadores cuando éste se encuentra en su punto más elevado, posee una gran vida pero su desventaja es que no puede usar magias."
        Case 4
            PClase = LoadPicture(App.path & "\Recursos\Clases\Asesino.jpg")
            Label1 = "Asesino es una clase Sigilosa y sanguinaria, la característica especial de esta clase es APUÑALAR, esto significa que de un golpe podrías dejar a tu enemigo prácticamente muerto, su evasión es la mejor y no conoce el miedo contra quienes intenten pegarle."
        Case 5
            PClase = LoadPicture(App.path & "\Recursos\Clases\Ladron.jpg")
            Label1 = "Ladron es una clase que pueden robar gran cantidad de objetos y oro al enemigo casi sin ser detectado. Dominan el arte del sigilo,al igual que los bandidos pueden caminar oculto entre las sombras sin ser detectado."
        Case 6
            PClase = LoadPicture(App.path & "\Recursos\Clases\Bardo.jpg")
            Label1 = "Bardo es una clase muy práctica frente a las clases de cuerpo a cuerpo ya que posee una gran evasión lo que resulta difícil para enemigo cuando intente acertarle un golpe."
        Case 7
            PClase = LoadPicture(App.path & "\Recursos\Clases\Druida.jpg")
            Label1 = "Druida es una clase que se especializa en domar criaturas y usarlas como mascotas, también usa conjuros para invocar otros tipos de criaturas que acuden a su ayuda, lo que permite que en su entrenamiento nunca esté solo."
        Case 8
            PClase = LoadPicture(App.path & "\Recursos\Clases\Bandido.jpg")
            Label1 = "Bandido es una clase prima-hermana del Ladròn solo que esta preparada para Luchar, se pueden ocultar muy bien, solo con puños dan un gran golpe."
        Case 9
            PClase = LoadPicture(App.path & "\Recursos\Clases\Paladin.jpg")
            Label1 = "Paladín es una clase con mucha vida y una gran fuerza, ideal para encuentro contra otra clase que lucha cuerpo a cuerpo ya que sus golpes son algo más débiles que los del guerrero."
        Case 10
            PClase = LoadPicture(App.path & "\Recursos\Clases\Cazador.jpg")
            Label1 = "Cazador es una clase que no usa magia, pero que es muy hábil usando armas a distancias, tiene la habilidad de poder ocultarse entre las sombras cuando éste usa su armadura de cazador."
        Case 11
            PClase = LoadPicture(App.path & "\Recursos\Clases\Plebeyo.jpg")
            Label1 = "Trabajador son fieles servidores capaces de elaborar artesanías con un poder extraordinario, son dedicados exclusivamente a la extracción de materia prima y creación de objetos de gran valor. Son expertos en actividades tales como la pesca, minería, tala, herrería y carpintería."
        Case 12
            PClase = LoadPicture(App.path & "\Recursos\Clases\Pirata.jpg")
            Label1 = "Pirata es una clase Reyes de los mares, los piratas aprenden a navegar más rápidamente que las demás clases de combate. Carece de maná y no son la mejor clase cuerpo-a-cuerpo, pero son indispensables en las líneas defensivas."
    End Select
    
    End Sub

Private Sub UpdateEspecialidad(ByVal eClase As eClass)
    lblEspecialidad.Caption = vEspecialidades(eClase)
End Sub

Private Sub lstRaza_Click()
    UserRaza = lstRaza.ListIndex + 1
    Call DarCuerpoYCabeza

    Call UpdateStats
End Sub

Private Sub picHead_Click(Index As Integer)
' No se mueve si clickea al medio
    If Index = 2 Then Exit Sub

    Dim Counter As Integer
    Dim Head As Integer

    Head = UserHead

    If Index > 2 Then
        For Counter = Index - 2 To 1 Step -1
            Head = CheckCabeza(Head + 1)
        Next Counter
    Else
        For Counter = 2 - Index To 1 Step -1
            Head = CheckCabeza(Head - 1)
        Next Counter
    End If

    UserHead = Head

    Call UpdateHeadSelection

End Sub

Private Sub tAnimacion_Timer()
    Dim SR As RECT
    Dim DR As RECT
    Dim Grh As Long
    Static Frame As Byte

    If currentGrh = 0 Then Exit Sub
    UserHead = CheckCabeza(UserHead)

    Frame = Frame + 1
    If Frame >= GrhData(currentGrh).NumFrames Then Frame = 1
    Call DrawImageInPicture(picPJ, Me.Picture, 0, 0, , , picPJ.Left, picPJ.Top)

    Grh = GrhData(currentGrh).Frames(Frame)

    With GrhData(Grh)
        SR.Left = .sX
        SR.Top = .sY
        SR.Right = SR.Left + .pixelWidth
        SR.Bottom = SR.Top + .pixelHeight

        DR.Left = (picPJ.Width - .pixelWidth) \ 2 - 2
        DR.Top = (picPJ.Height - .pixelHeight) \ 2 - 2
        DR.Right = DR.Left + .pixelWidth
        DR.Bottom = DR.Top + .pixelHeight

        picTemp.BackColor = picTemp.BackColor

        Call DrawGrhtoHdc(picTemp.hdc, Grh, SR, DR)
        Call DrawTransparentGrhtoHdc(picPJ.hdc, picTemp.hdc, DR, DR, vbBlack)
    End With

    Grh = HeadData(UserHead).Head(Dir).GrhIndex

    With GrhData(Grh)
        SR.Left = .sX
        SR.Top = .sY
        SR.Right = SR.Left + .pixelWidth
        SR.Bottom = SR.Top + .pixelHeight

        DR.Left = (picPJ.Width - .pixelWidth) \ 2 - 2
        DR.Top = DR.Bottom + BodyData(UserBody).HeadOffset.y - .pixelHeight
        DR.Right = DR.Left + .pixelWidth
        DR.Bottom = DR.Top + .pixelHeight

        picTemp.BackColor = picTemp.BackColor

        Call DrawGrhtoHdc(picTemp.hdc, Grh, SR, DR)
        Call DrawTransparentGrhtoHdc(picPJ.hdc, picTemp.hdc, DR, DR, vbBlack)
    End With
End Sub

Private Sub DrawHead(ByVal Head As Integer, ByVal PicIndex As Integer)

    Dim SR As RECT
    Dim DR As RECT
    Dim Grh As Long

    Call DrawImageInPicture(picHead(PicIndex), Me.Picture, 0, 0, , , picHead(PicIndex).Left, picHead(PicIndex).Top)

    Grh = HeadData(Head).Head(Dir).GrhIndex

    With GrhData(Grh)
        SR.Left = .sX
        SR.Top = .sY
        SR.Right = SR.Left + .pixelWidth
        SR.Bottom = SR.Top + .pixelHeight

        DR.Left = (picHead(0).Width - .pixelWidth) \ 2 + 1
        DR.Top = 0
        DR.Right = DR.Left + .pixelWidth
        DR.Bottom = DR.Top + .pixelHeight

        picTemp.BackColor = picTemp.BackColor

        Call DrawGrhtoHdc(picTemp.hdc, Grh, SR, DR)
        Call DrawTransparentGrhtoHdc(picHead(PicIndex).hdc, picTemp.hdc, DR, DR, vbBlack)
    End With

End Sub

Private Sub txtConfirmPasswd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieConfirmPasswd)
End Sub

Private Sub txtMail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieMail)
End Sub

Private Sub DarCuerpoYCabeza()

    Dim bVisible As Boolean
    Dim PicIndex As Integer
    Dim LineIndex As Integer

    Select Case UserSexo
        Case eGenero.Hombre
            Select Case UserRaza
                Case eRaza.Humano
                    UserHead = HUMANO_H_PRIMER_CABEZA
                    UserBody = HUMANO_H_CUERPO_DESNUDO

                Case eRaza.Elfo
                    UserHead = ELFO_H_PRIMER_CABEZA
                    UserBody = ELFO_H_CUERPO_DESNUDO

                Case eRaza.ElfoOscuro
                    UserHead = DROW_H_PRIMER_CABEZA
                    UserBody = DROW_H_CUERPO_DESNUDO

                Case eRaza.Enano
                    UserHead = ENANO_H_PRIMER_CABEZA
                    UserBody = ENANO_H_CUERPO_DESNUDO

                Case eRaza.Gnomo
                    UserHead = GNOMO_H_PRIMER_CABEZA
                    UserBody = GNOMO_H_CUERPO_DESNUDO

                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select

        Case eGenero.Mujer
            Select Case UserRaza
                Case eRaza.Humano
                    UserHead = HUMANO_M_PRIMER_CABEZA
                    UserBody = HUMANO_M_CUERPO_DESNUDO

                Case eRaza.Elfo
                    UserHead = ELFO_M_PRIMER_CABEZA
                    UserBody = ELFO_M_CUERPO_DESNUDO

                Case eRaza.ElfoOscuro
                    UserHead = DROW_M_PRIMER_CABEZA
                    UserBody = DROW_M_CUERPO_DESNUDO

                Case eRaza.Enano
                    UserHead = ENANO_M_PRIMER_CABEZA
                    UserBody = ENANO_M_CUERPO_DESNUDO

                Case eRaza.Gnomo
                    UserHead = GNOMO_M_PRIMER_CABEZA
                    UserBody = GNOMO_M_CUERPO_DESNUDO

                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
        Case Else
            UserHead = 0
            UserBody = 0
    End Select

    bVisible = UserHead <> 0 And UserBody <> 0

    HeadPJ(0).Visible = bVisible
    HeadPJ(1).Visible = bVisible
    DirPJ(0).Visible = bVisible
    DirPJ(1).Visible = bVisible

    For PicIndex = 0 To 4
        picHead(PicIndex).Visible = bVisible
    Next PicIndex

    For LineIndex = 0 To 3
        Line1(LineIndex).Visible = bVisible
    Next LineIndex

    'If bVisible Then Call UpdateHeadSelection

   ' currentGrh = BodyData(UserBody).Walk(Dir).GrhIndex
   ' If currentGrh > 0 Then _
   '    tAnimacion.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)
End Sub

Private Function CheckCabeza(ByVal Head As Integer) As Integer

    Select Case UserSexo
        Case eGenero.Hombre
            Select Case UserRaza
                Case eRaza.Humano
                    If Head > HUMANO_H_ULTIMA_CABEZA Then
                        CheckCabeza = HUMANO_H_PRIMER_CABEZA + (Head - HUMANO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < HUMANO_H_PRIMER_CABEZA Then
                        CheckCabeza = HUMANO_H_ULTIMA_CABEZA - (HUMANO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If

                Case eRaza.Elfo
                    If Head > ELFO_H_ULTIMA_CABEZA Then
                        CheckCabeza = ELFO_H_PRIMER_CABEZA + (Head - ELFO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < ELFO_H_PRIMER_CABEZA Then
                        CheckCabeza = ELFO_H_ULTIMA_CABEZA - (ELFO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If

                Case eRaza.ElfoOscuro
                    If Head > DROW_H_ULTIMA_CABEZA Then
                        CheckCabeza = DROW_H_PRIMER_CABEZA + (Head - DROW_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < DROW_H_PRIMER_CABEZA Then
                        CheckCabeza = DROW_H_ULTIMA_CABEZA - (DROW_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If

                Case eRaza.Enano
                    If Head > ENANO_H_ULTIMA_CABEZA Then
                        CheckCabeza = ENANO_H_PRIMER_CABEZA + (Head - ENANO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < ENANO_H_PRIMER_CABEZA Then
                        CheckCabeza = ENANO_H_ULTIMA_CABEZA - (ENANO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If

                Case eRaza.Gnomo
                    If Head > GNOMO_H_ULTIMA_CABEZA Then
                        CheckCabeza = GNOMO_H_PRIMER_CABEZA + (Head - GNOMO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < GNOMO_H_PRIMER_CABEZA Then
                        CheckCabeza = GNOMO_H_ULTIMA_CABEZA - (GNOMO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If

                Case Else
                    UserRaza = lstRaza.ListIndex + 1
                    CheckCabeza = CheckCabeza(Head)
            End Select

        Case eGenero.Mujer
            Select Case UserRaza
                Case eRaza.Humano
                    If Head > HUMANO_M_ULTIMA_CABEZA Then
                        CheckCabeza = HUMANO_M_PRIMER_CABEZA + (Head - HUMANO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < HUMANO_M_PRIMER_CABEZA Then
                        CheckCabeza = HUMANO_M_ULTIMA_CABEZA - (HUMANO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If

                Case eRaza.Elfo
                    If Head > ELFO_M_ULTIMA_CABEZA Then
                        CheckCabeza = ELFO_M_PRIMER_CABEZA + (Head - ELFO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < ELFO_M_PRIMER_CABEZA Then
                        CheckCabeza = ELFO_M_ULTIMA_CABEZA - (ELFO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If

                Case eRaza.ElfoOscuro
                    If Head > DROW_M_ULTIMA_CABEZA Then
                        CheckCabeza = DROW_M_PRIMER_CABEZA + (Head - DROW_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < DROW_M_PRIMER_CABEZA Then
                        CheckCabeza = DROW_M_ULTIMA_CABEZA - (DROW_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If

                Case eRaza.Enano
                    If Head > ENANO_M_ULTIMA_CABEZA Then
                        CheckCabeza = ENANO_M_PRIMER_CABEZA + (Head - ENANO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < ENANO_M_PRIMER_CABEZA Then
                        CheckCabeza = ENANO_M_ULTIMA_CABEZA - (ENANO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If

                Case eRaza.Gnomo
                    If Head > GNOMO_M_ULTIMA_CABEZA Then
                        CheckCabeza = GNOMO_M_PRIMER_CABEZA + (Head - GNOMO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < GNOMO_M_PRIMER_CABEZA Then
                        CheckCabeza = GNOMO_M_ULTIMA_CABEZA - (GNOMO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If

                Case Else
                    UserRaza = lstRaza.ListIndex + 1
                    CheckCabeza = CheckCabeza(Head)
            End Select
        Case Else
            UserSexo = lstGenero.ListIndex + 1
            CheckCabeza = CheckCabeza(Head)
    End Select
End Function

Private Function CheckDir(ByRef Dir As E_Heading) As E_Heading

    If Dir > E_Heading.WEST Then Dir = E_Heading.NORTH
    If Dir < E_Heading.NORTH Then Dir = E_Heading.WEST

    CheckDir = Dir

    currentGrh = BodyData(UserBody).Walk(Dir).GrhIndex
    If currentGrh > 0 Then _
       tAnimacion.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)

End Function

Private Sub LoadHelp()
    vHelp(eHelp.iePasswd) = "La contraseña que utilizarás para conectar tu personaje al juego."
    vHelp(eHelp.ieTirarDados) = "Presionando sobre la Esfera Roja, se modificarán al azar los atributos de tu personaje, de esta manera puedes elegir los que más te parezcan para definir a tu personaje."
    vHelp(eHelp.ieMail) = "Es sumamente importante que ingreses una dirección de correo electrónico válida, ya que en el caso de perder la contraseña de tu personaje, se te enviará cuando lo requieras, a esa dirección."
    vHelp(eHelp.ieNombre) = "Sé cuidadoso al seleccionar el nombre de tu personaje. Developed AO es un juego de rol, un mundo mágico y fantástico, y si seleccionás un nombre obsceno o con connotación política, los administradores borrarán tu personaje y no habrá ninguna posibilidad de recuperarlo."
    vHelp(eHelp.ieConfirmPasswd) = "La contraseña que utilizarás para conectar tu personaje al juego."
    vHelp(eHelp.ieAtributos) = "Son las cualidades que definen tu personaje. Generalmente se los llama ""Dados"". (Ver Tirar Dados)"
    vHelp(eHelp.ieD) = "Son los atributos que obtuviste al azar. Presioná la esfera roja para volver a tirarlos."
    vHelp(eHelp.ieM) = "Son los modificadores por raza que influyen en los atributos de tu personaje."
    vHelp(eHelp.ieF) = "Los atributos finales de tu personaje, de acuerdo a la raza que elegiste."
    vHelp(eHelp.ieFuerza) = "De ella dependerá qué tan potentes serán tus golpes, tanto con armas de cuerpo a cuerpo, a distancia o sin armas."
    vHelp(eHelp.ieAgilidad) = "Este atributo intervendrá en qué tan bueno seas, tanto evadiendo como acertando golpes, respecto de otros personajes como de las criaturas a las q te enfrentes."
    vHelp(eHelp.ieInteligencia) = "Influirá de manera directa en cuánto maná ganarás por nivel."
    vHelp(eHelp.ieCarisma) = "Será necesario tanto para la relación con otros personajes (entrenamiento en parties) como con las criaturas (domar animales)."
    vHelp(eHelp.ieConstitucion) = "Afectará a la cantidad de vida que podrás ganar por nivel."
    vHelp(eHelp.ieEvasion) = "Evalúa la habilidad esquivando ataques físicos."
    vHelp(eHelp.ieMagia) = "Puntúa la cantidad de maná que se tendrá."
    vHelp(eHelp.ieVida) = "Valora la cantidad de salud que se podrá llegar a tener."
    vHelp(eHelp.ieEscudos) = "Estima la habilidad para rechazar golpes con escudos."
    vHelp(eHelp.ieArmas) = "Evalúa la habilidad en el combate cuerpo a cuerpo con armas."
    vHelp(eHelp.ieArcos) = "Evalúa la habilidad en el combate a distancia con arcos. "
    vHelp(eHelp.ieEspecialidad) = ""
    vHelp(eHelp.iePuebloOrigen) = "Define el hogar de tu personaje. Sin embargo, el personaje nacerá en Nemahuak, la ciudad de los novatos."
    vHelp(eHelp.ieRaza) = "De la raza que elijas dependerá cómo se modifiquen los dados que saques. Podés cambiar de raza para poder visualizar cómo se modifican los distintos atributos."
    vHelp(eHelp.ieClase) = "La clase influirá en las características principales que tenga tu personaje, asi como en las magias e items que podrá utilizar. Las estrellas que ves abajo te mostrarán en qué habilidades se destaca la misma."
    vHelp(eHelp.ieGenero) = "Indica si el personaje será masculino o femenino. Esto influye en los items que podrá equipar."
    vHelp(eHelp.ieAlineacion) = "Indica si el personaje seguirá la senda del mal o del bien. (Actualmente deshabilitado)"
End Sub

Private Sub ClearLabel()
    LastPressed.ToggleToNormal
    lblHelp = ""
End Sub

Private Sub txtNombre_Change()
    txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.ieNombre)
End Sub

Private Sub txtPasswd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.Caption = vHelp(eHelp.iePasswd)
End Sub

Public Sub UpdateStats()

    Call UpdateRazaMod
    Call UpdateStars
End Sub

Private Sub UpdateRazaMod()
    Dim SelRaza As Integer
    Dim i As Integer


    If lstRaza.ListIndex > -1 Then

        SelRaza = lstRaza.ListIndex + 1

        With ModRaza(SelRaza)
            lblModRaza(eAtributos.Fuerza).Caption = IIf(.Fuerza >= 0, "+", "") & .Fuerza
            lblModRaza(eAtributos.Agilidad).Caption = IIf(.Agilidad >= 0, "+", "") & .Agilidad
            lblModRaza(eAtributos.Inteligencia).Caption = IIf(.Inteligencia >= 0, "+", "") & .Inteligencia
            lblModRaza(eAtributos.Carisma).Caption = IIf(.Carisma >= 0, "+", "") & .Carisma
            lblModRaza(eAtributos.Constitucion).Caption = IIf(.Constitucion >= 0, "+", "") & .Constitucion
        End With
    End If

    ' Atributo total
    For i = 1 To NUMATRIBUTES
        lblAtributoFinal(i).Caption = Val(lblAtributos(i).Caption) + Val(lblModRaza(i))
    Next i

End Sub

Private Sub UpdateStars()
    Dim NumStars As Double

    If UserClase = 0 Then Exit Sub

    ' Estrellas de evasion
    NumStars = (2.454 + 0.073 * Val(lblAtributoFinal(eAtributos.Agilidad).Caption)) * ModClase(UserClase).Evasion
    Call SetStars(imgEvasionStar, NumStars * 2)

    ' Estrellas de magia
    NumStars = ModClase(UserClase).Magia * Val(lblAtributoFinal(eAtributos.Inteligencia).Caption) * 0.085
    Call SetStars(imgMagiaStar, NumStars * 2)

    ' Estrellas de vida
    NumStars = 0.24 + (Val(lblAtributoFinal(eAtributos.Constitucion).Caption) * 0.5 - ModClase(UserClase).Vida) * 0.475
    Call SetStars(imgVidaStar, NumStars * 2)

    ' Estrellas de escudo
    NumStars = 4 * ModClase(UserClase).Escudo
    Call SetStars(imgEscudosStar, NumStars * 2)

    ' Estrellas de armas
    NumStars = (0.509 + 0.01185 * Val(lblAtributoFinal(eAtributos.Fuerza).Caption)) * ModClase(UserClase).Hit * _
               ModClase(UserClase).DañoArmas + 0.119 * ModClase(UserClase).AtaqueArmas * _
               Val(lblAtributoFinal(eAtributos.Agilidad).Caption)
    Call SetStars(imgArmasStar, NumStars * 2)

    ' Estrellas de arcos
    NumStars = (0.4915 + 0.01265 * Val(lblAtributoFinal(eAtributos.Fuerza).Caption)) * _
               ModClase(UserClase).DañoProyectiles * ModClase(UserClase).Hit + 0.119 * ModClase(UserClase).AtaqueProyectiles * _
               Val(lblAtributoFinal(eAtributos.Agilidad).Caption)
    Call SetStars(imgArcoStar, NumStars * 2)
End Sub

Private Sub SetStars(ByRef ImgContainer As Object, ByVal NumStars As Integer)
    Dim FullStars As Integer
    Dim HasHalfStar As Boolean
    Dim Index As Integer
    Dim Counter As Integer

    If NumStars > 0 Then

        If NumStars > 10 Then NumStars = 10

        FullStars = Int(NumStars / 2)

        ' Tienen brillo extra si estan todas
        If FullStars = 5 Then
            For Index = 1 To FullStars
                ImgContainer(Index).Picture = picGlowStar
            Next Index
        Else
            ' Numero impar? Entonces hay que poner "media estrella"
            If (NumStars Mod 2) > 0 Then HasHalfStar = True

            ' Muestro las estrellas enteras
            If FullStars > 0 Then
                For Index = 1 To FullStars
                    ImgContainer(Index).Picture = picFullStar
                Next Index

                Counter = FullStars
            End If

            ' Muestro la mitad de la estrella (si tiene)
            If HasHalfStar Then
                Counter = Counter + 1

                ImgContainer(Counter).Picture = picHalfStar
            End If

            ' Si estan completos los espacios, no borro nada
            If Counter <> 5 Then
                ' Limpio las que queden vacias
                For Index = Counter + 1 To 5
                    Set ImgContainer(Index).Picture = Nothing
                Next Index
            End If

        End If
    Else
        ' Limpio todo
        For Index = 1 To 5
            Set ImgContainer(Index).Picture = Nothing
        Next Index
    End If

End Sub

Private Sub LoadCharInfo()
    Dim SearchVar As String
    Dim i As Integer

    NroRazas = UBound(ListaRazas())
    NroClases = UBound(ListaClases())

    ReDim ModRaza(1 To NroRazas)
    ReDim ModClase(1 To NroClases)

    'Modificadores de Clase
    For i = 1 To NroClases
        With ModClase(i)
            SearchVar = ListaClases(i)

            .Evasion = Val(GetVar(IniPath & "CharInfo.dat", "MODEVASION", SearchVar))
            .AtaqueArmas = Val(GetVar(IniPath & "CharInfo.dat", "MODATAQUEARMAS", SearchVar))
            .AtaqueProyectiles = Val(GetVar(IniPath & "CharInfo.dat", "MODATAQUEPROYECTILES", SearchVar))
            .DañoArmas = Val(GetVar(IniPath & "CharInfo.dat", "MODDAÑOARMAS", SearchVar))
            .DañoProyectiles = Val(GetVar(IniPath & "CharInfo.dat", "MODDAÑOPROYECTILES", SearchVar))
            .Escudo = Val(GetVar(IniPath & "CharInfo.dat", "MODESCUDO", SearchVar))
            .Hit = Val(GetVar(IniPath & "CharInfo.dat", "HIT", SearchVar))
            .Magia = Val(GetVar(IniPath & "CharInfo.dat", "MODMAGIA", SearchVar))
            .Vida = Val(GetVar(IniPath & "CharInfo.dat", "MODVIDA", SearchVar))
        End With
    Next i

    'Modificadores de Raza
    For i = 1 To NroRazas
        With ModRaza(i)
            SearchVar = Replace(ListaRazas(i), " ", "")

            .Fuerza = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Fuerza"))
            .Agilidad = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Agilidad"))
            .Inteligencia = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Inteligencia"))
            .Carisma = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Carisma"))
            .Constitucion = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Constitucion"))
        End With
    Next i

End Sub
