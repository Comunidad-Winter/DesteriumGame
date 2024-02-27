VERSION 5.00
Begin VB.Form FrmInfoMaoPjs 
   BorderStyle     =   0  'None
   Caption         =   "Informacion de personaje"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmInfoMaoPjs.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstSkill 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000005&
      Height          =   3345
      Left            =   8280
      TabIndex        =   15
      Top             =   1800
      Width           =   2895
   End
   Begin VB.ListBox lstSpell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000005&
      Height          =   2370
      Left            =   6480
      TabIndex        =   14
      Top             =   5280
      Width           =   2055
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   3720
      ScaleHeight     =   162.727
      ScaleMode       =   0  'User
      ScaleWidth      =   163.012
      TabIndex        =   13
      Top             =   5280
      Width           =   2415
   End
   Begin VB.PictureBox PicBancoInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   4080
      ScaleHeight     =   2400
      ScaleWidth      =   3870
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3870
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   4800
      Top             =   8520
      Width           =   2535
   End
   Begin VB.Label lblAT 
      BackStyle       =   0  'Transparent
      Caption         =   "30000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   22
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lblAT 
      BackStyle       =   0  'Transparent
      Caption         =   "30000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   4
      Left            =   1680
      TabIndex        =   21
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label lblAT 
      BackStyle       =   0  'Transparent
      Caption         =   "30000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   20
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblAT 
      BackStyle       =   0  'Transparent
      Caption         =   "30000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   19
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblAT 
      BackStyle       =   0  'Transparent
      Caption         =   "30000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   18
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblBandido 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label lblAsesino 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   5610
      Width           =   2055
   End
   Begin VB.Label lblUserPremium 
      BackStyle       =   0  'Transparent
      Caption         =   "SI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2685
      TabIndex        =   11
      Top             =   7380
      Width           =   375
   End
   Begin VB.Label lblUserOro 
      BackStyle       =   0  'Transparent
      Caption         =   "SI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   7120
      Width           =   375
   End
   Begin VB.Label lblFundador 
      BackStyle       =   0  'Transparent
      Caption         =   "SI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudadano"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label lblGld 
      BackStyle       =   0  'Transparent
      Caption         =   "3000000000000000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   5350
      Width           =   2055
   End
   Begin VB.Label lblMana 
      BackStyle       =   0  'Transparent
      Caption         =   "30000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   4880
      Width           =   615
   End
   Begin VB.Label lblVida 
      BackStyle       =   0  'Transparent
      Caption         =   "30000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1350
      TabIndex        =   5
      Top             =   4605
      Width           =   615
   End
   Begin VB.Label lblFamas 
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1485
      TabIndex        =   4
      Top             =   2595
      Width           =   495
   End
   Begin VB.Label lblraza 
      BackStyle       =   0  'Transparent
      Caption         =   "Elfo Oscuro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   2340
      Width           =   2295
   End
   Begin VB.Label lblclase 
      BackStyle       =   0  'Transparent
      Caption         =   "Clerigo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   2080
      Width           =   2295
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      Caption         =   "47 (50%)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   1850
      Width           =   2295
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "LAUTARO MAO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   1440
      Width           =   2295
   End
End
Attribute VB_Name = "FrmInfoMaoPjs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
FrmInfoMao.Show
Unload Me
End Sub

