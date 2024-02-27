VERSION 5.00
Begin VB.Form frmCharInfo 
   BorderStyle     =   0  'None
   Caption         =   "Información del personaje"
   ClientHeight    =   6585
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6390
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCharInfo.frx":0000
   ScaleHeight     =   439
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   426
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMiembro 
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
      Height          =   1230
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   4740
      Width           =   5940
   End
   Begin VB.TextBox txtPeticiones 
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
      Height          =   1230
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   3120
      Width           =   5940
   End
   Begin VB.Image imgAceptar 
      Height          =   390
      Left            =   5040
      Tag             =   "1"
      Top             =   6075
      Width           =   1140
   End
   Begin VB.Image imgRechazar 
      Height          =   390
      Left            =   3840
      Tag             =   "1"
      Top             =   6075
      Width           =   1020
   End
   Begin VB.Image imgPeticion 
      Height          =   510
      Left            =   2760
      Tag             =   "1"
      Top             =   6000
      Width           =   1020
   End
   Begin VB.Image imgEchar 
      Height          =   510
      Left            =   1560
      Tag             =   "1"
      Top             =   5955
      Width           =   1020
   End
   Begin VB.Image imgCerrar 
      Height          =   510
      Left            =   240
      Tag             =   "1"
      Top             =   6000
      Width           =   1020
   End
   Begin VB.Label status 
      BackStyle       =   0  'Transparent
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
      Left            =   3120
      TabIndex        =   12
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label Nombre 
      BackStyle       =   0  'Transparent
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
      Left            =   1320
      TabIndex        =   11
      Top             =   700
      Width           =   1440
   End
   Begin VB.Label Nivel 
      BackStyle       =   0  'Transparent
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
      Left            =   1320
      TabIndex        =   10
      Top             =   1750
      Width           =   1185
   End
   Begin VB.Label Clase 
      BackStyle       =   0  'Transparent
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
      Left            =   1320
      TabIndex        =   9
      Top             =   1225
      Width           =   1575
   End
   Begin VB.Label Raza 
      BackStyle       =   0  'Transparent
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
      Left            =   1320
      TabIndex        =   8
      Top             =   960
      Width           =   1560
   End
   Begin VB.Label Genero 
      BackStyle       =   0  'Transparent
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
      Left            =   1320
      TabIndex        =   7
      Top             =   1500
      Width           =   1335
   End
   Begin VB.Label Oro 
      BackStyle       =   0  'Transparent
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
      Left            =   1320
      TabIndex        =   6
      Top             =   2010
      Width           =   1365
   End
   Begin VB.Label Banco 
      BackStyle       =   0  'Transparent
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
      Left            =   1320
      TabIndex        =   5
      Top             =   2250
      Width           =   1425
   End
   Begin VB.Label guildactual 
      BackStyle       =   0  'Transparent
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
      Left            =   3960
      TabIndex        =   4
      Top             =   960
      Width           =   2265
   End
   Begin VB.Label ejercito 
      BackStyle       =   0  'Transparent
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
      Left            =   3960
      TabIndex        =   3
      Top             =   1230
      Width           =   1785
   End
   Begin VB.Label Ciudadanos 
      BackStyle       =   0  'Transparent
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
      Left            =   4905
      TabIndex        =   2
      Top             =   1500
      Width           =   1185
   End
   Begin VB.Label criminales 
      BackStyle       =   0  'Transparent
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
      Left            =   4920
      TabIndex        =   1
      Top             =   1770
      Width           =   1185
   End
   Begin VB.Label reputacion 
      BackStyle       =   0  'Transparent
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
      Left            =   4905
      TabIndex        =   0
      Top             =   2040
      Width           =   1185
   End
End
Attribute VB_Name = "frmCharInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Developed AO 13.0
'FrmCharInfo

Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonCerrar As clsGraphicalButton
Private cBotonPeticion As clsGraphicalButton
Private cBotonRechazar As clsGraphicalButton
Private cBotonEchar As clsGraphicalButton
Private cBotonAceptar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Public Enum CharInfoFrmType
    frmMembers
    frmMembershipRequests
End Enum

Public frmType As CharInfoFrmType

Private Sub Form_Load()
' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

  '  Me.Picture = LoadPicture(App.path & "\Graficos\VentanaInfoPj.jpg")

   ' Call LoadButtons

End Sub

Private Sub LoadButtons()
    Dim GrhPath As String

    GrhPath = DirGraficos

    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonPeticion = New clsGraphicalButton
    Set cBotonRechazar = New clsGraphicalButton
    Set cBotonEchar = New clsGraphicalButton
    Set cBotonAceptar = New clsGraphicalButton

    Set LastPressed = New clsGraphicalButton


    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrarInfoChar.jpg", _
                                 GrhPath & "BotonCerrarClickInfoChar.jpg", _
                                 GrhPath & "BotonCerrarInfoChar.jpg", Me)

    Call cBotonPeticion.Initialize(imgPeticion, GrhPath & "BotonPeticion.jpg", _
                                   GrhPath & "BotonPeticionApretado.jpg", _
                                   GrhPath & "BotonPeticion.jpg", Me)

    Call cBotonRechazar.Initialize(ImgRechazar, GrhPath & "BotonRechazar.jpg", _
                                   GrhPath & "BotonRechazarClick.jpg", _
                                   GrhPath & "BotonRechazar.jpg", Me)

    Call cBotonEchar.Initialize(imgEchar, GrhPath & "BotonEchar.jpg", _
                                GrhPath & "BotonEcharclick.jpg", _
                                GrhPath & "BotonEchar.jpg", Me)

    Call cBotonAceptar.Initialize(ImgAceptar, GrhPath & "BotonAceptarInfoChar.jpg", _
                                  GrhPath & "BotonAceptarClickInfoChar.jpg", _
                                  GrhPath & "BotonAceptarInfoChar.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 '   LastPressed.ToggleToNormal
End Sub

Private Sub imgAceptar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteGuildAcceptNewMember(Nombre)
    Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
    Unload Me
End Sub

Private Sub imgCerrar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Unload Me
End Sub

Private Sub imgEchar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteGuildKickMember(Nombre)
    Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
    Unload Me
End Sub

Private Sub imgPeticion_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteGuildRequestJoinerInfo(Nombre)
End Sub

Private Sub imgRechazar_Click()
    Call Audio.PlayWave(SND_CLICK)
    frmCommet.t = RECHAZOPJ
    frmCommet.Nombre = Nombre.Caption
    frmCommet.Show vbModeless, frmCharInfo
End Sub

Private Sub txtMiembro_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    LastPressed.ToggleToNormal
End Sub
