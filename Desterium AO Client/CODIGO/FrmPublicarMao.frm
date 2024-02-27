VERSION 5.00
Begin VB.Form FrmPublicarMao 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   Picture         =   "FrmPublicarMao.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton OpVenta 
      BackColor       =   &H80000008&
      Caption         =   "Option1"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   3840
      Width           =   255
   End
   Begin VB.OptionButton OpCambio 
      BackColor       =   &H80000008&
      Caption         =   "Option1"
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox txtRecibidor 
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
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   2760
      TabIndex        =   5
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txtValor 
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
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   2180
      TabIndex        =   4
      Top             =   4460
      Width           =   1920
   End
   Begin VB.TextBox txtPin 
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
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtMail 
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
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   2940
      Width           =   2055
   End
   Begin VB.TextBox txtPw 
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
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
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
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   2040
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   4200
      Top             =   120
      Width           =   255
   End
   Begin VB.Image ImgPublicar 
      Height          =   495
      Left            =   1320
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   2640
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "FrmPublicarMao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsFormulario As clsFormMovementManager
Public LastPressed As clsGraphicalButton

Public BotonPublicar As clsGraphicalButton

Private Sub Form_Load()
    
        ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
End Sub

Private Sub Image1_Click()
Unload Me
'FrmMercado.SetFocus
End Sub

Private Sub Image2_Click()
Call Audio.PlayWave(SND_CLICK)
Unload Me
frmMain.SetFocus
End Sub

Private Sub ImgPublicar_Click()
Call Audio.PlayWave(SND_CLICK)
    If OpCambio.value = True Then
        Call WritePacketMercado(PublicarPersonaje, , txtNombre.Text, txtMail.Text, txtPin.Text, txtPw.Text, 0, vbNullString)
    ElseIf OpVenta.value = True Then
        Call WritePacketMercado(PublicarPersonaje, , txtNombre.Text, txtMail.Text, txtPin.Text, txtPw.Text, Val(txtValor.Text), txtRecibidor.Text)
    End If
End Sub

Private Sub OpCambio_Click()
Call Audio.PlayWave(SND_CLICK)
txtValor.Visible = False
txtRecibidor.Visible = False
End Sub

Private Sub OpVenta_Click()
Call Audio.PlayWave(SND_CLICK)
txtValor.Visible = True
txtRecibidor.Visible = True
End Sub

Private Sub txtValor_Change()
If Me.txtValor = "" Then
Exit Sub
End If
If txtValor.Text > 50000000 Then
txtValor.Text = 50000000
ShowConsoleMsg "Sólo se admiten hasta 50.000.000 monedas de oro como valor máximo del personaje.", 65, 190, 156, False, False
End If
End Sub
