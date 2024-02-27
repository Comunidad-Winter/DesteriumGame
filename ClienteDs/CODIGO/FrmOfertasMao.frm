VERSION 5.00
Begin VB.Form FrmOfertasMao 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmOfertasMao.frx":0000
   ScaleHeight     =   4515
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstPjs 
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
      Height          =   2790
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   255
   End
   Begin VB.Image ImgRechazar 
      Height          =   375
      Left            =   480
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Image ImgAceptar 
      Height          =   495
      Left            =   480
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   2640
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "FrmOfertasMao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsFormulario As clsFormMovementManager
Public LastPressed As clsGraphicalButton

Public BotonAceptar As clsGraphicalButton
Public BotonRechazar As clsGraphicalButton


Private Sub Form_Load()
    
    
        ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
End Sub




Private Sub Image1_Click()
Unload Me
FrmMercado.SetFocus
End Sub

Private Sub Image2_Click()
Call Audio.PlayWave(SND_CLICK)
    Dim Nick As String

    Nick = lstPjs.Text
    
    If LenB(Nick) <> 0 Then _
        Call WriteRequestCharInfoMercado(Nick)
End Sub

Private Sub imgAceptar_Click()
Call Audio.PlayWave(SND_CLICK)
If MsgBox("¿Seguro que desea aceptar esa oferta de intercambio? Tu contraseña/pin/email pasarán a ser los datos del personaje recibido.", vbYesNo + vbQuestion, "Desterium AO") = vbYes Then
    Call WritePacketMercado(AceptarOferta, lstPjs.ListIndex + 1)
Else
            Exit Sub
        End If
End Sub

Private Sub imgRechazar_Click()
If MsgBox("¿Deseas rechazar la oferta de cambio?", vbYesNo + vbQuestion, "Desterium AO") = vbYes Then
    Call WritePacketMercado(RechazarOferta, lstPjs.ListIndex + 1)
    Else
            Exit Sub
        End If
End Sub
