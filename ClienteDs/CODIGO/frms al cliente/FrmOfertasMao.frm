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
      Top             =   120
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
10        Set clsFormulario = New clsFormMovementManager
20        clsFormulario.Initialize Me
End Sub




Private Sub Image1_Click()
10    Unload Me
20    FrmMercado.SetFocus
End Sub

Private Sub Image2_Click()
10        Call Audio.PlayWave(SND_CLICK)
              
20        If lstPjs.ListIndex = -1 Then
30            MsgBox "Selecciona un personaje de la lista"
40            Exit Sub
50        End If
          
60        Call WriteRequestInfoCharMAO(lstPjs.List(lstPjs.ListIndex))
         ' If LenB(Nick) <> 0 Then Call WriteRequestCharInfoMercado(Nick)
End Sub

Private Sub imgAceptar_Click()
          Dim PIN As String
          
10        If lstPjs.ListIndex = -1 Then
20            MsgBox "Selecciona una invitación"
30            Exit Sub
40        End If
          
50        Call Audio.PlayWave(SND_CLICK)
          
60        If _
              MsgBox("¿Seguro que desea aceptar esa oferta de intercambio? Tu contraseña/pin/email pasarán a ser los datos del personaje recibido.", _
              vbYesNo + vbQuestion, "Desterium AO") = vbYes Then
70            PIN = _
                  InputBox("Escriba el pin de su personaje asi podrá aceptar la invitación")
80            Call WriteMercadoAcceptInvitation(lstPjs.List(lstPjs.ListIndex), PIN)
90        End If
          
100       Unload Me
End Sub

Private Sub imgRechazar_Click()
10        If lstPjs.ListIndex = -1 Then
20            MsgBox "Selecciona una invitación"
30            Exit Sub
40        End If
          
50        Call Audio.PlayWave(SND_CLICK)
          
60        If MsgBox("¿Deseas rechazar la oferta de cambio?", vbYesNo + vbQuestion, _
              "Desterium AO") = vbYes Then
70            Call WriteMercadoRechaceInvitation(lstPjs.List(lstPjs.ListIndex))
80        End If
          
90        Unload Me
End Sub
