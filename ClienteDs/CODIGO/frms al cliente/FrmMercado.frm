VERSION 5.00
Begin VB.Form FrmMercado 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   Picture         =   "FrmMercado.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image ImgMao 
      Height          =   615
      Left            =   720
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Image ImgOfertasRealizadas 
      Height          =   495
      Left            =   720
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Image ImgOfertasRecibidas 
      Height          =   495
      Left            =   720
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Image ImgPublicar 
      Height          =   495
      Left            =   840
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Image ImgQuitar 
      Height          =   495
      Left            =   720
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   840
      Top             =   6480
      Width           =   2895
   End
End
Attribute VB_Name = "FrmMercado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsFormulario As clsFormMovementManager
Public LastPressed As clsGraphicalButton

Public BotonMercado As clsGraphicalButton
Public BotonOfertasHechas As clsGraphicalButton
Public BotonRecibidas As clsGraphicalButton
Public BotonPublicar As clsGraphicalButton
Public BotonQuitar As clsGraphicalButton
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
10        LastPressed.ToggleToNormal
End Sub

Private Sub Form_Load()
10        Call LoadButtons
          
              ' Handles Form movement (drag and drop).
20        Set clsFormulario = New clsFormMovementManager
30        clsFormulario.Initialize Me
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image6_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    Unload Me
30    frmMain.SetFocus

End Sub

Private Sub ImgMAO_Click()
10        Call Audio.PlayWave(SND_CLICK)
          
20        Call WriteRequestMercado
          
End Sub

Private Sub ImgOfertasRealizadas_Click()
10        Call Audio.PlayWave(SND_CLICK)
20        WriteRequestOfferSentUser
End Sub

Private Sub ImgOfertasRecibidas_Click()
10        Call Audio.PlayWave(SND_CLICK)
20        WriteRequestOfferUser
          
End Sub

Private Sub ImgPersonajesPublicados_Click()

End Sub

Private Sub ImgPublicar_Click()
10        Call Audio.PlayWave(SND_CLICK)
20        FrmPublicarMao.Show , frmMain
30        Unload Me
End Sub

Private Sub imgQuitar_Click()
10        Call Audio.PlayWave(SND_CLICK)
20        Call WriteQuitarPj
End Sub
Private Sub LoadButtons()
          Dim GrhPath As String
          
10        GrhPath = DirGraficos
          
20        Set BotonMercado = New clsGraphicalButton
30        Set BotonOfertasHechas = New clsGraphicalButton
40        Set BotonRecibidas = New clsGraphicalButton
50        Set BotonPublicar = New clsGraphicalButton
60        Set BotonQuitar = New clsGraphicalButton
          
          
70        Set LastPressed = New clsGraphicalButton
          
         ' Call BotonMercado.Initialize(ImgMAO, grhpath & "BotonMAO.jpg", grhpath & "BotonMAO1.jpg", grhpath & "BotonMAO.jpg", Me)
          
        '  Call BotonOfertasHechas.Initialize(ImgOfertasRealizadas, grhpath & "BotonMisOfertas.jpg", grhpath & "BotonMisOfertas1.jpg", grhpath & "BotonMisOfertas.jpg", Me)
                                         
         ' Call BotonRecibidas.Initialize(ImgOfertasRecibidas, grhpath & "BotonVerOfertas.jpg", grhpath & "BotonVerOfertas1.jpg", grhpath & "BotonVerOfertas.jpg", Me)
                                          
          'Call BotonPublicar.Initialize(ImgPublicar, grhpath & "BotonPublicar.jpg", grhpath & "BotonPublicar1.jpg", grhpath & "BotonPublicar.jpg", Me)
                                          
         ' Call BotonQuitar.Initialize(ImgQuitar, grhpath & "BotonQuitar.jpg", grhpath & "BotonQuitar1.jpg", grhpath & "BotonQuitar.jpg", Me)
End Sub

