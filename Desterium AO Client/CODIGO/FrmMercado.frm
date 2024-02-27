VERSION 5.00
Begin VB.Form FrmMercado 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   Picture         =   "FrmMercado.frx":0000
   ScaleHeight     =   7470
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image6 
      Height          =   495
      Left            =   840
      Top             =   6480
      Width           =   2895
   End
   Begin VB.Image ImgQuitar 
      Height          =   495
      Left            =   720
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Image ImgPublicar 
      Height          =   495
      Left            =   840
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Image ImgOfertasRecibidas 
      Height          =   495
      Left            =   720
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Image ImgOfertasRealizadas 
      Height          =   495
      Left            =   720
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Image ImgMao 
      Height          =   495
      Left            =   720
      Top             =   1080
      Width           =   3015
   End
End
Attribute VB_Name = "FrmMercado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Public LastButtonPressed As clsGraphicalButton

Public BotonMercado As clsGraphicalButton
Public BotonOfertasHechas As clsGraphicalButton
Public BotonRecibidas As clsGraphicalButton
Public BotonPublicar As clsGraphicalButton
Public BotonQuitar As clsGraphicalButton

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_Load()
'    Call LoadButtons

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
End Sub

Private Sub Image6_Click()
    Unload Me
    frmMain.SetFocus

End Sub

Private Sub ImgMAO_Click()
    Call WritePacketMercado(SolicitarLista)
End Sub

Private Sub ImgOfertasRealizadas_Click()
    Call WritePacketMercado(SolicitarListaHechas)
End Sub

Private Sub ImgOfertasRecibidas_Click()
    Call WritePacketMercado(SolicitarListaRecibidas)
End Sub

Private Sub ImgPublicar_Click()
    FrmPublicarMao.Show
    Unload Me
End Sub

Private Sub ImgQuitar_Click()
    Call WritePacketMercado(QuitarVenta)
End Sub

Private Sub LoadButtons()

    Dim GrhPath As String

    GrhPath = DirButtons

    Set BotonMercado = New clsGraphicalButton
    Set BotonOfertasHechas = New clsGraphicalButton
    Set BotonRecibidas = New clsGraphicalButton
    Set BotonPublicar = New clsGraphicalButton
    Set BotonQuitar = New clsGraphicalButton


    Set LastButtonPressed = New clsGraphicalButton

    Call BotonMercado.Initialize(ImgMao, GrhPath & "BotonMAO.jpg", _
                                 GrhPath & "BotonMAO1.jpg", _
                                 GrhPath & "BotonMAO.jpg", Me)

    Call BotonOfertasHechas.Initialize(ImgOfertasRealizadas, GrhPath & "BotonMisOfertas.jpg", _
                                       GrhPath & "BotonMisOfertas1.jpg", _
                                       GrhPath & "BotonMisOfertas.jpg", Me)

    Call BotonRecibidas.Initialize(ImgOfertasRecibidas, GrhPath & "BotonVerOfertas.jpg", _
                                   GrhPath & "BotonVerOfertas1.jpg", _
                                   GrhPath & "BotonVerOfertas.jpg", Me)

    Call BotonPublicar.Initialize(ImgPublicar, GrhPath & "BotonPublicar.jpg", _
                                  GrhPath & "BotonPublicar1.jpg", _
                                  GrhPath & "BotonPublicar.jpg", Me)

    Call BotonQuitar.Initialize(ImgQuitar, GrhPath & "BotonQuitar.jpg", _
                                GrhPath & "BotonQuitar1.jpg", _
                                GrhPath & "BotonQuitar.jpg", Me)
End Sub

