VERSION 5.00
Begin VB.Form FrmRanking 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   Picture         =   "FrmRanking.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2640
      Top             =   120
      Width           =   255
   End
   Begin VB.Image ImgOro 
      Height          =   495
      Left            =   360
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Image ImgFrags 
      Height          =   495
      Left            =   360
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Image ImgReto 
      Height          =   495
      Left            =   360
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Image ImgNivel 
      Height          =   495
      Left            =   360
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Image ImgTorneo 
      Height          =   495
      Left            =   360
      Top             =   480
      Width           =   2295
   End
   Begin VB.Image ImgClan 
      Height          =   495
      Left            =   360
      Top             =   2640
      Width           =   2295
   End
End
Attribute VB_Name = "FrmRanking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Public LastPressed As clsGraphicalButton

' @ Botones
Public BotonClan As clsGraphicalButton
Public BotonFrags As clsGraphicalButton
Public BotonOro As clsGraphicalButton
Public BotonTorneos As clsGraphicalButton
Public BotonRetos As clsGraphicalButton
Public BotonNivel As clsGraphicalButton

Private Sub Form_Load()

          ' Handles Form movement (drag and drop).
10        Set clsFormulario = New clsFormMovementManager
20        clsFormulario.Initialize Me
          

End Sub
Private Sub LoadButtons()
10        Set BotonClan = New clsGraphicalButton
20        Set BotonFrags = New clsGraphicalButton
30        Set BotonOro = New clsGraphicalButton
40        Set BotonTorneos = New clsGraphicalButton
50        Set BotonRetos = New clsGraphicalButton
60        Set BotonNivel = New clsGraphicalButton
70        Set LastPressed = New clsGraphicalButton

              
80        Call BotonFrags.Initialize(ImgFrags, DirGraficos & "BotonFrags.jpg", _
              DirGraficos & "BotonFrags1.jpg", DirGraficos & "BotonFrags.jpg", Me)
                                          
90        Call BotonClan.Initialize(ImgClan, DirGraficos & "BotonClanes.jpg", _
              DirGraficos & "BotonClanes1.jpg", DirGraficos & "BotonClanes.jpg", Me)
                                          
100       Call BotonOro.Initialize(ImgOro, DirGraficos & "BotonOro.jpg", DirGraficos _
              & "BotonOro1.jpg", DirGraficos & "BotonOro.jpg", Me)
                                          
110       Call BotonRetos.Initialize(ImgReto, DirGraficos & "BotonRetos.jpg", _
              DirGraficos & "BotonRetos1.jpg", DirGraficos & "BotonRetos.jpg", Me)
                                          
120       Call BotonTorneos.Initialize(ImgTorneo, DirGraficos & "Botontorneos.jpg", _
              DirGraficos & "BotonTorneos1.jpg", DirGraficos & "Botontorneos.jpg", Me)
                                          
130       Call BotonNivel.Initialize(ImgNivel, DirGraficos & "BotonNivel.jpg", _
              DirGraficos & "BotonNivel1.jpg", DirGraficos & "BotonNivel.jpg", Me)
                                          
End Sub

Private Sub Image1_Click()
10    Unload Me
20    frmMain.SetFocus
End Sub


Private Sub Image6_Click()

End Sub

Private Sub ImgClan_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    FrmRanking2.Picture = LoadPicture(App.path & "\Recursos\CriminalesMatados.jpg")
30    Call WriteSolicitarRanking(TopClanes)
End Sub


Private Sub ImgFrags_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    FrmRanking2.Picture = LoadPicture(App.path & "\Recursos\RankingFrags.jpg")
30        Call WriteSolicitarRanking(TopFrags)
End Sub

Private Sub ImgNivel_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    FrmRanking2.Picture = LoadPicture(App.path & "\Recursos\CiudadanosMatados.jpg")
30    Call WriteSolicitarRanking(TopLevel)
End Sub

Private Sub ImgOro_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    FrmRanking2.Picture = LoadPicture(App.path & "\Recursos\RankingOro.jpg")
30    Call WriteSolicitarRanking(TopOro)
End Sub

Private Sub ImgReto_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    FrmRanking2.Picture = LoadPicture(App.path & "\Recursos\RankingRetos.jpg")
30    Call WriteSolicitarRanking(TopRetos)
End Sub

Private Sub ImgTorneo_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    FrmRanking2.Picture = LoadPicture(App.path & "\Recursos\RankingTorneos.jpg")
30    Call WriteSolicitarRanking(TopTorneos)
End Sub
