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
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    

End Sub
Private Sub LoadButtons()
    Set BotonClan = New clsGraphicalButton
    Set BotonFrags = New clsGraphicalButton
    Set BotonOro = New clsGraphicalButton
    Set BotonTorneos = New clsGraphicalButton
    Set BotonRetos = New clsGraphicalButton
    Set BotonNivel = New clsGraphicalButton
    Set LastPressed = New clsGraphicalButton

        
    Call BotonFrags.Initialize(ImgFrags, DirGraficos & "BotonFrags.jpg", _
                                    DirGraficos & "BotonFrags1.jpg", _
                                    DirGraficos & "BotonFrags.jpg", Me)
                                    
    Call BotonClan.Initialize(ImgClan, DirGraficos & "BotonClanes.jpg", _
                                    DirGraficos & "BotonClanes1.jpg", _
                                    DirGraficos & "BotonClanes.jpg", Me)
                                    
    Call BotonOro.Initialize(ImgOro, DirGraficos & "BotonOro.jpg", _
                                    DirGraficos & "BotonOro1.jpg", _
                                    DirGraficos & "BotonOro.jpg", Me)
                                    
    Call BotonRetos.Initialize(ImgReto, DirGraficos & "BotonRetos.jpg", _
                                    DirGraficos & "BotonRetos1.jpg", _
                                    DirGraficos & "BotonRetos.jpg", Me)
                                    
    Call BotonTorneos.Initialize(ImgTorneo, DirGraficos & "Botontorneos.jpg", _
                                    DirGraficos & "BotonTorneos1.jpg", _
                                    DirGraficos & "Botontorneos.jpg", Me)
                                    
    Call BotonNivel.Initialize(ImgNivel, DirGraficos & "BotonNivel.jpg", _
                                    DirGraficos & "BotonNivel1.jpg", _
                                    DirGraficos & "BotonNivel.jpg", Me)
                                    
End Sub

Private Sub Image1_Click()
Unload Me
frmMain.SetFocus
End Sub


Private Sub Image6_Click()

End Sub

Private Sub ImgClan_Click()
Call Audio.PlayWave(SND_CLICK)
FrmRanking2.Picture = LoadPicture(App.path & "\Recursos\CriminalesMatados.jpg")
Call WriteSolicitarRanking(TopClanes)
End Sub


Private Sub ImgFrags_Click()
Call Audio.PlayWave(SND_CLICK)
FrmRanking2.Picture = LoadPicture(App.path & "\Recursos\RankingFrags.jpg")
    Call WriteSolicitarRanking(TopFrags)
End Sub

Private Sub ImgNivel_Click()
Call Audio.PlayWave(SND_CLICK)
FrmRanking2.Picture = LoadPicture(App.path & "\Recursos\CiudadanosMatados.jpg")
Call WriteSolicitarRanking(TopLevel)
End Sub

Private Sub ImgOro_Click()
Call Audio.PlayWave(SND_CLICK)
FrmRanking2.Picture = LoadPicture(App.path & "\Recursos\RankingOro.jpg")
Call WriteSolicitarRanking(TopOro)
End Sub

Private Sub ImgReto_Click()
Call Audio.PlayWave(SND_CLICK)
FrmRanking2.Picture = LoadPicture(App.path & "\Recursos\RankingRetos.jpg")
Call WriteSolicitarRanking(TopRetos)
End Sub

Private Sub ImgTorneo_Click()
Call Audio.PlayWave(SND_CLICK)
FrmRanking2.Picture = LoadPicture(App.path & "\Recursos\RankingTorneos.jpg")
Call WriteSolicitarRanking(TopTorneos)
End Sub
