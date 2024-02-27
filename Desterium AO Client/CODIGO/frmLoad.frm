VERSION 5.00
Begin VB.Form frmLoad 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   4905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   Picture         =   "frmLoad.frx":0000
   ScaleHeight     =   4905
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgUP 
      Height          =   465
      Left            =   644
      Picture         =   "frmLoad.frx":3ECC2
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Image imgforo 
      Height          =   450
      Left            =   645
      Picture         =   "frmLoad.frx":3F9DF
      Top             =   3055
      Width           =   2655
   End
   Begin VB.Image imgActualizar 
      Height          =   465
      Left            =   644
      Picture         =   "frmLoad.frx":407B5
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Image imgCancelar 
      Height          =   480
      Left            =   635
      Picture         =   "frmLoad.frx":41558
      Tag             =   "1"
      Top             =   4155
      Width           =   2670
   End
   Begin VB.Image imgSiguiente 
      Height          =   450
      Left            =   644
      Picture         =   "frmLoad.frx":42094
      Tag             =   "1"
      Top             =   1944
      Width           =   2640
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private clsFormulario As clsFormMovementManager

Private cBotonSiguiente As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton
Private cBotonActualizar As clsGraphicalButton
Private cBotonForo As clsGraphicalButton
Private cBotonup As clsGraphicalButton
Public LastPressed As clsGraphicalButton

Private Sub Form_Deactivate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Call LoadButtons
    

End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    
    GrhPath = DirGraficos

    Set cBotonSiguiente = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
    Set cBotonActualizar = New clsGraphicalButton
    Set cBotonForo = New clsGraphicalButton
Set cBotonup = New clsGraphicalButton
    Set LastPressed = New clsGraphicalButton
    

    
    Call cBotonSiguiente.Initialize(imgSiguiente, GrhPath & "imgBotonSiguiente.jpg", _
                                    GrhPath & "imgBotonSiguiente2.jpg", _
                                    GrhPath & "imgBotonSiguiente.jpg", Me)

    Call cBotonCancelar.Initialize(imgCancelar, GrhPath & "imgBotonCancelar.jpg", _
                                    GrhPath & "imgBotonCancelar2.jpg", _
                                    GrhPath & "imgBotonCancelar.jpg", Me)
      
    Call cBotonActualizar.Initialize(imgActualizar, GrhPath & "imgBotonActualizar.jpg", _
                                    GrhPath & "imgBotonActualizar2.jpg", _
                                    GrhPath & "imgBotonActualizar.jpg", Me)
                                    
    Call cBotonForo.Initialize(imgforo, GrhPath & "imgForo.jpg", _
                                    GrhPath & "imgForo2.jpg", _
                                    GrhPath & "imgForo.jpg", Me)
 Call cBotonup.Initialize(imgUP, GrhPath & "imgBotonUp.jpg", _
                                    GrhPath & "imgBotonUp2.jpg", _
                                    GrhPath & "imgBotonUp.jpg", Me)

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub imgCancelar_Click()
    Unload Me
End Sub

Private Sub imgSiguiente_Click()
Call Main
Unload Me
End Sub
