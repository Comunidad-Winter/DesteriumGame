VERSION 5.00
Begin VB.Form FrmOfertasMao2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   Picture         =   "FrmOfertasMao2.frx":0000
   ScaleHeight     =   4500
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
      Height          =   2985
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   2640
      Top             =   120
      Width           =   375
   End
   Begin VB.Image OfertasMAO2 
      Height          =   375
      Left            =   480
      Top             =   3960
      Width           =   2175
   End
End
Attribute VB_Name = "FrmOfertasMao2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Public LastButtonPressed As clsGraphicalButton

Public BotonCancelar As clsGraphicalButton

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_Load()
'    Call LoadButtons

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
End Sub

Private Sub LoadButtons()

    Set BotonCancelar = New clsGraphicalButton

    Set LastButtonPressed = New clsGraphicalButton

    Call BotonCancelar.Initialize(OfertasMAO2, DirButtons & "BotonCancelarMAO.jpg", _
                                  DirButtons & "BotonCancelarMAO1.jpg", _
                                  DirButtons & "BotonCancelarMAO.jpg", Me)
End Sub

Private Sub OfertasMAO2_Click()
    If lstPjs.ListIndex < 0 Then Exit Sub
    Call WritePacketMercado(EliminarOferta, lstPjs.ListIndex + 1)
End Sub

Private Sub Image1_Click()
    Unload Me
    FrmMercado.SetFocus
End Sub
