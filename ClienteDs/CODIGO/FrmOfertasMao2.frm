VERSION 5.00
Begin VB.Form FrmOfertasMao2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "FrmOfertasMao2.frx":0000
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
      Height          =   2985
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.Image OfertasMAO2 
      Height          =   375
      Left            =   480
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   2640
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "FrmOfertasMao2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsFormulario As clsFormMovementManager
Public LastPressed As clsGraphicalButton

Public BotonCancelar As clsGraphicalButton

Private Sub Form_Load()
    
        ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
End Sub

Private Sub OfertasMAO2_Click()
    Call WritePacketMercado(EliminarOferta, lstPjs.ListIndex + 1)
End Sub




Private Sub Image1_Click()
Unload Me
FrmMercado.SetFocus
End Sub


Private Sub ImgOfertasMao2_Click()

End Sub
