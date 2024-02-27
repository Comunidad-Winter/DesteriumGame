VERSION 5.00
Begin VB.Form frmPanelSeguimiento 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "              Panel"
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   1680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Dejar de seguir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblVida 
      BackStyle       =   0  'Transparent
      Caption         =   "Vida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblMana 
      BackStyle       =   0  'Transparent
      Caption         =   "Mana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image ImgMana 
      Height          =   165
      Left            =   120
      Picture         =   "frmPanelSeguimiento.frx":0000
      Top             =   720
      Width           =   1410
   End
   Begin VB.Image ImgVida 
      Height          =   165
      Left            =   120
      Picture         =   "frmPanelSeguimiento.frx":060C
      Top             =   240
      Width           =   1410
   End
   Begin VB.Label Slot 
      BackStyle       =   0  'Transparent
      Caption         =   "Slot: Sin seleccionar"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label InvOrSpell 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "frmPanelSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Public LastPressed As clsGraphicalButton

Private Sub Command1_Click()
    Call WriteSeguimiento("1")
    Unload Me
End Sub

Private Sub Form_Load()

' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Set LastPressed = New clsGraphicalButton

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastPressed.ToggleToNormal
End Sub


