VERSION 5.00
Begin VB.Form frmPanelSeguimiento 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Miralo al chitero sucio"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   2070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Dejar de Seguir"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Image ImgVida 
      Height          =   165
      Left            =   240
      Picture         =   "frmPanelSeguimiento.frx":0000
      Top             =   360
      Width           =   1410
   End
   Begin VB.Image ImgMana 
      Height          =   165
      Left            =   240
      Picture         =   "frmPanelSeguimiento.frx":05A0
      Top             =   840
      Width           =   1410
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
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1215
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPanelSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call WriteSeguimiento("1")
Unload Me
End Sub

