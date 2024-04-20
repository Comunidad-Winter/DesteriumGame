VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "El programa que no tenia ganas de hacer yo"
   ClientHeight    =   1065
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1065
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbInd 
      Caption         =   "Comprimir Personajes.ind"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2730
      TabIndex        =   1
      Top             =   210
      Width           =   2355
   End
   Begin VB.CommandButton cmbINI 
      Caption         =   "Extraer Personajes.ind"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   2145
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   315
      TabIndex        =   2
      Top             =   735
      Width           =   4665
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbInd_Click()
    
    Call mIni.Ini_Load_Body
    Call Write_Body
End Sub

Private Sub cmbINI_Click()
    Call Read_Body
    Call Ini_Generate_Body
End Sub

Private Sub Form_Load()
    Call Load_Main
End Sub

