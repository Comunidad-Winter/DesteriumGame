VERSION 5.00
Begin VB.Form frmrecuperar 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   Picture         =   "frmrecuperar.frx":0000
   ScaleHeight     =   3300
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPIN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox DATO2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1400
      Width           =   2655
   End
   Begin VB.TextBox DATO1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   870
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   4440
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   1680
      Top             =   2760
      Width           =   1815
   End
End
Attribute VB_Name = "frmrecuperar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If EstadoLogin = E_MODO.RecuperarPJ Then
Me.Caption = "Recuperar Pj"
lbl2.Caption = "E-Mail"
Command1.Caption = "Recuperar"
Modo = 2
End If
End Sub

Private Sub Image1_Click()
If Modo = 2 Then
        UserName = DATO1
        UserEmail = DATO2
        UserPin = txtPIN
Else


'Si por X razon llegamos aca y no se asigno el modo
MsgBox "Ocurrio un error En el proceso. Reintentelo."
        Unload Me
        Exit Sub
        End If
Call Login

Unload Me
End Sub

Private Sub Image2_Click()
Unload Me
End Sub
