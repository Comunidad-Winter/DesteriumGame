VERSION 5.00
Begin VB.Form FrmRECBORR 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   ClientHeight    =   3285
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   4950
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrnRECBORR.frx":0000
   ScaleHeight     =   3285
   ScaleWidth      =   4950
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
      Height          =   375
      Left            =   4560
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   1560
      Top             =   2760
      Width           =   1815
   End
End
Attribute VB_Name = "FrmRECBORR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Modo As Byte

Private Sub Image1_Click()

10    If Modo = 1 Then
20    Me.Picture = LoadPicture(App.path & "\Recursos\VentanaBorrar.jpg")
30            UserName = DATO1
40            UserPassword = DATO2
50            UserPin = txtPin
60    ElseIf Modo = 2 Then
70    Me.Picture = LoadPicture(App.path & "\Recursos\VentanaRecuperar.jpg")
80            UserName = DATO1
90            UserEmail = DATO2
100           UserPin = txtPin
110   Else


      'Si por X razon llegamos aca y no se asigno el modo
120   MsgBox "Ocurrio un error En el proceso. Reintentelo."
130           Unload Me
140           Exit Sub
150           End If
160   Call Login

170   Unload Me
End Sub

Private Sub Form_Load()
10    If EstadoLogin = E_MODO.BorrarPJ Then
20    Me.Picture = LoadPicture(App.path & "\Recursos\VentanaBorrar.jpg")
30    Modo = 1
40    ElseIf EstadoLogin = E_MODO.RecuperarPJ Then
50    Me.Picture = LoadPicture(App.path & "\Recursos\VentanaRecuperar.jpg")
60    Modo = 2
70    End If
End Sub


Private Sub Image2_Click()
10    Unload Me
End Sub
