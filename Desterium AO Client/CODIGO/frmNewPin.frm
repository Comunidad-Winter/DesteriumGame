VERSION 5.00
Begin VB.Form frmNewpin 
   BorderStyle     =   0  'None
   Caption         =   "Cambiar Contraseña"
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewPin.frx":0000
   ScaleHeight     =   220
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1920
      Width           =   2760
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1380
      Width           =   2760
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   855
      Width           =   2760
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4560
      Top             =   0
      Width           =   375
   End
   Begin VB.Label ImgCancelar 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Image imgAceptar 
      Height          =   495
      Left            =   1800
      Tag             =   "1"
      Top             =   2640
      Width           =   1575
   End
End
Attribute VB_Name = "frmNewpin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonAceptar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
  '  Me.Picture = LoadPicture(App.path & "\Recursos\VentanaCambiarcontrasenia.jpg")
    
    
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub imgAceptar_Click()

Call Audio.PlayWave(SND_CLICK)

If Len(Text2.Text) > 0 Then Unload Me
    If Text2.Text <> Text3.Text Then
        Call MsgBox("Las claves no coinciden", vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Cambiar Contraseña")
        Exit Sub
    End If
    
    Call WriteChangePin(Text1.Text, Text2.Text)
    
    Unload Me
End Sub

Private Sub imgCancelar_Click()
Call Audio.PlayWave(SND_CLICK)
Unload Me
End Sub
