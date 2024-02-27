VERSION 5.00
Begin VB.Form frmMessageTxt 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Mensajes Predefinidos"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmMessageTxt.frx":0000
   ScaleHeight     =   4695
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   1320
      TabIndex        =   9
      Top             =   3720
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   8
      Left            =   1320
      TabIndex        =   8
      Top             =   3360
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   1320
      TabIndex        =   7
      Top             =   3000
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   1320
      TabIndex        =   6
      Top             =   2630
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      Top             =   2250
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   4
      Top             =   1870
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Text            =   " "
      Top             =   1490
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   1120
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   780
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   420
      Width           =   3165
   End
   Begin VB.Image ImgCancelar 
      Height          =   375
      Left            =   120
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Image ImgGuardar 
      Height          =   375
      Left            =   3240
      Top             =   4185
      Width           =   1335
   End
End
Attribute VB_Name = "frmMessageTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonGuardar As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Form_Load()
          Dim i As Long
          
          ' Handles Form movement (drag and drop).
10        Set clsFormulario = New clsFormMovementManager
20        clsFormulario.Initialize Me
          
30        For i = 0 To 9
40            messageTxt(i) = CustomMessages.Message(i)
50        Next i

        '  Me.Picture = LoadPicture(App.path & "\graficos\VentanaMensajesPersonalizados.jpg")
          
60        LoadButtons
          
End Sub

Private Sub LoadButtons()
          Dim GrhPath As String
          
10        GrhPath = DirGraficos
          
20        Set cBotonGuardar = New clsGraphicalButton
30        Set cBotonCancelar = New clsGraphicalButton
          
40        Set LastPressed = New clsGraphicalButton

         ' Call cBotonGuardar.Initialize(imgGuardar, GrhPath & "BotonGuardarCustomMsg.jpg", GrhPath & "BotonGuardarRolloverCustomMsg.jpg", GrhPath & "BotonGuardarClickCustomMsg.jpg", Me)
         ' Call cBotonCancelar.Initialize(ImgCancelar, GrhPath & "BotonCancelarCustomMsg.jpg", GrhPath & "BotonCancelarRolloverCustomMsg.jpg", GrhPath & "BotonCancelarClickCustomMsg.jpg", Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub imgCancelar_Click()
10        Unload Me
End Sub

Private Sub imgGuardar_Click()
10    On Error GoTo ErrHandler
          Dim i As Long
          
20        For i = 0 To 9
30            CustomMessages.Message(i) = messageTxt(i)
40        Next i
          
50        Unload Me
60    Exit Sub

ErrHandler:
          'Did detected an invalid message??
70        If Err.number = CustomMessages.InvalidMessageErrCode Then
80            Call MsgBox("El Mensaje " & CStr(i + 1) & _
                  " es inválido. Modifiquelo por favor.")
90        End If

End Sub

Private Sub messageTxt_MouseMove(Index As Integer, Button As Integer, Shift As _
    Integer, X As Single, Y As Single)
    'LastPressed.ToggleToNormal
End Sub
