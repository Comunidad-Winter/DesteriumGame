VERSION 5.00
Begin VB.Form frmConnect2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   2535
   ClientTop       =   3060
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "frmConnect2.frx":0000
   ScaleHeight     =   9000
   ScaleMode       =   0  'User
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNombre 
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
      Height          =   225
      Left            =   5040
      TabIndex        =   2
      Top             =   3000
      Width           =   3165
   End
   Begin VB.TextBox txtPasswd 
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   5040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3840
      Width           =   3165
   End
   Begin VB.CheckBox checkrecu 
      Caption         =   "Recordar Password"
      Height          =   195
      Left            =   8000
      TabIndex        =   0
      Top             =   4200
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recordar Contraseña"
      Height          =   255
      Left            =   6360
      TabIndex        =   3
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Image ImgSalir 
      Height          =   255
      Left            =   3840
      Top             =   4560
      Width           =   975
   End
   Begin VB.Image imgConectarse 
      Height          =   255
      Left            =   7080
      Top             =   4560
      Width           =   1095
   End
End
Attribute VB_Name = "frmConnect2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub imgConectarse_Click()

    If checkrecu.value = 1 Then
    If Not StringIsRecup(UCase$(txtNombre.Text)) Then
        Call SaveRecu(txtNombre.Text, txtPasswd.Text)
    End If
End If
#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
#Else
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
#End If
    
    'update user info
    UserName = txtNombre.Text
    
    Dim aux As String
    aux = txtPasswd.Text
    
#If SeguridadAlkon Then
    UserPassword = md5.GetMD5String(aux)
    Call md5.MD5Reset
#Else
    UserPassword = aux
#End If
    If CheckUserData(False) = True Then
        EstadoLogin = Normal
        
#If UsarWrench = 1 Then
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
#Else
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If

    End If
    
End Sub

Private Sub imgSalir_Click()
    Unload Me
End Sub

Private Sub txtNombre_Change()
 
Dim f As Long
For f = 1 To MaxRecu
    If UCase$(txtNombre.Text) = UCase$(Recu(f).Nick) Then
        txtPasswd.Text = Recu(f).Password
    End If
Next f
 
End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then imgConectarse_Click
End Sub
