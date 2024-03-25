VERSION 5.00
Begin VB.Form frmConnectAcc 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Conexion de la cuenta"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3990
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   360
      Left            =   390
      TabIndex        =   5
      Top             =   2160
      Width           =   1440
   End
   Begin VB.CommandButton cmdConectar 
      Caption         =   "Conectar"
      Height          =   360
      Left            =   2220
      TabIndex        =   4
      Top             =   2160
      Width           =   1440
   End
   Begin VB.TextBox txtAccPass 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "123456"
      Top             =   1500
      Width           =   2265
   End
   Begin VB.TextBox txtAccName 
      Height          =   315
      Left            =   810
      TabIndex        =   1
      Text            =   "BetaTester@test.com"
      Top             =   780
      Width           =   2265
   End
   Begin VB.Label lblContraseña 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   870
      TabIndex        =   2
      Top             =   1170
      Width           =   2280
   End
   Begin VB.Label lblNombreDe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email de la cuenta:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   900
      TabIndex        =   0
      Top             =   450
      Width           =   1815
   End
End
Attribute VB_Name = "frmConnectAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Reparado por Lorwik
' El cliente original carga los datos de la cuenta del launcher, asi que hice este form

Option Explicit

Private Sub cmdCerrar_Click()
    Call CloseClient
End Sub

Private Sub cmdConectar_Click()
    
    LastDataAccount = txtAccName.Text
    LastDataPasswd = txtAccPass.Text
    
    ' Solicitamos la conexión
         
    Account.Email = LastDataAccount
    Account.Passwd = LastDataPasswd
    Prepare_And_Connect E_MODO.e_LoginAccount
    
    FrmConectando.visible = False

End Sub

