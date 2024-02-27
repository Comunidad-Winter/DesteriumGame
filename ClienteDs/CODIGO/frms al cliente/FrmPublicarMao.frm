VERSION 5.00
Begin VB.Form FrmPublicarMao 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   Picture         =   "FrmPublicarMao.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMail 
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   210
      Left            =   1440
      TabIndex        =   6
      Top             =   2190
      Width           =   2055
   End
   Begin VB.TextBox txtPin 
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   210
      Left            =   1440
      TabIndex        =   5
      Top             =   1750
      Width           =   2055
   End
   Begin VB.PictureBox PicExtra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   840
      Picture         =   "FrmPublicarMao.frx":16AF7
      ScaleHeight     =   1170
      ScaleWidth      =   4335
      TabIndex        =   4
      Top             =   3720
      Width           =   4335
      Begin VB.TextBox txtOro 
         BackColor       =   &H00404080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Left            =   2880
         TabIndex        =   9
         Text            =   "0"
         Top             =   880
         Width           =   1335
      End
      Begin VB.TextBox txtDsp 
         BackColor       =   &H00404080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Left            =   720
         TabIndex        =   8
         Text            =   "0"
         Top             =   870
         Width           =   1335
      End
      Begin VB.TextBox txtNamePago 
         BackColor       =   &H00404080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.OptionButton OpVenta 
      BackColor       =   &H80000008&
      Caption         =   "Option1"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.OptionButton OpCambio 
      BackColor       =   &H80000008&
      Caption         =   "Option1"
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txtPw 
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   1440
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   4200
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "NICKNAME JEJEJE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   890
      Width           =   2295
   End
   Begin VB.Image ImgPublicar 
      Height          =   375
      Left            =   2400
      Top             =   5280
      Width           =   1335
   End
End
Attribute VB_Name = "FrmPublicarMao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsFormulario As clsFormMovementManager
Public LastPressed As clsGraphicalButton

Public BotonPublicar As clsGraphicalButton

Private Sub Form_Load()
          
          ' Handles Form movement (drag and drop).
10        Set clsFormulario = New clsFormMovementManager
20        clsFormulario.Initialize Me
          
30        lblName.Caption = UCase$(UserName)
End Sub

Private Sub Image1_Click()
10    Unload Me
      'FrmMercado.SetFocus
End Sub

Private Sub Image2_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    Unload Me
30    frmMain.SetFocus
End Sub

Private Sub ImgPublicar_Click()
10        Call Audio.PlayWave(SND_CLICK)
          
20        If Not CheckMailString(txtMail.Text) Then
30            MsgBox "Dirección de email inválida. Ingrese una nuevamente."
40            txtMail.Text = vbNullString
50            Exit Sub
60        End If
          
70        If txtPw.Text = vbNullString Then
80            MsgBox "Ingrese la contraseña del personaje " & UserName
90            Exit Sub
100       End If
          
110       If txtPin.Text = vbNullString Then
120           MsgBox "Ingrese el pin del personaje " & UserName
130           Exit Sub
140       End If
          
150       If OpCambio.value = True Then
160           Protocol.WritePublicationMAO txtMail.Text, txtPw.Text, txtPin.Text
              
170       ElseIf OpVenta.value = True Then
180           If txtOro.Text = vbNullString Or txtDsp.Text = vbNullString Then
190               MsgBox "¡¡ATENCIÓN!! El valor debe ser 0 como mínimo!"
200               Exit Sub
210           End If
              
220           If Val(txtOro.Text) <= 0 And Val(txtDsp.Text <= 0) Then
230               MsgBox "¡¡ATENCIÓN!! No se puede vender GRATIS!"
240               Exit Sub
250           End If
              
260           Protocol.WritePublicationMAO txtMail.Text, txtPw.Text, txtPin.Text, _
                  txtNamePago.Text, Val(txtOro.Text), Val(txtDsp.Text)
270       End If
          
280       Unload Me
End Sub

Private Sub OpCambio_Click()
10        Call Audio.PlayWave(SND_CLICK)
          
20        PicExtra.Visible = False
End Sub

Private Sub OpVenta_Click()
10        Call Audio.PlayWave(SND_CLICK)
          
20        PicExtra.Visible = True
End Sub
