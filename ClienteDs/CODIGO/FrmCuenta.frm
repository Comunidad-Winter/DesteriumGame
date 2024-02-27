VERSION 5.00
Begin VB.Form FrmCuenta 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Lista de Personajes"
   ClientHeight    =   8685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   Picture         =   "FrmCuenta.frx":0000
   ScaleHeight     =   579
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicCambiarPw 
      BackColor       =   &H80000012&
      Height          =   4215
      Left            =   2745
      ScaleHeight     =   4155
      ScaleWidth      =   5715
      TabIndex        =   6
      Top             =   2805
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox txtPwNew 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   2205
         TabIndex        =   29
         Top             =   2550
         Width           =   2910
      End
      Begin VB.TextBox txtAcc 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2220
         TabIndex        =   16
         Top             =   1050
         Width           =   2895
      End
      Begin VB.TextBox txtPwOld 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2220
         TabIndex        =   15
         Top             =   2130
         Width           =   2895
      End
      Begin VB.TextBox txtMail 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2220
         TabIndex        =   14
         Top             =   1755
         Width           =   2895
      End
      Begin VB.TextBox txtPinaso 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2220
         TabIndex        =   13
         Top             =   1395
         Width           =   2880
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CAMBIAR CONTRASEÑA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1200
         TabIndex        =   12
         Top             =   3120
         Width           =   3375
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Nueva Contraseña:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   300
         Left            =   315
         TabIndex        =   28
         Top             =   2535
         Width           =   1740
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
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
         Left            =   5520
         TabIndex        =   17
         Top             =   3840
         Width           =   135
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña Actual:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   255
         TabIndex        =   11
         Top             =   2175
         Width           =   1875
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   1365
         TabIndex        =   10
         Top             =   1785
         Width           =   660
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "PIN:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1575
         TabIndex        =   9
         Top             =   1380
         Width           =   390
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de la Cuenta:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   150
         TabIndex        =   8
         Top             =   1125
         Width           =   2010
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmCuenta.frx":59F4E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   765
         Left            =   360
         TabIndex        =   7
         Top             =   225
         Width           =   5295
      End
   End
   Begin VB.PictureBox PicCambiarPwS 
      BorderStyle     =   0  'None
      Height          =   3945
      Left            =   960
      Picture         =   "FrmCuenta.frx":59FE9
      ScaleHeight     =   3945
      ScaleWidth      =   6000
      TabIndex        =   18
      Top             =   9630
      Visible         =   0   'False
      Width           =   6000
      Begin VB.TextBox txtPwNewS 
         BackColor       =   &H00000000&
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
         Height          =   285
         Left            =   2805
         TabIndex        =   27
         Top             =   3015
         Width           =   2415
      End
      Begin VB.TextBox txtPwOldS 
         BackColor       =   &H00000000&
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
         Height          =   285
         Left            =   2805
         TabIndex        =   25
         Top             =   2760
         Width           =   2445
      End
      Begin VB.TextBox txtAccs 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   285
         Left            =   2730
         TabIndex        =   21
         Top             =   1450
         Width           =   2550
      End
      Begin VB.TextBox txtMails 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   285
         Left            =   2790
         TabIndex        =   20
         Top             =   2430
         Width           =   2550
      End
      Begin VB.TextBox txtPinasos 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   285
         Left            =   2760
         TabIndex        =   19
         Top             =   1935
         Width           =   2550
      End
      Begin VB.Label Label10 
         Caption         =   "pw nueva"
         Height          =   180
         Left            =   1920
         TabIndex        =   26
         Top             =   3060
         Width           =   720
      End
      Begin VB.Label Label9 
         Caption         =   "pw actual"
         Height          =   225
         Left            =   1875
         TabIndex        =   24
         Top             =   2775
         Width           =   795
      End
      Begin VB.Label Label8 
         Caption         =   "Meil"
         Height          =   255
         Left            =   2175
         TabIndex        =   23
         Top             =   2415
         Width           =   465
      End
      Begin VB.Label Label7 
         Caption         =   "Pin"
         Height          =   240
         Left            =   2115
         TabIndex        =   22
         Top             =   1950
         Width           =   375
      End
      Begin VB.Image Image4 
         Height          =   405
         Left            =   390
         Top             =   3315
         Width           =   2355
      End
      Begin VB.Image Image5 
         Height          =   405
         Left            =   3315
         Top             =   3300
         Width           =   2355
      End
   End
   Begin VB.CommandButton MOMENT 
      Caption         =   "AGREGAR PJS EXISTENTES      (BOTÓN TEMPORAL)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7935
      TabIndex        =   5
      Top             =   10530
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.ListBox lstPjs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   3270
      ItemData        =   "FrmCuenta.frx":789FF
      Left            =   585
      List            =   "FrmCuenta.frx":78A06
      TabIndex        =   0
      Top             =   3900
      Width           =   2940
   End
   Begin VB.Image Image6 
      Height          =   360
      Left            =   8520
      Top             =   3105
      Width           =   2370
   End
   Begin VB.Image Image3 
      Height          =   405
      Left            =   12270
      Top             =   750
      Width           =   2475
   End
   Begin VB.Image Image2 
      Height          =   1590
      Left            =   8385
      Top             =   3675
      Width           =   2550
   End
   Begin VB.Image imgSalir 
      Height          =   210
      Left            =   8385
      Top             =   5655
      Width           =   2550
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   8385
      Top             =   2400
      Width           =   2550
   End
   Begin VB.Image imgRemove 
      Height          =   405
      Left            =   780
      Top             =   7695
      Width           =   2550
   End
   Begin VB.Image imgConnect 
      Height          =   405
      Left            =   780
      Top             =   7215
      Width           =   2550
   End
   Begin VB.Label lblMap 
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   405
      Left            =   4095
      TabIndex        =   4
      Top             =   5475
      Width           =   3135
   End
   Begin VB.Label lblRaza 
      BackStyle       =   0  'Transparent
      Caption         =   "Raza:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   405
      Left            =   4095
      TabIndex        =   3
      Top             =   4980
      Width           =   3135
   End
   Begin VB.Label lblNivel 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   405
      Left            =   4095
      TabIndex        =   2
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label lblClase 
      BackStyle       =   0  'Transparent
      Caption         =   "Clase:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   405
      Left            =   4125
      TabIndex        =   1
      Top             =   4470
      Width           =   3135
   End
End
Attribute VB_Name = "FrmCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
          
    With Cuenta
        .Account = txtAcc.Text
        .PIN = txtPinaso.Text
        .Email = txtMail.Text
        .Passwd = txtPwOld.Text
        .NewPasswd = txtPwNew.Text
    End With
    
    'If CheckDataAccount = False Then Exit Sub

    EstadoLogin = e_ChangePasswdAccount

    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
    
    Call Login

    PicCambiarPw.Visible = False
    
End Sub

Private Sub Image1_Click()
    frmCrearPersonaje.Show vbModeless, frmMain
End Sub

Private Sub Image2_Click()
    MsgBox "Botones deshabilitados momentaneamente"
End Sub

Private Sub Image3_Click()
PicCambiarPw.Visible = True
End Sub



Private Sub Image5_Click()
If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
          
    With Cuenta
        .Account = txtAcc.Text
        .PIN = txtPinaso.Text
        .Email = txtMail.Text
        .Passwd = txtPwOld.Text
        .NewPasswd = txtPwNew.Text
    End With
    
    'If CheckDataAccount = False Then Exit Sub

    EstadoLogin = e_ChangePasswdAccount

    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
    
    Call Login

PicCambiarPw.Visible = False
End Sub

Private Sub Image6_Click()
PicCambiarPw.Visible = True
End Sub

Private Sub imgConnect_Click()
    If SelectedChar = 0 Then
        MsgBox "Selecciona el personaje al que deseas entrar."
        Exit Sub
    End If
    
    
    Dim TempName As String
    
    TempName = lstPjs.List(SelectedChar - 1)
    
    UserName = InputBox("Ingrese como le gustaría que su nick luzca ACTUALMENTE LUCE: '" & TempName & " '", , TempName)
    
    If UCase$(UserName) <> UCase$(TempName) Then Exit Sub
    'UserName = lstPjs.List(SelectedChar - 1)
    EstadoLogin = e_LoginCharAccount
    
    Call Login
End Sub

Private Sub imgRemove_Click()
    If SelectedChar = 0 Then
        MsgBox "Selecciona el personaje al que deseas entrar."
        Exit Sub
    End If
    
    
    UserName = lstPjs.List(SelectedChar - 1)
    
    If MsgBox("¿Estás seguro que deseas borrar el personaje " & UserName & "?", vbYesNo) = vbYes Then
    
        EstadoLogin = e_RemoveCharAccount
    
        Call Login
    End If
End Sub

Private Sub Label3_Click()
    frmCrearPersonaje.Show
End Sub


Private Sub lblKillChar_Click()
    
    If SelectedChar = 0 Then
        MsgBox "Selecciona el personaje al que deseas entrar."
        Exit Sub
    End If
    
    UserName = lstPjs.List(SelectedChar - 1)
    EstadoLogin = e_RemoveCharAccount
    
    Call Login
End Sub

Private Sub lblLogin_Click()
    

End Sub

Private Sub imgSalir_Click()
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    
    frmConnect.Show
    Unload Me
End Sub

Private Sub Label6_Click()
 PicCambiarPw.Visible = False
End Sub

Private Sub lstPjs_Click()
    Dim Index As Integer
    
    Index = lstPjs.ListIndex + 1
    
    If Index <= 0 Then Exit Sub
    
    SelectedChar = Index
    
    If lstPjs.List(lstPjs.ListIndex) = "(Vacio)" Then Exit Sub
    
    
    With CuentaChars(Index)
        If .Clase = 0 Then Exit Sub
        lblClase.Caption = "Clase: " & ListaClases(.Clase)
        lblRaza.Caption = "Raza: " & ListaRazas(.Raza)
        lblNivel.Caption = "Nivel: " & .Elv
    
    End With
    
    
End Sub

