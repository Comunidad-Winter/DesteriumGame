VERSION 5.00
Begin VB.Form FrmNuevaCuenta 
   BorderStyle     =   0  'None
   Caption         =   "Creando una nueva cuenta..."
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmNuevaCuenta.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPin 
      BackColor       =   &H80000008&
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
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Left            =   4750
      TabIndex        =   4
      Top             =   6720
      Width           =   2550
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H80000008&
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
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Left            =   4750
      TabIndex        =   3
      Top             =   5950
      Width           =   2550
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H80000008&
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
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Index           =   1
      Left            =   4750
      TabIndex        =   2
      Top             =   5200
      Width           =   2550
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H80000008&
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
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Index           =   0
      Left            =   4750
      TabIndex        =   1
      Top             =   4400
      Width           =   2550
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H80000008&
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
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Left            =   4750
      TabIndex        =   0
      Top             =   3600
      Width           =   2550
   End
   Begin VB.Image imgVolver 
      Height          =   405
      Left            =   3315
      Top             =   7100
      Width           =   2550
   End
   Begin VB.Image imgCrear 
      Height          =   405
      Left            =   6045
      Top             =   7120
      Width           =   2550
   End
End
Attribute VB_Name = "FrmNuevaCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function CheckDataAccount() As Boolean
    
    Dim LoopC As Integer
    
    With Cuenta
    
        'Largo de los datos
        
        If Len(.Account) = 0 Or Len(.Account) > 15 Then
            MsgBox "Nombre nulo o supera los caracteres permitidos (15)"
            Exit Function
        End If
        
        If Len(.Email) = 0 Or Len(.Email) > 30 Then
            MsgBox "Email nulo o supera los caracteres permitidos (30)"
            Exit Function
        End If
        
        If Len(.Passwd) = 0 Or Len(.Passwd) > 20 Then
            MsgBox "Contraseña nula o supera los caracteres permitidos (20)"
            Exit Function
        End If
        
        If .Passwd <> txtPasswd(1).Text Then
            MsgBox "Las contraseña no coinciden."
            Exit Function
        End If
        
        
        If Len(.PIN) = 0 Or Len(.PIN) > 30 Then
            MsgBox "Pin nulo o supera los caracteres permitidos (30)"
            Exit Function
        End If
        
        
        For LoopC = 1 To Len(.Passwd)
           CharAscii = Asc(mid$(.Passwd, LoopC, 1))
           
           If Not LegalCharacter(CharAscii) Then
               MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & _
                      " no está permitido.")
               Exit Function
           End If
       Next LoopC
       
        For LoopC = 1 To Len(.Account)
           CharAscii = Asc(mid$(.Account, LoopC, 1))
           
           If Not LegalCharacter(CharAscii) Then
               MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & _
                      " no está permitido.")
               Exit Function
           End If
       Next LoopC
    End With
    
    CheckDataAccount = True
    
End Function

Private Sub Command_Click()


End Sub

Private Sub Label_Click(Index As Integer)

End Sub

Private Sub imgCrear_Click()
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
          
    With Cuenta
        .Account = txtName.Text
        .Email = txtEmail.Text
        .Passwd = txtPasswd(0).Text
        .PIN = txtPin.Text
    End With
    
    If CheckDataAccount = False Then Exit Sub

    EstadoLogin = e_NewAccount

    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
    
    Call Login
    
    frmConnect.Show
    Unload Me
End Sub

Private Sub imgVolver_Click()
    frmConnect.Show
    Unload Me
End Sub

