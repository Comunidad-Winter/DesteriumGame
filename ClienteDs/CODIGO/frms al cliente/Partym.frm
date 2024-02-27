VERSION 5.00
Begin VB.Form frmParty 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4890
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Partym.frx":0000
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   326
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picReward 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   0
      Picture         =   "Partym.frx":2655A
      ScaleHeight     =   2895
      ScaleWidth      =   4815
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Image imgVolverObtenido 
         Height          =   375
         Left            =   3240
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblOroObtenido 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100.000.000.000"
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
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   1250
         Width           =   2175
      End
      Begin VB.Label lblExpObtenida 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100.000.000.000"
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
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   960
         Width           =   2175
      End
   End
   Begin VB.PictureBox PicSolicitudes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   0
      Picture         =   "Partym.frx":3406B
      ScaleHeight     =   2895
      ScaleWidth      =   4815
      TabIndex        =   16
      Top             =   2160
      Visible         =   0   'False
      Width           =   4815
      Begin VB.ListBox lstRequest 
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   2190
         Left            =   360
         TabIndex        =   19
         Top             =   240
         Width           =   4095
      End
      Begin VB.Image imgVolver 
         Height          =   375
         Left            =   3240
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Image imgRechazar 
         Height          =   375
         Left            =   1680
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Image imgAceptar 
         Height          =   375
         Left            =   240
         Top             =   2520
         Width           =   1455
      End
   End
   Begin VB.Label lblSavePorcentaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Guardar porcentajes"
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
      Height          =   255
      Left            =   2160
      TabIndex        =   21
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   4395
      TabIndex        =   20
      Top             =   90
      Width           =   375
   End
   Begin VB.Image imgBons 
      Height          =   240
      Index           =   3
      Left            =   240
      Picture         =   "Partym.frx":424CC
      Top             =   5955
      Width           =   240
   End
   Begin VB.Image imgBons 
      Height          =   240
      Index           =   2
      Left            =   240
      Picture         =   "Partym.frx":466CD
      Top             =   5670
      Width           =   240
   End
   Begin VB.Image imgBons 
      Height          =   240
      Index           =   1
      Left            =   240
      Picture         =   "Partym.frx":4A8CE
      Top             =   5400
      Width           =   240
   End
   Begin VB.Image imgBons 
      Height          =   240
      Index           =   0
      Left            =   240
      Picture         =   "Partym.frx":4EACF
      Top             =   5160
      Width           =   240
   End
   Begin VB.Image imgSolicitudes 
      Height          =   375
      Left            =   3240
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Image imgReward 
      Height          =   375
      Left            =   1680
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Image imgAbandonate 
      Height          =   375
      Left            =   240
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   14
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   13
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   12
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   11
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lblOro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   10
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lblOro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   9
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblOro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   8
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label lblOro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   7
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   6
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblOro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   5
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Vacio>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Vacio>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Vacio>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Vacio>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Vacio>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Image boton 
      Height          =   255
      Index           =   0
      Left            =   3720
      Top             =   7500
      Visible         =   0   'False
      Width           =   15
   End
End
Attribute VB_Name = "frmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Picture = LoadPicture(App.path & "\Recursos\Grupos\GrupoMiembros.jpg")
End Sub

Private Sub imgAbandonate_Click()
    Protocol.WritePartyClient 4
    Unload Me
End Sub

Private Sub imgAceptar_Click()
    If lstRequest.ListIndex = -1 Then Exit Sub
    
    WriteGroupMember 1, lstRequest.List(lstRequest.ListIndex)
    Unload Me
End Sub

Private Sub imgRechazar_Click()
    If lstRequest.ListIndex = -1 Then Exit Sub
    
    WriteGroupMember 2, lstRequest.List(lstRequest.ListIndex)
    
    Unload Me
End Sub

Private Sub imgReward_Click()
    Protocol.WritePartyClient 3
    Me.Picture = LoadPicture(App.path & "\Recursos\Grupos\GrupoObtenido.jpg")
    
    picReward.Visible = True
End Sub

Private Sub imgSolicitudes_Click()
    Protocol.WritePartyClient 2
    Me.Picture = LoadPicture(App.path & "\Recursos\Grupos\GrupoSolicitudes.jpg")
    
    PicSolicitudes.Visible = True
End Sub

Private Sub imgVolver_Click()
    PicSolicitudes.Visible = False
    
    Me.Picture = LoadPicture(App.path & "\Recursos\Grupos\GrupoMiembros.jpg")
End Sub

Private Sub imgVolverObtenido_Click()
    picReward.Visible = False
    
    Me.Picture = LoadPicture(App.path & "\Recursos\Grupos\GrupoMiembros.jpg")
End Sub

Private Sub Label1_Click()
    Default
    Unload Me
End Sub

Private Sub lblExp_Click(Index As Integer)
    Dim temp
    temp = InputBox("Elige el porcentaje de experiencia que deseas para este personaje", "Grupos Desterium: Edición de porcentaje", lblExp(Index).Caption)

    lblExp(Index).Caption = Val(temp)
End Sub

Private Sub lblOro_Click(Index As Integer)
    Dim temp
    
    temp = InputBox("Elige el porcentaje  de oro que deseas para este personaje", "Grupos Desterium: Edición de porcentaje", lblOro(Index).Caption)
    
    lblOro(Index).Caption = Val(temp)
End Sub

Private Sub lblSavePorcentaje_Click()
    Dim Exp(4) As Byte
    Dim Oro(4) As Byte
    Dim A As Byte
    Dim TotalExp As Integer
    Dim TotalGld As Integer
    
    For A = 0 To 4
        Exp(A) = lblExp(A)
        Oro(A) = lblOro(A)
        
        TotalExp = TotalExp + Exp(A)
        TotalGld = TotalGld + Oro(A)
        
    Next A
    
    If TotalExp <> 100 Then Exit Sub
    If TotalGld <> 100 Then Exit Sub
    
    WriteGroupChangePorc Exp, Oro
End Sub

Private Sub Default()
    Dim A As Long
    
    For A = 0 To 4
        lblUser(A).Caption = "<Vacio>"
        lblExp(A).Caption = "0"
        lblOro(A).Caption = "0"
    Next A
    
End Sub

