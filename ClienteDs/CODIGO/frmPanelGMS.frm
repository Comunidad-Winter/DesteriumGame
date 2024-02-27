VERSION 5.00
Begin VB.Form frmPanelGMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desterium AO Panel GM"
   ClientHeight    =   4470
   ClientLeft      =   3225
   ClientTop       =   1500
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command15 
      Caption         =   "Dungeon Magma"
      Height          =   495
      Left            =   6360
      TabIndex        =   20
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Activar Global"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4080
      Width           =   3015
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Atención si usas este comando 1 no se caeran y a la segunda vez caeran los items"
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   3015
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Isla Veriil"
      Height          =   495
      Left            =   3240
      TabIndex        =   16
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Arghal"
      Height          =   495
      Left            =   4800
      TabIndex        =   15
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Fuerte Pretoriano"
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Dungeon Newbie"
      Height          =   495
      Left            =   4800
      TabIndex        =   13
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Sala Teleports"
      Height          =   495
      Left            =   6360
      TabIndex        =   12
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Ciudad Oscura"
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Dungeon Marabel"
      Height          =   495
      Left            =   6360
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Lindos"
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Banderbill"
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Mapa GM"
      Height          =   495
      Left            =   6360
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Nix"
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Ullathorpe"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   2955
   End
   Begin VB.CommandButton Command2 
      Caption         =   "/VERPROCESO"
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1600
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "/CAPTIONS"
      CausesValidation=   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4850
      TabIndex        =   18
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPanelGMS.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AntiCheat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmPanelGMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
          '/CAPTIONS
          Dim Nick As String

10        Nick = Combo1.Text
          
20        If LenB(Nick) <> 0 Then Call WriteRequieredCaptions(Nick)
End Sub

Private Sub Command15_Click()
10    WriteWarpChar UserName, 175, 50, 50
End Sub

Private Sub Command2_Click()
          '/VERPROCESOS
          Dim Nick As String

10        Nick = Combo1.Text
          
20        If LenB(Nick) <> 0 Then Call WriteLookProcess(Nick)


End Sub

Private Sub Command16_Click()
10    WriteWarpChar UserName, 34, 26, 72
End Sub

Private Sub Command17_Click()
10    WriteWarpChar UserName, 1, 50, 50
End Sub

Private Sub Command3_Click()
10    WriteWarpChar UserName, 194, 26, 33
End Sub

Private Sub Command4_Click()
10    WriteWarpChar UserName, 59, 43, 49
End Sub

Private Sub Command5_Click()
10    WriteWarpChar UserName, 62, 71, 41
End Sub

Private Sub Command6_Click()
10    WriteWarpChar UserName, 115, 45, 91
End Sub

Private Sub Command7_Click()
10    WriteWarpChar UserName, 185, 53, 22
End Sub

Private Sub Command8_Click()
10    WriteWarpChar UserName, 197, 49, 44
End Sub

Private Sub Command9_Click()
10    WriteWarpChar UserName, 168, 50, 57
End Sub

Private Sub Command10_Click()
10    WriteWarpChar UserName, 196, 35, 30
End Sub

Private Sub Command11_Click()
10    WriteWarpChar UserName, 151, 40, 48
End Sub

Private Sub Command12_Click()
10    Call ParseUserCommand("/caer")
End Sub

Private Sub Command13_Click()
10    WriteWarpChar UserName, 98, 46, 52
End Sub


Private Sub Command14_Click()
10    Call ParseUserCommand("/ACTIVARGLOBAL")
End Sub

Private Sub sa_Click()
10    MsgBox _
          "Para enviar un torneo automatico deathmatch debes ingresar el comando /EVENTO 2@0 Si es por item /evento 2@1 siempre 2 por que si o si los cupos tienen que ser 2."
End Sub

