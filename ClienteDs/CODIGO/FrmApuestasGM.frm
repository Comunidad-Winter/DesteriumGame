VERSION 5.00
Begin VB.Form FrmApuestasGM 
   Caption         =   "Realización de apuestas"
   ClientHeight    =   9135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form4"
   ScaleHeight     =   9135
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Otorgar PREMIO (Se cierra el sistema)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   195
      TabIndex        =   18
      Top             =   7020
      Width           =   4305
      Begin VB.TextBox txtApuesta 
         Height          =   285
         Left            =   585
         TabIndex        =   21
         Top             =   780
         Width           =   2745
      End
      Begin VB.CommandButton Command4 
         Caption         =   "OTORGAR PREMIO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   780
         TabIndex        =   19
         Top             =   1365
         Width           =   2550
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Elige el nombre de la apuesta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   195
         TabIndex        =   20
         Top             =   390
         Width           =   3915
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Anti boludos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   4875
      TabIndex        =   15
      Top             =   6045
      Width           =   5280
      Begin VB.CommandButton Command3 
         Caption         =   "CANCELAR APUESTAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1365
         TabIndex        =   17
         Top             =   1365
         Width           =   2550
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Si te confundiste podrás cancelar la apuesta. Y las apuestas que ya fueron hechas seran devueltas a los usuarios correspondientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   195
         TabIndex        =   16
         Top             =   390
         Width           =   3915
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Apuestas de los personajes"
      Height          =   5475
      Left            =   4875
      TabIndex        =   9
      Top             =   390
      Width           =   5280
      Begin VB.ListBox lstUsers 
         Height          =   3375
         Left            =   195
         TabIndex        =   11
         Top             =   975
         Width           =   2745
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ACTUALIZAR INFORMACIÓN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1170
         TabIndex        =   10
         Top             =   390
         Width           =   2550
      End
      Begin VB.Label lblGld 
         BackStyle       =   0  'Transparent
         Caption         =   "Oro apostado: PROXIMAMENTE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   13
         Top             =   4875
         Width           =   3915
      End
      Begin VB.Label lblDsp 
         BackStyle       =   0  'Transparent
         Caption         =   "Dsp apostado: PROXIMAMENTE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   12
         Top             =   4485
         Width           =   3915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nueva apuesta"
      Height          =   6645
      Left            =   195
      TabIndex        =   0
      Top             =   195
      Width           =   4305
      Begin VB.CommandButton Command2 
         Caption         =   "ENVIAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   780
         TabIndex        =   14
         Top             =   5850
         Width           =   2550
      End
      Begin VB.ComboBox cmbTime 
         Height          =   315
         Left            =   2730
         TabIndex        =   8
         Text            =   "Seleccionar"
         Top             =   3120
         Width           =   1380
      End
      Begin VB.TextBox txtDesc 
         Height          =   990
         Left            =   195
         TabIndex        =   6
         Top             =   4485
         Width           =   2745
      End
      Begin VB.ComboBox cmbAmount 
         Height          =   315
         Left            =   2730
         TabIndex        =   5
         Text            =   "Seleccionar"
         Top             =   2340
         Width           =   1380
      End
      Begin VB.TextBox txtApuestas 
         Height          =   990
         Left            =   195
         TabIndex        =   1
         Top             =   975
         Width           =   2745
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo que tienen los usuarios en apostar (en minutos) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   195
         TabIndex        =   7
         Top             =   2925
         Width           =   2355
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Una breve descripción que quieras poner:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   195
         TabIndex        =   4
         Top             =   3900
         Width           =   2355
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "¿Cuantas personas podrán apostar máximo?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   195
         TabIndex        =   3
         Top             =   2145
         Width           =   2355
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "¿A quien se le va a apostar? (separados por ',')"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   195
         TabIndex        =   2
         Top             =   390
         Width           =   3720
      End
   End
End
Attribute VB_Name = "FrmApuestasGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
10        Protocol.WriteRequestUsersApostando
          
End Sub

Private Sub Command2_Click()
10        If txtDesc.Text = vbNullString Then
20            MsgBox "No puedes dejar la descripción vacia"
30            Exit Sub
40        End If
          
50        If cmbTime.ListIndex = -1 Then
60            MsgBox "Selecciona un tiempo para que terminen las apuestas"
70            Exit Sub
80        End If
          
90        If cmbAmount.ListIndex = -1 Then
100           MsgBox "Selecciona la cantidad de usuarios máximos que pueden apostar"
110           Exit Sub
120       End If
          
130       WriteNewGamble txtDesc.Text, Val(cmbTime.List(cmbTime.ListIndex)), _
              Val(cmbAmount.List(cmbAmount.ListIndex)), txtApuestas.Text
End Sub

Private Sub Command3_Click()
10        WriteCancelGamble
End Sub

Private Sub Command4_Click()
10        If txtApuesta.Text = vbNullString Then
20            MsgBox "Elije el nombre de la apuesta ganadora"
30            Exit Sub
40        End If
          
50        Protocol.WriteWinGamble txtApuesta.Text
          
End Sub

Private Sub Form_Load()
10        cmbAmount.AddItem "10"
20        cmbAmount.AddItem "20"
30        cmbAmount.AddItem "30"
40        cmbAmount.AddItem "40"
50        cmbAmount.AddItem "50"
60        cmbAmount.AddItem "60"
70        cmbAmount.AddItem "70"
80        cmbAmount.AddItem "80"
90        cmbAmount.AddItem "90"
100       cmbAmount.AddItem "100"
          
110       cmbTime.AddItem "5"
120       cmbTime.AddItem "10"
130       cmbTime.AddItem "15"
140       cmbTime.AddItem "20"
150       cmbTime.AddItem "25"
160       cmbTime.AddItem "30"
End Sub

Private Sub lstUsers_Click()
    'Protocol.WriteRequestInfoUserApostando lstUsers.List(lstUsers.ListIndex)
    
End Sub
