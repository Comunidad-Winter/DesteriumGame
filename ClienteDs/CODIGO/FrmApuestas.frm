VERSION 5.00
Begin VB.Form FrmApuestas 
   Caption         =   "¡Apuesta por tus favoritos!"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5625
   LinkTopic       =   "Form4"
   ScaleHeight     =   5475
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Realiza tu apuesta aquí"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3330
      Left            =   195
      TabIndex        =   1
      Top             =   1755
      Width           =   5085
      Begin VB.CommandButton Command1 
         Caption         =   "¡APOSTAR!"
         Height          =   405
         Left            =   2925
         TabIndex        =   7
         Top             =   2535
         Width           =   1575
      End
      Begin VB.TextBox txtGld 
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Top             =   1950
         Width           =   1770
      End
      Begin VB.TextBox txtDsp 
         Height          =   285
         Left            =   3120
         TabIndex        =   5
         Top             =   1560
         Width           =   1770
      End
      Begin VB.ListBox lstApuestas 
         Height          =   1815
         Left            =   195
         TabIndex        =   2
         Top             =   1170
         Width           =   2160
      End
      Begin VB.Label Label4 
         Caption         =   "¡Selecciona de la lista al personaje o a los personajes que aparezcan. Podrás apostar por un solo bando a la vez."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   390
         TabIndex        =   8
         Top             =   390
         Width           =   4500
      End
      Begin VB.Label Label3 
         Caption         =   "Oro:"
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
         Left            =   2535
         TabIndex        =   4
         Top             =   1950
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "Dsp:"
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
         Left            =   2535
         TabIndex        =   3
         Top             =   1560
         Width           =   600
      End
   End
   Begin VB.Label Label1 
      Caption         =   "JUGAR COMPULSIVAMENTE ES PERJUDICIAL PARA LA SALUD. EL STAFF NO SE RESPONSABILIZA DE MALAS APUESTAS."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1575
      Left            =   195
      TabIndex        =   0
      Top             =   195
      Width           =   5475
   End
End
Attribute VB_Name = "FrmApuestas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
10        If lstApuestas.ListIndex = -1 Then
20            MsgBox "Selecciona a quien quieres realizarle la apuesta"
30            Exit Sub
40        End If
          
50        If Val(txtDsp.Text) = 0 And Val(txtGld.Text) = 0 Then
60            MsgBox "Debes elegir que apostar"
70            Exit Sub
80        End If
          
90        Protocol.WriteSendGamble lstApuestas.ListIndex, Val(txtDsp.Text), _
              Val(txtGld.Text)
          
End Sub

