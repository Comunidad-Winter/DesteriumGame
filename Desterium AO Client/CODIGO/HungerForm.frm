VERSION 5.00
Begin VB.Form HungerForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Juegos del Hambre"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   2805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Salir de los Juegos del Hambre"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2880
      Width           =   2800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Iniciar Juegos del Hambre"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   2800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar Juegos del Hambre"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   2800
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   2788
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Caen items Juegos del Hambre?"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Caen items:"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Valor de la inscripción:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad de Cupos:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "HungerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
WriteHungerGamesDelete
Unload Me
End Sub

Private Sub Command2_Click()
WriteHungerGamesCreate Val(Text1.Text), Val(Text2.Text), Check1.value
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
