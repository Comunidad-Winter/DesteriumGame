VERSION 5.00
Begin VB.Form FrMToRNeoC 
   Caption         =   "Form1"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   3600
      TabIndex        =   3
      Text            =   "itemS"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   1920
      TabIndex        =   2
      Text            =   "precio"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Text            =   "Cupos"
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      Caption         =   "Crear Torneo"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "FrMToRNeoC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Click()
Call WriteTorneoAutomatico(ByVal (Text3), ByVal (Text1), ByVal (Text2))
Unload Me
End Sub

Private Sub Form_Load()

End Sub
