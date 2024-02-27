VERSION 5.00
Begin VB.Form FrmRGM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Soporte"
   ClientHeight    =   4125
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   4890
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   2775
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nick del Usuario a enviar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "FrmRGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call WriteResponderGm((Text1.Text), (Text2.Text))
End Sub

Private Sub Form_Load()

End Sub
