VERSION 5.00
Begin VB.Form frmParty 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Porcentajes"
   ClientHeight    =   2625
   ClientLeft      =   5925
   ClientTop       =   4170
   ClientWidth     =   2865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   2865
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtPorcentaje 
      Height          =   285
      Index           =   4
      Left            =   1680
      TabIndex        =   9
      Text            =   "100"
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtPorcentaje 
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   8
      Text            =   "100"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtPorcentaje 
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   7
      Text            =   "100"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtPorcentaje 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   6
      Text            =   "100"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtPorcentaje 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Text            =   "100"
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblUser 
      Caption         =   "USUARIO"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblUser 
      Caption         =   "USUARIO"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblUser 
      Caption         =   "USUARIO"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblUser 
      Caption         =   "USUARIO"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblUser 
      Caption         =   "USUARIO"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
