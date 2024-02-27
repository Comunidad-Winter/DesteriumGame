VERSION 5.00
Begin VB.Form FrmRECBORR 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tierras del Noroeste"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox DATO1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox DATO2 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox txtPIN 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblNOMBRE 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   660
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   540
   End
   Begin VB.Label lblPIN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   255
   End
End
Attribute VB_Name = "FrmRECBORR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
