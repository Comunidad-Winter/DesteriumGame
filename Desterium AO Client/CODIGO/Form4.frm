VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Postear personaje."
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5505
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Text            =   "Oro"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Text            =   "1"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Text            =   "Pj que recibe"
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Postear!"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del personaje que recibe el oro"
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel mínimo intercambio"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad de oro"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
WritePostearPJ Text1.Text, Text2.Text, Text3.Text
Unload Me
End Sub
