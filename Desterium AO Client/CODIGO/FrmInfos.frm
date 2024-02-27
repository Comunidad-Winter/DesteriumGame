VERSION 5.00
Begin VB.Form FrmInfos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Información del Personaje"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Oro 
      Caption         =   "Label2"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblName 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Mana 
      Caption         =   "Label3"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Clase 
      Caption         =   "Label4"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Raza 
      Caption         =   "Label5"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Statu 
      Caption         =   "Label6"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label NIVEL 
      Caption         =   "Label6"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "FrmInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

