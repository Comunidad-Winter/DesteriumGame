VERSION 5.00
Begin VB.Form frmSubastar 
   BorderStyle     =   0  'None
   Caption         =   "Compra y venta de objetos"
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   30
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSubastarsa.frx":0000
   ScaleHeight     =   5850
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstObjetos 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4530
      Left            =   390
      TabIndex        =   0
      Top             =   360
      Width           =   3000
   End
   Begin VB.Image Command1 
      Height          =   615
      Left            =   3600
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Image Command2 
      Height          =   495
      Left            =   3600
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Subastas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   1
      Top             =   1800
      Width           =   9495
   End
End
Attribute VB_Name = "frmSubastar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'JAO
Call WriteCompraSubasta(0, lstObjetos.ListIndex + 1)
End Sub

Private Sub Command2_Click()
Unload frmSubastar
End Sub

