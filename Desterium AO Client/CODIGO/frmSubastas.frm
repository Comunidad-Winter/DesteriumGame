VERSION 5.00
Begin VB.Form frmSubastas 
   BorderStyle     =   0  'None
   Caption         =   "Subastas"
   ClientHeight    =   5805
   ClientLeft      =   0
   ClientTop       =   30
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSubastas.frx":0000
   ScaleHeight     =   5805
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPrecio 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   810
      Left            =   888
      TabIndex        =   2
      Text            =   "100.000"
      Top             =   1977
      Width           =   3733
   End
   Begin VB.Image Command1 
      Height          =   495
      Left            =   1920
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Estime un precio :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   1920
      Width           =   6495
   End
   Begin VB.Label lblItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Objeto seleccionado."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   0
      Top             =   960
      Width           =   6735
   End
End
Attribute VB_Name = "frmSubastas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Not IsNumeric(txtPrecio) Then
Call MsgBox("El precio no es numérico.", vbInformation)
Else
Call WriteSubastarObjeto(txtPrecio)
Unload frmSubastas
End If
End Sub

Private Sub Form_Load()
Dim ItemSlot As Byte
ItemSlot = Inventario.SelectedItem
lblItem.Caption = Inventario.ItemName(ItemSlot)
End Sub

