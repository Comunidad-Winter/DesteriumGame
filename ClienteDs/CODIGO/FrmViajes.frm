VERSION 5.00
Begin VB.Form FrmViajes 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   Picture         =   "FrmViajes.frx":0000
   ScaleHeight     =   5040
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
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
      Height          =   2595
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   2565
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   120
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   2400
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label LblPrecio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   3720
      Width           =   1455
   End
End
Attribute VB_Name = "FrmViajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub Form_Load()
          ' Handles Form movement (drag and drop).
10        Set clsFormulario = New clsFormMovementManager
20        clsFormulario.Initialize Me
          
30        List1.AddItem "Muelle de Lindos"
40        List1.AddItem "Muelle de Nix"
50        List1.AddItem "Muelle de Arghâl"
60        List1.AddItem "Muelle de Banderbill"
70        List1.AddItem "Fortaleza del Rey Pretoriano"
80        LblPrecio.Caption = ""
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image3_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    Call WriteViajar(List1.ListIndex)
30    Unload Me
End Sub

Private Sub Image4_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    Unload Me
End Sub

Private Sub List1_Click()

10    If List1.ListIndex = 0 Then
20        LblPrecio.Caption = "10.000"
30    ElseIf List1.ListIndex = 1 Then
40        LblPrecio = "7.000"
50        ElseIf List1.ListIndex = 2 Then
60        LblPrecio = "15.000"
70    ElseIf List1.ListIndex = 3 Then
80        LblPrecio = "15.000"
90    ElseIf List1.ListIndex = 4 Then
100       LblPrecio = "350.000"
110   End If
End Sub
