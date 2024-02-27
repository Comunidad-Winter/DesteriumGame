VERSION 5.00
Begin VB.Form FrmOfertasMao 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmOfertasMao.frx":0000
   ScaleHeight     =   4500
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstPjs 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   2790
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2640
      Top             =   0
      Width           =   375
   End
   Begin VB.Image ImgAceptar 
      Height          =   495
      Left            =   480
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Image ImgRechazar 
      Height          =   375
      Left            =   480
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "FrmOfertasMao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Public LastButtonPressed As clsGraphicalButton

Public BotonAceptar As clsGraphicalButton
Public BotonRechazar As clsGraphicalButton

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_Load()
'    Call LoadButtons

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
End Sub

Private Sub LoadButtons()

    Dim GrhPath As String

    GrhPath = DirButtons

    Set BotonAceptar = New clsGraphicalButton
    Set BotonRechazar = New clsGraphicalButton

    Set LastButtonPressed = New clsGraphicalButton

    Call BotonAceptar.Initialize(ImgAceptar, GrhPath & "BotonAceptarMAO.jpg", _
                                 GrhPath & "BotonAceptarMAO1.jpg", _
                                 GrhPath & "BotonAceptarMAO.jpg", Me)

    Call BotonRechazar.Initialize(ImgRechazar, GrhPath & "BotonRechazarMAO.jpg", _
                                  GrhPath & "BotonRechazarMAO1.jpg", _
                                  GrhPath & "BotonRechazarMAO.jpg", Me)
End Sub

Private Sub Image1_Click()
    On Error Resume Next
    Unload Me
    FrmMercado.SetFocus
End Sub

Private Sub Image2_Click()
    If lstPjs.ListIndex < 0 Then Exit Sub
    If Not lstPjs.Text = vbNullString Then
        Call WriteRequestCharInfo(lstPjs.ListIndex + 1)      'Devuelve un numero, un index
    End If

End Sub

Private Sub imgAceptar_Click()
    If lstPjs.ListIndex < 0 Then Exit Sub
    Call WritePacketMercado(AceptarOferta, lstPjs.ListIndex + 1)
End Sub

Private Sub imgRechazar_Click()
    If lstPjs.ListIndex < 0 Then Exit Sub
    Call WritePacketMercado(RechazarOferta, lstPjs.ListIndex + 1)
End Sub
 
