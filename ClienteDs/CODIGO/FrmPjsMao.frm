VERSION 5.00
Begin VB.Form FrmPjsMao 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   Picture         =   "FrmPjsMao.frx":0000
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
      Left            =   135
      TabIndex        =   0
      Top             =   545
      Width           =   2760
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   255
   End
   Begin VB.Image ImgComprar 
      Height          =   375
      Left            =   600
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Image ImgOfrecer 
      Height          =   375
      Left            =   600
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2640
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "FrmPjsMao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Image1_Click()
Unload Me
'FrmMercado.SetFocus
End Sub


Private Sub Image2_Click()
Call AddtoRichTextBox(frmMain.RecTxt, "Presiona doble click para ver las estadísticas del personaje. Sólo se muestran las estadísticas de los personajes en modo cambio.", 0, 200, 200, False, False)
End Sub
Private Sub image2_dblclick()
Call Audio.PlayWave(SND_CLICK)
    Dim Nick As String

    Nick = lstPjs.Text
    

    If LenB(Nick) <> 0 Then
        If InStr(1, Nick, "-") Then
        Dim tmpstr() As String
        
        
        'Funciona, es una pelotudes lo que hice pero funciona.
        tmpstr = Split(Nick, "-")
        Call WriteRequestCharInfoMercado(tmpstr(0))
        Exit Sub
        End If
        
        Call WriteRequestCharInfoMercado(Nick)
    End If
    
End Sub

Private Sub imgComprar_Click()
Call Audio.PlayWave(SND_CLICK)
If lstPjs.List(lstPjs.ListIndex) <> vbNullString Then
If MsgBox("¿Seguro que desea comprar ese personaje?", vbYesNo + vbQuestion, "Desterium AO") = vbYes Then
Call WritePacketMercado(ComprarPJ, lstPjs.ListIndex + 1)
Else
            Exit Sub
        End If
    End If
End Sub

Private Sub ImgOfrecer_Click()
Call Audio.PlayWave(SND_CLICK)

If lstPjs.List(lstPjs.ListIndex) <> vbNullString Then
If MsgBox("¿Seguro que deseas ofertar tu personaje a cambio de este? Tu contraseña/pin/email pasarán a ser los datos del personaje recibido.", vbYesNo + vbQuestion, "Desterium AO") = vbYes Then
Call WritePacketMercado(EnviarOferta, lstPjs.ListIndex + 1)
Else
            Exit Sub
            End If
            End If
End Sub
