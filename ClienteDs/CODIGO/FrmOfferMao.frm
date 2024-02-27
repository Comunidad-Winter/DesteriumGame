VERSION 5.00
Begin VB.Form FrmOfferMao 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmOfferMao.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstOfferSend 
      Appearance      =   0  'Flat
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
      Height          =   2370
      Left            =   600
      TabIndex        =   1
      Top             =   1920
      Width           =   10800
   End
   Begin VB.ListBox lstOfferReceive 
      Appearance      =   0  'Flat
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
      Height          =   2370
      Left            =   720
      TabIndex        =   0
      Top             =   5160
      Width           =   10680
   End
   Begin VB.Image imgAccept 
      Height          =   375
      Left            =   600
      Top             =   7680
      Width           =   2415
   End
   Begin VB.Image imgBorrarOferta 
      Height          =   375
      Left            =   8520
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Image ImgUnload 
      Height          =   375
      Left            =   8760
      Top             =   7680
      Width           =   2535
   End
End
Attribute VB_Name = "FrmOfferMao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imgAccept_Click()
    
    If lstOfferReceive.ListIndex = -1 Then Exit Sub
    
    Dim temp As String
    
    temp = InputBox("Escribe la clave PIN de la cuenta para confirmar.", "Control de seguridad")
    
    If temp = vbNullString Then
        MsgBox "No puede estar vacío"
        Exit Sub
    End If
    
    Protocol.WriteMercadoAcceptInvitation temp, lstOfferReceive.ListIndex
End Sub

Private Sub ImgUnload_Click()
    Protocol.WriteRequestMercado
    Unload Me
End Sub
