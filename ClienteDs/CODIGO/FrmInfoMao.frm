VERSION 5.00
Begin VB.Form FrmInfoMao 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmInfoMao.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCopyAccount 
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
      Height          =   1590
      Left            =   2520
      TabIndex        =   5
      Top             =   5760
      Width           =   1320
   End
   Begin VB.ListBox lstAccount 
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
      Height          =   1590
      Left            =   720
      TabIndex        =   4
      Top             =   5760
      Width           =   1320
   End
   Begin VB.ListBox lstPjs 
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
      Height          =   1590
      Left            =   720
      TabIndex        =   3
      Top             =   3840
      Width           =   1320
   End
   Begin VB.Image ImgSalir 
      Height          =   375
      Left            =   8400
      Top             =   7560
      Width           =   2535
   End
   Begin VB.Image imgConfirm 
      Height          =   375
      Left            =   8400
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Label lblDsp 
      BackStyle       =   0  'Transparent
      Caption         =   "100.000.000.000.000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label lblGld 
      BackStyle       =   0  'Transparent
      Caption         =   "100.000.000.000.000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Hola esta es la poronga de publicación y aca no se va a ver una mierda jajajajaja xD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1920
      Width           =   9735
   End
End
Attribute VB_Name = "FrmInfoMao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imgConfirm_Click()
    Dim A As Long
    Dim Users As String
    
    If lstCopyAccount.ListCount > 0 Then
        For A = 0 To lstCopyAccount.ListCount - 1
            Users = Users & lstCopyAccount.List(A) & "-"
        Next A
        
        Users = mid$(Users, 1, Len(Users) - 1)
    End If
    
    Protocol.WriteSendOfferAccount SelectedListMAO, Users
    Unload Me
End Sub

Private Sub imgSalir_Click()
    Unload Me
End Sub

Private Sub lstAccount_dblClick()
    Dim SelectedPj As String
    Dim A As Long
    
    SelectedPj = lstAccount.List(lstAccount.ListIndex)
    
    For A = 0 To lstCopyAccount.ListCount - 1
        If UCase$(SelectedPj) = lstCopyAccount.List(A) Then
            MsgBox "El personaje ya está agregado"
            Exit Sub
        End If
    Next A
    
    lstCopyAccount.AddItem SelectedPj
End Sub

Private Sub lstCopyAccount_dblClick()
    lstCopyAccount.RemoveItem lstCopyAccount.ListIndex
End Sub

Private Sub lstPjs_dblClick()
    WriteRequestOfferSentUser lstPjs.List(lstPjs.ListIndex)
    
    FrmInfoMaoPjs.lblName.Caption = UCase$(lstPjs.List(lstPjs.ListIndex))
End Sub
