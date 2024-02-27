VERSION 5.00
Begin VB.Form frmPartyPorc 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "Acomodar Porcentajes"
   ClientHeight    =   2985
   ClientLeft      =   4305
   ClientTop       =   3105
   ClientWidth     =   3000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPartyPorc.frx":0000
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Porc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   2280
      TabIndex        =   9
      Text            =   "0"
      Top             =   2010
      Width           =   375
   End
   Begin VB.TextBox Porc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   8
      Text            =   "0"
      Top             =   1650
      Width           =   375
   End
   Begin VB.TextBox Porc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Porc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   5
      Text            =   "0"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Porc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   7
      Text            =   "0"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image Boton 
      Height          =   375
      Index           =   1
      Left            =   150
      Top             =   2490
      Width           =   975
   End
   Begin VB.Image Boton 
      Height          =   375
      Index           =   0
      Left            =   1800
      Top             =   2520
      Width           =   975
   End
   Begin VB.Line Lin 
      BorderColor     =   &H00E0E0E0&
      Index           =   5
      Visible         =   0   'False
      X1              =   120
      X2              =   3120
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Lin 
      BorderColor     =   &H00E0E0E0&
      Index           =   4
      Visible         =   0   'False
      X1              =   120
      X2              =   3120
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Lin 
      BorderColor     =   &H00E0E0E0&
      Index           =   3
      Visible         =   0   'False
      X1              =   120
      X2              =   3120
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Lin 
      BorderColor     =   &H00E0E0E0&
      Index           =   2
      Visible         =   0   'False
      X1              =   120
      X2              =   3120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Lin 
      BorderColor     =   &H00E0E0E0&
      Index           =   1
      Visible         =   0   'False
      X1              =   120
      X2              =   3120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2400
      TabIndex        =   11
      Top             =   360
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje"
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
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Pj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pj1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   285
   End
   Begin VB.Label Pj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pj1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   285
   End
   Begin VB.Label Pj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pj1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   285
   End
   Begin VB.Label Pj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pj1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   285
   End
   Begin VB.Label Pj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pj1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   285
   End
End
Attribute VB_Name = "frmPartyPorc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Boton_Click(Index As Integer)
10    Call Audio.PlayWave(SND_CLICK)
20    Select Case Index
          Case 0
30        Unload Me
40        Case 1
       Dim fin1$
       Dim lin1$
       Dim loopX As Long

50    For loopX = 0 To 4
60    If frmPartyPorc.Porc(loopX).Text <> "%" Then
70        fin1 = fin1 & frmPartyPorc.Pj(loopX).Caption & "*" & _
              frmPartyPorc.Porc(loopX).Text & "*" & ","
80        End If
90    Next loopX


100   writeSetPartyPorcentajes fin1
110       Unload Me
120   End Select
End Sub

Private Sub Boton_MouseMove(Index As Integer, Button As Integer, Shift As _
    Integer, X As Single, Y As Single)
10    Select Case Index
          Case 0
          'Me.boton(index).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bCancelPartyPorcS.jpg")
20        Case 1
          'Me.boton(index).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bAcceptPartyPorcS.jpg")
30    End Select
End Sub

Private Sub Form_Load()
      Dim i As Long
10    For i = 0 To 4
20    Pj(i).Caption = frmParty.Label5(i).Caption
30    If frmParty.Label8(i).Caption <> vbNullString Then
40    Porc(i).Text = frmParty.Label8(i).Caption
50    End If
60    Next i
70    For PT = 0 To 4
80    If frmPartyPorc.Pj(PT).Caption = "Personaje1" Then
90    frmPartyPorc.Pj(PT).Visible = False
100   End If
110   If frmPartyPorc.Porc(PT).Text = "%" Then
120   frmPartyPorc.Porc(PT).Enabled = False
130   frmPartyPorc.Porc(PT).Visible = False
140   End If
150   Next PT
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
10    Me.boton(0).Picture = LoadPicture("")
20    Me.boton(1).Picture = LoadPicture("")
End Sub
