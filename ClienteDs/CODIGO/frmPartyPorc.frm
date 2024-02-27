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
Private Sub Boton_Click(index As Integer)
Call Audio.PlayWave(SND_CLICK)
Select Case index
    Case 0
    Unload Me
    Case 1
 Dim fin1$
 Dim lin1$
 Dim loopX As Long

For loopX = 0 To 4
If frmPartyPorc.Porc(loopX).Text <> "%" Then
    fin1 = fin1 & frmPartyPorc.Pj(loopX).Caption & "*" & frmPartyPorc.Porc(loopX).Text & "*" & ","
    End If
Next loopX


writeSetPartyPorcentajes fin1
    Unload Me
End Select
End Sub

Private Sub Boton_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case index
    Case 0
    'Me.boton(index).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bCancelPartyPorcS.jpg")
    Case 1
    'Me.boton(index).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bAcceptPartyPorcS.jpg")
End Select
End Sub

Private Sub Form_Load()
Dim i As Long
For i = 0 To 4
Pj(i).Caption = frmParty.Label5(i).Caption
If frmParty.Label8(i).Caption <> vbNullString Then
Porc(i).Text = frmParty.Label8(i).Caption
End If
Next i
For PT = 0 To 4
If frmPartyPorc.Pj(PT).Caption = "Personaje1" Then
frmPartyPorc.Pj(PT).Visible = False
End If
If frmPartyPorc.Porc(PT).Text = "%" Then
frmPartyPorc.Porc(PT).Enabled = False
frmPartyPorc.Porc(PT).Visible = False
End If
Next PT
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.Boton(0).Picture = LoadPicture("")
Me.Boton(1).Picture = LoadPicture("")
End Sub
