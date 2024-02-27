VERSION 5.00
Begin VB.Form FrmCanje 
   BorderStyle     =   0  'None
   Caption         =   "Canjes Desterium"
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8040
   LinkTopic       =   "Form4"
   Picture         =   "FrmCanje.frx":0000
   ScaleHeight     =   6210
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicCanje 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   780
      ScaleHeight     =   2400
      ScaleWidth      =   3870
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1365
      Width           =   3870
   End
   Begin VB.Image imgCanje 
      Height          =   405
      Left            =   1560
      Top             =   3900
      Width           =   2160
   End
   Begin VB.Label lblRequired 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   795
      Left            =   4875
      TabIndex        =   6
      Top             =   4290
      Width           =   2940
   End
   Begin VB.Label lblPuntos 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   210
      Left            =   6630
      TabIndex        =   5
      Top             =   3260
      Width           =   1380
   End
   Begin VB.Label lblSeCae 
      BackStyle       =   0  'Transparent
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
      Height          =   210
      Left            =   6435
      TabIndex        =   4
      Top             =   2900
      Width           =   1380
   End
   Begin VB.Label lblAtaqueFisico 
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   210
      Left            =   6435
      TabIndex        =   3
      Top             =   2700
      Width           =   1380
   End
   Begin VB.Label lblRM 
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   210
      Left            =   6435
      TabIndex        =   2
      Top             =   2310
      Width           =   1380
   End
   Begin VB.Label lblDef 
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   210
      Left            =   6435
      TabIndex        =   1
      Top             =   2130
      Width           =   1380
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   6825
      Top             =   390
      Width           =   795
   End
End
Attribute VB_Name = "FrmCanje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Image1_Click()
10        HandleCanjeEnd
End Sub

Private Sub imgCanje_Click()
10        If InvCanje.SelectedItem = 0 Then Exit Sub
          
          If MsgBox("¿Estás seguro que deseas comprar este objeto?", vbYesNo) = vbYes Then
20          Call _
                Protocol.WriteCanjeItem(Canjes(InvCanje.SelectedItem).ObjCanje.OBJIndex, _
                Canjes(InvCanje.SelectedItem).ObjRequired(1).OBJIndex, _
                Canjes(InvCanje.SelectedItem).Points)
                
                HandleCanjeEnd
40          Unload Me
          End If
          
30
          
End Sub

Private Sub PicCanje_Click()
10        If InvCanje.SelectedItem = 0 Then Exit Sub
20        If InvCanje.SelectedItem > NumCanjes Then Exit Sub
30        Call _
              Protocol.WriteCanjeInfo(Canjes(InvCanje.SelectedItem).ObjCanje.OBJIndex, _
              Canjes(InvCanje.SelectedItem).ObjRequired(1).OBJIndex, _
              Canjes(InvCanje.SelectedItem).Points)
End Sub
