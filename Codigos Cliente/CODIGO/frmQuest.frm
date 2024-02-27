VERSION 5.00
Begin VB.Form frmQuest 
   BackColor       =   &H80000008&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Quest (beta)"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4515
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "¡ ACEPTAR !"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2775
      Left            =   2160
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Recompensa: 300.000 monedas y 1.000.000 de exp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   2280
      TabIndex        =   7
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Debes asesinar 33 personas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Debes matar : 33 lobos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre: Guerra de lobos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Quest actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de quest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "frmQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Debug.Print List1.ListIndex
WriteQuestAcepta List1.ListIndex + 2
End Sub

Private Sub Form_Load()
Dim i As Long
Dim NumQ As Byte
NumQ = 2
For i = 1 To NumQ
List1.AddItem "Quest" & i
Next i
List1.ListIndex = 0
Call List1_Click
End Sub

Private Sub List1_Click()
'Label4.Caption = "Debes matar " & quest(List1.ListIndex + 1).CantidadNpcs & " " & quest(List1.ListIndex + 1).NpcNamees & "s"
'Label5.Caption = "Debes asesinar " & quest(List1.ListIndex + 1).CantidadUsers
'Label6.Caption = quest(List1.ListIndex + 1).Recompense
End Sub
