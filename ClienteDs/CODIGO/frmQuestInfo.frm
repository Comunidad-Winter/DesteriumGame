VERSION 5.00
Begin VB.Form frmQuestInfo 
   BorderStyle     =   0  'None
   Caption         =   "Información de la misión"
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmQuestInfo.frx":0000
   ScaleHeight     =   3720
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtInfo 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   2970
      Left            =   440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   310
      Width           =   2175
   End
   Begin VB.Image CmdRechazar 
      Height          =   255
      Left            =   1560
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Image CmdAceptar 
      Height          =   255
      Left            =   120
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "frmQuestInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
10    Call Audio.PlayWave(SND_CLICK)

20        Call WriteQuestAccept
30        Unload Me
End Sub

Private Sub cmdRechazar_Click()
10    Call Audio.PlayWave(SND_CLICK)

20        Unload Me
End Sub

