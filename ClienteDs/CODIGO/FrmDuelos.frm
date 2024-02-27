VERSION 5.00
Begin VB.Form FrmDuelos 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   8235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8265
   LinkTopic       =   "Form4"
   Picture         =   "FrmDuelos.frx":0000
   ScaleHeight     =   8235
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCompa 
      Caption         =   "Check1"
      Height          =   210
      Left            =   2925
      TabIndex        =   10
      Top             =   3900
      Width           =   210
   End
   Begin VB.TextBox txtLider 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   210
      Left            =   1755
      TabIndex        =   9
      Top             =   6550
      Width           =   1185
   End
   Begin VB.TextBox txtEnemy 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   210
      Left            =   1560
      TabIndex        =   8
      Top             =   4290
      Width           =   1185
   End
   Begin VB.TextBox txtTeam 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   210
      Left            =   1560
      TabIndex        =   7
      Top             =   3900
      Width           =   1185
   End
   Begin VB.TextBox txtRojas 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   210
      Left            =   1170
      TabIndex        =   6
      Text            =   "0"
      Top             =   3250
      Width           =   1185
   End
   Begin VB.TextBox txtOro 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   210
      Left            =   1170
      TabIndex        =   5
      Text            =   "0"
      Top             =   3000
      Width           =   1185
   End
   Begin VB.TextBox txtDsp 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   210
      Left            =   1170
      TabIndex        =   4
      Text            =   "0"
      Top             =   2730
      Width           =   1185
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   6825
      Top             =   7605
      Width           =   1185
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   2535
      Top             =   6825
      Width           =   1380
   End
   Begin VB.Image imgSendSolicitud 
      Height          =   405
      Left            =   2340
      Top             =   4680
      Width           =   1380
   End
   Begin VB.Label lblClan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<CLANTAGNAME>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   210
      Left            =   5175
      TabIndex        =   3
      Top             =   4980
      Width           =   1770
   End
   Begin VB.Label lblTercer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<TERCER PUESTO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   795
      Left            =   3960
      TabIndex        =   2
      Top             =   3510
      Width           =   4110
   End
   Begin VB.Label lblSegundo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<SEGUNDO PUESTO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   795
      Left            =   4020
      TabIndex        =   1
      Top             =   2700
      Width           =   4110
   End
   Begin VB.Label lblPrimer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<PRIMER PUESTO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   795
      Left            =   4040
      TabIndex        =   0
      Top             =   1680
      Width           =   4110
   End
End
Attribute VB_Name = "FrmDuelos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
10        If txtLider.Text = vbNullString Then
20            MsgBox "Elije el nombre del lider del clan al que quieras enfrentar."
30            Exit Sub
40        End If
          
50        Call Protocol.WriteSendFightClan(txtLider.Text)
          
End Sub

Private Sub Image2_Click()
10        Unload Me
End Sub

Private Sub imgSendSolicitud_Click()
          Dim lstTeam() As String
          Dim lstOponent() As String
          Dim Team(20) As Byte
          Dim strTemp As String
          Dim tmpTeam As Byte
          
10        If Val(txtDsp.Text) > 30000 Then
20            MsgBox "No puedes elegir más de 30.000 DSP"
30            Exit Sub
40        End If
          
50        If Val(txtOro.Text) > 50000000 Then
60            MsgBox "No puedes elegir más de 50.000.000 monedas de oro."
70            Exit Sub
80        End If
          
90        If chkCompa.value = 1 Then
100           If txtTeam.Text = vbNullString Then
110               MsgBox "Selecciona a tu/s compañero/s"
120               Exit Sub
130           End If
140       End If
          
150       If txtEnemy.Text = vbNullString Then
160           MsgBox "Debes elegir a tu/s enemigo/s"
170           Exit Sub
180       End If
          
          
          ' 1vs1
190       If chkCompa.value = 0 Then
200           If InStr(UCase$(txtEnemy.Text), UCase$("-")) > 0 Then
210               MsgBox _
                      "Si lo que buscas es realizar un enfrentamiento 1vs1 no hace falta poner el '-'"
220               Exit Sub
230           End If
240           strTemp = txtEnemy.Text
250       Else
              
260           lstTeam = Split(txtTeam.Text, "-")
270           lstOponent = Split(txtEnemy.Text, "-")
              
280           If UBound(lstOponent) > 2 Then
290               MsgBox "El máximo de luchas es de 3vs3"
300               Exit Sub
310           End If
320       End If
          

          
330       If chkCompa.value = 1 Then
340           If UBound(lstTeam) + 1 <> UBound(lstOponent) Then
350               MsgBox _
                      "Recuerda y ten encuenta que por cada compañero que elijas, deberás poner un enemigo más."
360               Exit Sub
370           End If
              
              ' ENEMIGO
380           For LoopC = LBound(lstOponent) To UBound(lstOponent)
390               strTemp = strTemp & lstOponent(LoopC) & "-"
400           Next LoopC
              
              ' TEAM
410           For LoopC = LBound(lstTeam) To UBound(lstTeam)
420               strTemp = strTemp & lstTeam(LoopC) & "-"
430           Next LoopC
              
440           strTemp = Left(strTemp, Len(strTemp) - 1)
450       End If

460       Protocol.WriteSendFight Val(txtOro.Text), Val(txtDsp.Text), _
              Val(txtRojas.Text), strTemp
470       Unload Me
End Sub

