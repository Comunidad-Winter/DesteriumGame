VERSION 5.00
Begin VB.Form frmRetos 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   7515
   ClientLeft      =   0
   ClientTop       =   30
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmRetos.frx":0000
   ScaleHeight     =   7515
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check5 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   310
      TabIndex        =   17
      Top             =   3130
      Value           =   2  'Grayed
      Width           =   200
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   400
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "///////////////////////////////"
      Top             =   6600
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   400
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "///////////////////////////////"
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Left            =   400
      TabIndex        =   14
      Top             =   5240
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Index           =   0
      Left            =   400
      TabIndex        =   13
      Top             =   4570
      Width           =   2175
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check1"
      Height          =   195
      Left            =   310
      TabIndex        =   12
      Top             =   3680
      Width           =   200
   End
   Begin VB.TextBox bName 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   275
      Index           =   0
      Left            =   400
      TabIndex        =   11
      Top             =   6600
      Width           =   2175
   End
   Begin VB.TextBox bName 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   400
      TabIndex        =   10
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox bName 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   285
      Index           =   2
      Left            =   400
      TabIndex        =   9
      Top             =   4570
      Width           =   2175
   End
   Begin VB.TextBox bGold 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   250
      Left            =   400
      TabIndex        =   8
      Top             =   5240
      Width           =   2175
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   310
      TabIndex        =   7
      Top             =   2570
      Value           =   2  'Grayed
      Width           =   200
   End
   Begin VB.CheckBox cDrop 
      Caption         =   "Check1"
      Height          =   195
      Left            =   310
      TabIndex        =   6
      Top             =   3680
      Width           =   200
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   310
      TabIndex        =   5
      Top             =   2040
      Value           =   2  'Grayed
      Width           =   200
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   310
      TabIndex        =   4
      Top             =   1440
      Value           =   2  'Grayed
      Width           =   200
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Height          =   195
      Left            =   1740
      TabIndex        =   3
      Top             =   780
      Width           =   195
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Height          =   195
      Left            =   340
      TabIndex        =   2
      Top             =   780
      Value           =   -1  'True
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   1560
      Top             =   7005
      Width           =   1335
   End
   Begin VB.Image CmdSend 
      Height          =   400
      Left            =   120
      Top             =   7000
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   420
      Left            =   1560
      TabIndex        =   1
      Top             =   7005
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   7035
      Width           =   1335
   End
End
Attribute VB_Name = "frmRetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bGold_Change()
'Error de letras
If bGold.Text = "" Then
Exit Sub
End If
'Cantidad
If Val(bGold.Text) > 2000000 Then
bGold.Text = 2000000
'ShowConsoleMsg "La apuesta máxima es de 2.000.000 monedas de oro.", 65, 190, 156, False, False
Else
If Val(bGold.Text) < 50000 Then
'ShowConsoleMsg "La apuesta mínima es de 50.000 monedas de oro.", 65, 190, 156, False, False
bGold.Text = 50000
End If
End If
End Sub

Private Sub CmdSend_Click()
Call Audio.PlayWave(SND_CLICK)
 
Dim sText As String
Dim i     As Long
   
For i = 0 To 2
 
    sText = sText & bName(i).Text & IIf(i <> 2, "*", vbNullString)
 
Next i
   
Call Protocol.WriteSendReto(sText, Val(bGold.Text), (cDrop.value <> 0))
Unload Me
End Sub

Private Sub Check3_Click()
If frmRetos.Check3.value = vbChecked Then
MsgBox "No disponible"
Check3.value = vbUnchecked
End If
End Sub

Private Sub Form_Load()
'System invisible desde entrada asi queda chevere
'System 2vs2 Invisible
bName(0).Visible = False
bName(1).Visible = False
bName(2).Visible = False
bGold.Visible = False
cDrop.Visible = False
'Label5.Visible = False
CmdSend.Visible = False
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label2_Click()
Call Audio.PlayWave(SND_CLICK)
Unload Me
End Sub

Private Sub Label5_Click()
Call Audio.PlayWave(SND_CLICK)

'Mira si mando 5000e en el server explota todo.
'Ayer vi una cosa que es que explotaba el server creo.
'//Reto 1vs1
If Check3.value = 1 And Option1.value = True Then
WriteRetos Replace(Text2(0), " ", "+"), Val(Text1.Text), True
ElseIf Check3.value = 0 And Option1.value = True Then
WriteRetos Replace(Text2(0), " ", "+"), Val(Text1.Text), False
End If
Unload Me
End Sub

Private Sub Option1_Click()
'1vs1
Call Audio.PlayWave(SND_CLICK)
Text4.Visible = True
Text3.Visible = True
Text1.Visible = True
Text2(0).Visible = True
Check3.Visible = True
Label5.Visible = True
Image1.Visible = True
Option1.Visible = True
'2vs2
bName(0).Visible = False
bName(1).Visible = False
bName(2).Visible = False
bGold.Visible = False
cDrop.Visible = False
CmdSend.Visible = False
End Sub

Private Sub Option2_Click()
Call Audio.PlayWave(SND_CLICK)
Text4.Visible = False
Text3.Visible = False
Text1.Visible = False
Text2(0).Visible = False
Check3.Visible = False
Label5.Visible = False
Image1.Visible = False
'2vs2
bName(0).Visible = True
bName(1).Visible = True
bName(2).Visible = True
bGold.Visible = True
cDrop.Visible = True
CmdSend.Visible = True
End Sub

Private Sub Text1_Change()
'Error de Letras
If Me.bGold.Text = "" Then
Exit Sub
End If
'Cantidad
If Val(Text1.Text) > 2000000 Then
Text1.Text = 2000000
ShowConsoleMsg "La apuesta máxima es de 2.000.000 monedas de oro.", 65, 190, 156, False, False
End If
End Sub

Private Sub Text3_Change()
Text3.Enabled = False
End Sub

Private Sub Text4_Change()
Text4.Enabled = False
End Sub

