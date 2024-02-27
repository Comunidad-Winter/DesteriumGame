VERSION 5.00
Begin VB.Form frmretos 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   Caption         =   "frmretos"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmRetos.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox bGold 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   250
      Left            =   405
      TabIndex        =   15
      Top             =   5235
      Width           =   2175
   End
   Begin VB.TextBox bName 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   285
      Index           =   2
      Left            =   405
      TabIndex        =   14
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox bName 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   405
      TabIndex        =   13
      Top             =   5865
      Width           =   2175
   End
   Begin VB.TextBox bName 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   275
      Index           =   0
      Left            =   405
      TabIndex        =   12
      Top             =   6585
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Index           =   0
      Left            =   405
      TabIndex        =   11
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Left            =   405
      TabIndex        =   10
      Top             =   5235
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   405
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "///////////////////////////////"
      Top             =   5865
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   405
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "///////////////////////////////"
      Top             =   6585
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Height          =   195
      Left            =   250
      TabIndex        =   7
      Top             =   760
      Value           =   -1  'True
      Width           =   195
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Height          =   195
      Left            =   1690
      TabIndex        =   6
      Top             =   760
      Width           =   195
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   310
      TabIndex        =   5
      Top             =   1420
      Value           =   2  'Grayed
      Width           =   200
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   310
      TabIndex        =   4
      Top             =   2010
      Value           =   2  'Grayed
      Width           =   200
   End
   Begin VB.CheckBox cDrop 
      Caption         =   "Check1"
      Height          =   195
      Left            =   310
      TabIndex        =   3
      Top             =   3650
      Width           =   200
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   310
      TabIndex        =   2
      Top             =   2550
      Value           =   2  'Grayed
      Width           =   200
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check1"
      Height          =   195
      Left            =   315
      TabIndex        =   1
      Top             =   3650
      Width           =   200
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   310
      TabIndex        =   0
      Top             =   3100
      Value           =   2  'Grayed
      Width           =   200
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   7020
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   420
      Left            =   1560
      TabIndex        =   16
      Top             =   6990
      Width           =   1335
   End
   Begin VB.Image CmdSend 
      Height          =   405
      Left            =   120
      Top             =   6990
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   1560
      Top             =   6990
      Width           =   1335
   End
End
Attribute VB_Name = "frmretos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSend_Click()

 
Dim sText As String
Dim i     As Long
   
For i = 0 To 2
 
    sText = sText & bName(i).Text & IIf(i <> 2, "*", vbNullString)
 
Next i

Call Protocol.WriteSendReto(sText, Val(bGold.Text), (cDrop.value <> 0))
Unload Me
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

Unload Me
End Sub

Private Sub Label5_Click()
If Me.Text1 = "" Then
ShowConsoleMsg "Requisito para retar invalido. Verifica tu oro y el de tu oponente y la condición de ambos.", 65, 190, 156, False, False
Exit Sub
End If
'//Reto 1vs1
If Check3.value = 1 And Option1.value = True Then
WriteRetos Replace(Text2(0), " ", "+"), Text1.Text, True
ElseIf Check3.value = 0 And Option1.value = True Then
WriteRetos Replace(Text2(0), " ", "+"), Text1.Text, False
End If
Unload Me
End Sub

Private Sub Option1_Click()
'1vs1
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
bName(0).ForeColor = &HE0E0E0
bName(1).ForeColor = &HE0E0E0
End Sub

Private Sub Option2_Click()
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
bName(0).ForeColor = &HFFFFFF
bName(1).ForeColor = &HFFFFFF
End Sub

Private Sub Text3_Change()
Text3.Enabled = False
End Sub

Private Sub Text4_Change()
Text4.Enabled = False
End Sub

Private Sub Text1_Change()
'Error de Letras
If Me.bGold = "" Then
Exit Sub
End If
'Cantidad
If Text1.Text > 2000000 Then
Text1.Text = 2000000
ShowConsoleMsg "La apuesta máxima es de 2.000.000 monedas de oro.", 65, 190, 156, False, False
End If
End Sub

Private Sub bGold_Change()
'Error de letras
If Me.Text1 = "" Then
Exit Sub
End If
'Cantidad
If bGold.Text > 2000000 Then
bGold.Text = 2000000
ShowConsoleMsg "La apuesta máxima es de 2.000.000 monedas de oro.", 65, 190, 156, False, False
End If
End Sub

