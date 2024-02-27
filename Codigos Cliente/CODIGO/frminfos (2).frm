VERSION 5.00
Begin VB.Form frmretos 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   Caption         =   "frmretos"
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7365
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox bName 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   275
      Index           =   0
      Left            =   720
      TabIndex        =   15
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox bName 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   14
      Top             =   5160
      Width           =   2175
   End
   Begin VB.TextBox bName 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   285
      Index           =   2
      Left            =   720
      TabIndex        =   13
      Top             =   4455
      Width           =   2175
   End
   Begin VB.TextBox bGold 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   250
      Left            =   720
      TabIndex        =   12
      Top             =   3700
      Width           =   2175
   End
   Begin VB.CheckBox cDrop 
      Caption         =   "Check1"
      Height          =   195
      Left            =   600
      TabIndex        =   11
      Top             =   2870
      Width           =   200
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Height          =   195
      Left            =   1920
      TabIndex        =   10
      Top             =   960
      Width           =   195
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Height          =   195
      Left            =   600
      TabIndex        =   9
      Top             =   960
      Value           =   -1  'True
      Width           =   195
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "///////////////////////////////"
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "///////////////////////////////"
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   600
      TabIndex        =   6
      Top             =   1920
      Value           =   2  'Grayed
      Width           =   200
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Value           =   2  'Grayed
      Width           =   200
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   2400
      Value           =   2  'Grayed
      Width           =   200
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Left            =   720
      TabIndex        =   3
      Top             =   3710
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Top             =   4490
      Width           =   2175
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check1"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   2870
      Width           =   200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   1920
      TabIndex        =   16
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Image CmdSend 
      Height          =   615
      Left            =   120
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1920
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Width           =   1695
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

Private Sub Label5_Click()
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
End Sub

Private Sub Text3_Change()
Text3.Enabled = False
End Sub

Private Sub Text4_Change()
Text4.Enabled = False
End Sub
