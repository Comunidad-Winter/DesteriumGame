VERSION 5.00
Begin VB.Form frmParty 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Partym.frx":0000
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   4575
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Experiencia total:"
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
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "10.000.000"
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
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   4560
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   4560
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   4560
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   4560
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   4560
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje"
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
         Height          =   255
         Left            =   2880
         TabIndex        =   23
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Experiencia"
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
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   4
         Left            =   3120
         TabIndex        =   20
         Top             =   1920
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   3
         Left            =   3120
         TabIndex        =   19
         Top             =   1560
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   18
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   1
         Left            =   3120
         TabIndex        =   17
         Top             =   840
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   4
         Left            =   1560
         TabIndex        =   16
         Top             =   1920
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   3
         Left            =   1560
         TabIndex        =   15
         Top             =   1560
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   2
         Left            =   1560
         TabIndex        =   14
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   13
         Top             =   840
         Width           =   45
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   8
         Top             =   480
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   0
         Left            =   3120
         TabIndex        =   7
         Top             =   480
         Width           =   45
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   2175
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   2175
      ItemData        =   "Partym.frx":CACD
      Left            =   2760
      List            =   "Partym.frx":CACF
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Experiencia total:"
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
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Image boton 
      Height          =   255
      Index           =   0
      Left            =   3720
      Top             =   7500
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image boton 
      Height          =   480
      Index           =   3
      Left            =   3480
      Picture         =   "Partym.frx":CAD1
      Top             =   3360
      Width           =   1650
   End
   Begin VB.Image boton 
      Height          =   480
      Index           =   7
      Left            =   1800
      Picture         =   "Partym.frx":10B5C
      Top             =   3360
      Width           =   1650
   End
   Begin VB.Image boton 
      Height          =   480
      Index           =   6
      Left            =   120
      Picture         =   "Partym.frx":14BA8
      Top             =   3360
      Width           =   1650
   End
   Begin VB.Image boton 
      Height          =   480
      Index           =   5
      Left            =   2760
      Picture         =   "Partym.frx":18CA0
      Top             =   3840
      Width           =   2325
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Height          =   255
      Left            =   4920
      TabIndex        =   25
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitudes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Integrantes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image boton 
      Height          =   480
      Index           =   2
      Left            =   120
      Picture         =   "Partym.frx":1D110
      Top             =   3840
      Width           =   2325
   End
   Begin VB.Image boton 
      Height          =   480
      Index           =   1
      Left            =   120
      Picture         =   "Partym.frx":2133E
      Top             =   3840
      Width           =   2325
   End
End
Attribute VB_Name = "frmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function prepararString() As String
    
    Dim J    As Long
    Dim tStr As String
    
    For J = 0 To 4
        If (tStr = vbNullString) Then
            tStr = Label5(J).Caption & "*" & Label8(J).Caption
        Else
            tStr = tStr & "," & Label5(J).Caption & "*" & Label7(J).Caption
        End If
    Next J
    
    prepararString = tStr

End Function

Public Sub prepararForm(ByRef sourceString As String)
   
    Dim loopC   As Long
    Dim temp()  As String
    Dim nUser   As Byte
    Dim tmpName As String

 '   If Not InStr(1, sourceString, ",") Then
  '      Label5(0).Caption = ReadField(1, sourceString, Asc("*"))
   '     Label7(0).Caption = ReadField(2, sourceString, Asc("*"))
    '    Label8(0).Caption = ReadField(3, sourceString, Asc("*"))
     '   Exit Sub
   ' End If
    
    temp() = Split(sourceString, ",")
    
    For loopC = 0 To UBound(temp())
        If Not temp(loopC) = vbNullString Then
           tmpName = ReadField(1, temp(loopC), Asc("*"))
           
           If Not tmpName = vbNullString Then
              Label5(nUser).Caption = tmpName
              Label7(nUser).Caption = ReadField(2, temp(loopC), Asc("*"))
              Label8(nUser).Caption = ReadField(3, temp(loopC), Asc("*"))
              
              nUser = nUser + 1
           End If
        End If
    Next loopC
    
End Sub



Private Sub Boton_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)
Dim i As Long
Select Case Index
    Case 1
    Me.Boton(1).Visible = False
    Me.Boton(2).Visible = True
    Me.Boton(3).Enabled = False
    Me.Boton(5).Enabled = False
    Me.Boton(6).Enabled = False
    Me.Boton(7).Enabled = False
    Case 2
    Me.Boton(2).Visible = False
    Me.Boton(1).Visible = True
    Me.Boton(3).Enabled = False
    Me.Boton(5).Enabled = False
    Me.Boton(6).Enabled = False
    Me.Boton(7).Enabled = False
    WritePartyLeave
    Unload Me
    Case 3
    Case 5
    frmPartyPorc.Show , frmParty
    Case 6
    For i = 0 To (List2.ListCount - 1)
    If List2.List(i) <> vbNullString Then
    WritePartyKick List1.List(i)
    End If
    Next i
    List2.Clear
    Case 7
    Me.Boton(7).Enabled = False
    Me.Boton(3).Enabled = False
    For i = 0 To (List1.ListCount - 1)
    If List1.List(i) <> vbNullString Then
    WritePartyAcceptMember List1.List(i)
    List2.AddItem List1.List(i)
    End If
    Next i
    List1.Clear
End Select
End Sub

Private Sub Boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Index
    Case 1
    Boton(Index).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bCreatePartyS.jpg")
    Case 2
    Boton(Index).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bQuitPartyS.jpg")
    Case 3
    If Boton(Index).Enabled = True Then
    Boton(Index).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bRejectPartyS.jpg")
    'Else
    'boton(Index).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bRejectPartyN.jpg")
    End If
    Case 5
    If Boton(Index).Enabled = True Then
    Boton(Index).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bChangePorcS.jpg")
    'Else
    'boton(Index).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bChangePorcN.jpg")
    End If
    Case 6
    If Boton(Index).Enabled = True Then
    Boton(Index).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bRemovePartyS.jpg")
    'else
    'boton(Index).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bRemovePartyN.jpg")
    End If
    Case 7
    If Boton(Index).Enabled = True Then
    Boton(Index).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bAcceptPartyS.jpg")
    'else
    'boton(Index).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bAcceptPartyN.jpg")
    End If
End Select
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Boton(1).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bCreateParty.jpg")
Boton(2).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bQuitParty.jpg")
If Boton(3).Enabled = True Then
Boton(3).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bRejectParty.jpg")
Else
Boton(3).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bRejectPartyN.jpg")
End If
If Boton(5).Enabled = True Then
Boton(5).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bChangePorc.jpg")
Else
Boton(5).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bChangePorcN.jpg")
End If
If Boton(6).Enabled = True Then
Boton(6).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bRemoveParty.jpg")
Else
Boton(6).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bRemovePartyN.jpg")
End If
If Boton(7).Enabled = True Then
Boton(7).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bAcceptParty.jpg")
Else
Boton(7).Picture = LoadPicture(App.path & "\Recursos\Button\Party\bAcceptPartyN.jpg")
End If
End Sub

Private Sub Label1_Click()

    Call Unload(Me)
    
    Call frmMain.SetFocus

End Sub


Private Sub Label6_Click()

    If (Label6.Caption = ">>") Then
        Label6.Caption = "<<"
        List1.Visible = True
        List2.Visible = True
        Label2.Visible = True
        Label3.Visible = True
        Frame1.Visible = False
    Else
        Label6.Caption = ">>"
        List1.Visible = False
        List2.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Frame1.Visible = True
    End If
    
End Sub
