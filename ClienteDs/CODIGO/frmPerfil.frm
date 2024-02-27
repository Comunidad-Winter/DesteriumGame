VERSION 5.00
Begin VB.Form frmPerfil 
   BorderStyle     =   0  'None
   ClientHeight    =   8970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   LinkTopic       =   "Form4"
   Picture         =   "frmPerfil.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picInfo 
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   8580
      Picture         =   "frmPerfil.frx":41306
      ScaleHeight     =   3061.224
      ScaleMode       =   0  'User
      ScaleWidth      =   3000
      TabIndex        =   28
      Top             =   585
      Visible         =   0   'False
      Width           =   3000
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "CLERIGO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   2550
         Left            =   195
         TabIndex        =   29
         Top             =   195
         Width           =   2550
      End
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   3
      Left            =   5900
      Picture         =   "frmPerfil.frx":49527
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   30
      Top             =   2826
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   23
      Left            =   8430
      Picture         =   "frmPerfil.frx":4CAB7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   27
      Top             =   5340
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   22
      Left            =   7605
      Picture         =   "frmPerfil.frx":50047
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   26
      Top             =   5340
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   21
      Left            =   6750
      Picture         =   "frmPerfil.frx":535D7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   25
      Top             =   5340
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   20
      Left            =   5900
      Picture         =   "frmPerfil.frx":56B67
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   24
      Top             =   5340
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   19
      Left            =   5070
      Picture         =   "frmPerfil.frx":5A0F7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   23
      Top             =   5340
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   18
      Left            =   4230
      Picture         =   "frmPerfil.frx":5D687
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   22
      Top             =   5340
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   17
      Left            =   3390
      Picture         =   "frmPerfil.frx":60C17
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   21
      Top             =   5340
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   16
      Left            =   7605
      Picture         =   "frmPerfil.frx":641A7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   20
      Top             =   4485
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   15
      Left            =   6750
      Picture         =   "frmPerfil.frx":67737
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   19
      Top             =   4485
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   14
      Left            =   5900
      Picture         =   "frmPerfil.frx":6ACC7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   18
      Top             =   4485
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   13
      Left            =   5070
      Picture         =   "frmPerfil.frx":6E257
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   17
      Top             =   4485
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   12
      Left            =   4230
      Picture         =   "frmPerfil.frx":717E7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   16
      Top             =   4485
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   11
      Left            =   3390
      Picture         =   "frmPerfil.frx":74D77
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   15
      Top             =   4485
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   10
      Left            =   7590
      Picture         =   "frmPerfil.frx":78307
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   14
      Top             =   3660
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   9
      Left            =   6750
      Picture         =   "frmPerfil.frx":7B897
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   13
      Top             =   3660
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   8
      Left            =   5900
      Picture         =   "frmPerfil.frx":7EE27
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   12
      Top             =   3660
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   7
      Left            =   5070
      Picture         =   "frmPerfil.frx":823B7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   11
      Top             =   3680
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   6
      Left            =   4230
      Picture         =   "frmPerfil.frx":85947
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   10
      Top             =   3660
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   5
      Left            =   3390
      Picture         =   "frmPerfil.frx":88ED7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   9
      Top             =   3660
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   4
      Left            =   6750
      Picture         =   "frmPerfil.frx":8C467
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   8
      Top             =   2826
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   2
      Left            =   5070
      Picture         =   "frmPerfil.frx":8F9F7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   7
      Top             =   2820
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   1
      Left            =   4230
      Picture         =   "frmPerfil.frx":92F87
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   6
      Top             =   2830
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   0
      Left            =   3390
      Picture         =   "frmPerfil.frx":96517
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   5
      Top             =   2800
      Width           =   735
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   23
      Left            =   8385
      Top             =   5265
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   22
      Left            =   7605
      Top             =   5265
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   21
      Left            =   6630
      Top             =   5265
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   20
      Left            =   5850
      Top             =   5265
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   19
      Left            =   5070
      Top             =   5265
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   18
      Left            =   4290
      Top             =   5265
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   17
      Left            =   3315
      Top             =   5265
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   16
      Left            =   7605
      Top             =   4485
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   15
      Left            =   6825
      Top             =   4485
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   14
      Left            =   5850
      Top             =   4485
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   13
      Left            =   5070
      Top             =   4485
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   12
      Left            =   4290
      Top             =   4485
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   11
      Left            =   3315
      Top             =   4485
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   10
      Left            =   7605
      Top             =   3705
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   9
      Left            =   6825
      Top             =   3705
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   8
      Left            =   5850
      Top             =   3705
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   7
      Left            =   5070
      Top             =   3705
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   6
      Left            =   4290
      Top             =   3705
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   5
      Left            =   3315
      Top             =   3705
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   4
      Left            =   6825
      Top             =   2730
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   3
      Left            =   5850
      Top             =   2730
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   2
      Left            =   5070
      Top             =   2730
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   1
      Left            =   4290
      Top             =   2730
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   0
      Left            =   3315
      Top             =   2730
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   9165
      Top             =   7800
      Width           =   2160
   End
   Begin VB.Label lblElv 
      BackStyle       =   0  'Transparent
      Caption         =   "47"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   5745
      Width           =   615
   End
   Begin VB.Label lblRaza 
      BackStyle       =   0  'Transparent
      Caption         =   "HUMANO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   5300
      Width           =   1815
   End
   Begin VB.Label lblClase 
      BackStyle       =   0  'Transparent
      Caption         =   "CLERIGO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblClan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<MI GUILD ES CARLO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   3680
      TabIndex        =   1
      Top             =   1890
      Width           =   4455
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MI NOMBRE ES CARLITOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   1200
      Width           =   4455
   End
End
Attribute VB_Name = "frmPerfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
10        Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
10        picInfo.Visible = False
End Sub
Private Sub ImgInfo_MouseMove(Index As Integer, Button As Integer, Shift As _
    Integer, X As Single, Y As Single)
          
10        picInfo.Visible = True
20        picInfo.Move (PicLogro(Index).Left + X + 20), (PicLogro(Index).Top + Y + 40)
          
          
30        InfoRequest Index
End Sub
Private Sub PicLogro_MouseMove(Index As Integer, Button As Integer, Shift As _
    Integer, X As Single, Y As Single)
          
10        picInfo.Visible = True
20        picInfo.Move (PicLogro(Index).Left + X + 20), (PicLogro(Index).Top + Y + 40)
          
          
30        InfoRequest Index
End Sub
Private Sub InfoRequest(ByVal Index As Integer)

          Dim tmpStr As String
          
10        Select Case Index
              Case 0 'Usuario recien registrado
20                tmpStr = "REGISTRADO" & vbCrLf
30                tmpStr = tmpStr & _
                      "Logro otorgado como agradecimiento por jugar DesteriumAO" & vbCrLf
40                tmpStr = tmpStr & "Recompensa: 10.000 monedas de oro"
                  
50            Case 1 ' Nivel máximo
60                tmpStr = "NIVEL MÁXIMO" & vbCrLf
70                tmpStr = tmpStr & _
                      "Logro otorgado como recompensa por alcanzar el nivel máximo." & _
                      vbCrLf
80                tmpStr = tmpStr & "Recompensa: 1.000.000 monedas de oro"
              
90            Case 2 ' Fundar clan
100               tmpStr = "FUNDADOR" & vbCrLf
110               tmpStr = tmpStr & "Logro otorgado por haber fundado un clan."
              
120           Case 3 ' Oro
130               tmpStr = "USUARIO ORO" & vbCrLf
140               tmpStr = tmpStr & "Logro otorgado a los usuarios ORO."
                  
150           Case 4 ' Premium
160               tmpStr = "USUARIO PREMIUM" & vbCrLf
170               tmpStr = tmpStr & "Logro otorgado a los usuarios PREMIUM."
                  
180           Case 5 ' Usuarios matados
190               tmpStr = "100 FRAGS" & vbCrLf
200               tmpStr = tmpStr & "Logro otorgado por alcanzar los 100 Frags" & _
                      vbCrLf & vbCrLf
210               tmpStr = tmpStr & "Recompensa: 250.000 monedas de oro"
220           Case 6 ' Usuarios matados
230               tmpStr = "200 FRAGS" & vbCrLf
240               tmpStr = tmpStr & "Logro otorgado por alcanzar los 200 Frags" & _
                      vbCrLf & vbCrLf
250               tmpStr = tmpStr & "Recompensa: 500.000 monedas de oro"
260           Case 7 ' Usuarios matados
270               tmpStr = "400 FRAGS" & vbCrLf
280               tmpStr = tmpStr & "Logro otorgado por alcanzar los 400 Frags" & _
                      vbCrLf & vbCrLf
290               tmpStr = tmpStr & "Recompensa: 750.000 monedas de oro"
300           Case 8 ' Usuarios matados
310               tmpStr = "800 FRAGS" & vbCrLf
320               tmpStr = tmpStr & "Logro otorgado por alcanzar los 800 Frags" & _
                      vbCrLf & vbCrLf
330               tmpStr = tmpStr & "Recompensa: 1.000.000 monedas de oro"
340           Case 9 ' Usuarios matados
350               tmpStr = "1600 FRAGS" & vbCrLf
360               tmpStr = tmpStr & "Logro otorgado por alcanzar los 1600 Frags" & _
                      vbCrLf & vbCrLf
370               tmpStr = tmpStr & "Recompensa: 2.000.000 monedas de oro"
380           Case 10 ' Usuarios matados
390               tmpStr = "5000 FRAGS" & vbCrLf
400               tmpStr = tmpStr & "Logro otorgado por alcanzar los 5000 Frags" & _
                      vbCrLf & vbCrLf
410               tmpStr = tmpStr & "Recompensa: 4.000.000 monedas de oro y 500 DSP"
420           Case 11 ' Retos
430               tmpStr = "5 RETOS" & vbCrLf
440               tmpStr = tmpStr & "Logro otorgado por alcanzar los 5 Retos GANADOS" _
                      & vbCrLf & vbCrLf
450               tmpStr = tmpStr & "Recompensa: 150.000 monedas de oro "
460           Case 12 ' Retos
470               tmpStr = "10 RETOS" & vbCrLf
480               tmpStr = tmpStr & "Logro otorgado por alcanzar los 10 Retos GANADOS" _
                      & vbCrLf & vbCrLf
490               tmpStr = tmpStr & "Recompensa: 300.000 monedas de oro "
500           Case 13 ' Retos
510               tmpStr = "50 RETOS" & vbCrLf
520               tmpStr = tmpStr & "Logro otorgado por alcanzar los 50 Retos GANADOS" _
                      & vbCrLf & vbCrLf
530               tmpStr = tmpStr & "Recompensa: 500.000 monedas de oro "
540           Case 14 ' Retos
550               tmpStr = "100 RETOS" & vbCrLf
560               tmpStr = tmpStr & _
                      "Logro otorgado por alcanzar los 100 Retos GANADOS" & vbCrLf & _
                      vbCrLf
570               tmpStr = tmpStr & "Recompensa: 600.000 monedas de oro "
580           Case 15 ' Retos
590               tmpStr = "250 RETOS" & vbCrLf
600               tmpStr = tmpStr & _
                      "Logro otorgado por alcanzar los 250 Retos GANADOS" & vbCrLf & _
                      vbCrLf
610               tmpStr = tmpStr & "Recompensa: 1.000.000 monedas de oro "
620           Case 16 ' Retos
630               tmpStr = "1000 RETOS" & vbCrLf
640               tmpStr = tmpStr & _
                      "Logro otorgado por alcanzar los 1000 Retos GANADOS" & vbCrLf & _
                      vbCrLf
650               tmpStr = tmpStr & "Recompensa: 3.500.000 monedas de oro "
660           Case 17 ' Eventos
670               tmpStr = "5 EVENTOS" & vbCrLf
680               tmpStr = tmpStr & _
                      "Logro otorgado por alcanzar los 5 Eventos GANADOS" & vbCrLf & _
                      vbCrLf
690               tmpStr = tmpStr & "Recompensa: 300.000 monedas de oro "
700           Case 18 ' Eventos
710               tmpStr = "10 EVENTOS" & vbCrLf
720               tmpStr = tmpStr & _
                      "Logro otorgado por alcanzar los 10 Eventos GANADOS" & vbCrLf & _
                      vbCrLf
730               tmpStr = tmpStr & "Recompensa: 500.000 monedas de oro "
740           Case 19 ' Eventos
750               tmpStr = "20 EVENTOS" & vbCrLf
760               tmpStr = tmpStr & _
                      "Logro otorgado por alcanzar los 20 Eventos GANADOS" & vbCrLf & _
                      vbCrLf
770               tmpStr = tmpStr & "Recompensa: 600.000 monedas de oro "
780           Case 20 ' Eventos
790               tmpStr = "30 EVENTOS" & vbCrLf
800               tmpStr = tmpStr & _
                      "Logro otorgado por alcanzar los 30 Eventos GANADOS" & vbCrLf & _
                      vbCrLf
810               tmpStr = tmpStr & "Recompensa: 700.000 monedas de oro "
820           Case 21 ' Eventos
830               tmpStr = "40 EVENTOS" & vbCrLf
840               tmpStr = tmpStr & _
                      "Logro otorgado por alcanzar los 40 Eventos GANADOS" & vbCrLf & _
                      vbCrLf
850               tmpStr = tmpStr & "Recompensa: 800.000 monedas de oro "
860           Case 22 ' Eventos
870               tmpStr = "50 EVENTOS" & vbCrLf
880               tmpStr = tmpStr & _
                      "Logro otorgado por alcanzar los 50 Eventos GANADOS" & vbCrLf & _
                      vbCrLf
890               tmpStr = tmpStr & "Recompensa: 900.000 monedas de oro "
900           Case 23 ' Eventos
910               tmpStr = "100 EVENTOS" & vbCrLf
920               tmpStr = tmpStr & _
                      "Logro otorgado por alcanzar los 100 Eventos GANADOS" & vbCrLf & _
                      vbCrLf
930               tmpStr = tmpStr & "Recompensa: 1.000.000 monedas de oro "
940       End Select
          
          
950       lblInfo.Caption = tmpStr
          
End Sub
'ahora vas a entenderlo dame 2m OKIS
Private Sub picInfo_Click()

End Sub

