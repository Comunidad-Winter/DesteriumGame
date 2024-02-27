VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   ClientHeight    =   8970
   ClientLeft      =   360
   ClientTop       =   300
   ClientWidth     =   17010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmMain.frx":1CCA
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1134
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   6480
      Top             =   2880
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   10240
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   10000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   8880
      ScaleHeight     =   162.727
      ScaleMode       =   0  'User
      ScaleWidth      =   163.012
      TabIndex        =   83
      Top             =   2475
      Width           =   2415
   End
   Begin VB.TextBox SendRmstxt 
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   90
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1980
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.TextBox SendGms 
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   90
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   71
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1980
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      CausesValidation=   0   'False
      Height          =   195
      Left            =   13320
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   3000
      Width           =   180
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   90
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1980
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.TextBox SendCMSTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1980
      Visible         =   0   'False
      Width           =   8175
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7560
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      _Version        =   393216
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   3
      Left            =   7800
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   10
      Top             =   9240
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   2
      Left            =   7320
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   9
      Top             =   9240
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   6840
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   8
      Top             =   9240
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   0
      Left            =   5880
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   7
      Top             =   9360
      Visible         =   0   'False
      Width           =   420
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   5760
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   5640
      Top             =   2520
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   5160
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5760
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   4680
      Top             =   2520
   End
   Begin VB.ListBox hlst 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2790
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":2C2D3
      Left            =   8610
      List            =   "frmMain.frx":2C2D5
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2355
      Visible         =   0   'False
      Width           =   2595
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1470
      Left            =   75
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   420
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   2593
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":2C2D7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox MainViewPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6210
      Left            =   240
      MousePointer    =   99  'Custom
      ScaleHeight     =   414
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   542
      TabIndex        =   20
      Top             =   2280
      Width           =   8130
      Begin VB.Timer TimerF8 
         Interval        =   30000
         Left            =   2730
         Top             =   390
      End
      Begin VB.Timer Timer1 
         Interval        =   40
         Left            =   3900
         Top             =   1755
      End
      Begin VB.Timer TimerPing 
         Interval        =   1010
         Left            =   2340
         Top             =   1755
      End
      Begin VB.Timer tSec 
         Interval        =   60000
         Left            =   1755
         Top             =   390
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10995
      TabIndex        =   99
      Top             =   780
      Width           =   660
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   14280
      TabIndex        =   98
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   12480
      TabIndex        =   97
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11040
      TabIndex        =   55
      Top             =   975
      Width           =   540
   End
   Begin VB.Label imgAsignarSkill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
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
      Left            =   8640
      TabIndex        =   57
      Top             =   720
      Width           =   225
   End
   Begin VB.Label lblclan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Desterium AO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   8700
      TabIndex        =   66
      Top             =   990
      Width           =   2355
   End
   Begin VB.Image ImageRanking 
      Height          =   300
      Left            =   10320
      Top             =   7290
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   9960
      TabIndex        =   74
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   8400
      TabIndex        =   73
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Height          =   9015
      Left            =   -15
      TabIndex        =   70
      Top             =   -15
      Width           =   90
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Height          =   2655
      Left            =   10200
      TabIndex        =   69
      Top             =   5760
      Width           =   255
   End
   Begin VB.Image DropGold 
      Height          =   255
      Left            =   10320
      Top             =   6360
      Width           =   255
   End
   Begin VB.Image ImageQUest 
      Height          =   300
      Left            =   10320
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label lblclan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Desterium AO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   2
      Left            =   8715
      TabIndex        =   68
      Top             =   975
      Width           =   2355
   End
   Begin VB.Label lblclan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Desterium AO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   1
      Left            =   8685
      TabIndex        =   67
      Top             =   975
      Width           =   2355
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   65
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   64
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblhelm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1725
      TabIndex        =   63
      ToolTipText     =   " "
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label IconosegD 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   9840
      TabIndex        =   62
      Top             =   5505
      Width           =   435
   End
   Begin VB.Image ImagePARTY 
      Height          =   375
      Left            =   12480
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   10800
      TabIndex        =   60
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label IconoSeg 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   8985
      TabIndex        =   59
      Top             =   5505
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajo"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10800
      TabIndex        =   58
      Top             =   5760
      Width           =   735
   End
   Begin VB.Image imgEstadisticas 
      Height          =   345
      Left            =   10320
      MouseIcon       =   "frmMain.frx":2C354
      MousePointer    =   99  'Custom
      Top             =   7680
      Width           =   1170
   End
   Begin VB.Label Labelgm4 
      BackStyle       =   0  'Transparent
      Caption         =   "Panel 1"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8520
      TabIndex        =   56
      Top             =   8280
      Width           =   615
   End
   Begin VB.Image cmdInfo 
      Height          =   525
      Left            =   10440
      MouseIcon       =   "frmMain.frx":2C4A6
      MousePointer    =   99  'Custom
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image CmdLanzar 
      Height          =   615
      Left            =   8520
      MouseIcon       =   "frmMain.frx":2C5F8
      MousePointer    =   99  'Custom
      Top             =   5160
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Anstirion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   8280
      TabIndex        =   54
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Anstirion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   2
      Left            =   8295
      TabIndex        =   53
      Top             =   705
      Width           =   3255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Anstirion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   1
      Left            =   8265
      TabIndex        =   52
      Top             =   705
      Width           =   3255
   End
   Begin VB.Label lblporclvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   11100
      TabIndex        =   51
      Top             =   990
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lvllbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "47"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   11205
      TabIndex        =   50
      Top             =   990
      Width           =   225
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   10785
      TabIndex        =   49
      Top             =   6405
      Width           =   90
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   8790
      TabIndex        =   48
      Top             =   7740
      Width           =   495
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   4
      Left            =   12600
      TabIndex        =   47
      Top             =   7680
      Width           =   1350
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   3
      Left            =   12465
      TabIndex        =   46
      Top             =   7695
      Width           =   1350
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   2
      Left            =   12465
      TabIndex        =   45
      Top             =   7665
      Width           =   1350
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   1
      Left            =   12450
      TabIndex        =   44
      Tag             =   "5"
      Top             =   7680
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   9480
      TabIndex        =   43
      Top             =   7740
      Width           =   615
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   4
      Left            =   12480
      TabIndex        =   42
      Top             =   7320
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   3
      Left            =   12465
      TabIndex        =   41
      Top             =   7335
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   2
      Left            =   12465
      TabIndex        =   40
      Top             =   7305
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   1
      Left            =   12450
      TabIndex        =   39
      Top             =   7320
      Width           =   1350
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   8520
      TabIndex        =   38
      Top             =   7335
      Width           =   1470
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   3
      Left            =   13800
      TabIndex        =   37
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   2
      Left            =   12960
      TabIndex        =   36
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   1
      Left            =   13200
      TabIndex        =   35
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   8520
      TabIndex        =   34
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   4
      Left            =   13920
      TabIndex        =   33
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   3
      Left            =   13200
      TabIndex        =   32
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   2
      Left            =   13200
      TabIndex        =   31
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   1
      Left            =   15000
      TabIndex        =   30
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   8520
      TabIndex        =   29
      Top             =   6570
      Width           =   1455
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   1
      Left            =   14520
      TabIndex        =   28
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   2
      Left            =   14760
      TabIndex        =   27
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   3
      Left            =   14760
      TabIndex        =   26
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   4
      Left            =   14520
      TabIndex        =   25
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   4
      Left            =   12360
      TabIndex        =   24
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Image STAShp 
      Height          =   165
      Left            =   8535
      Picture         =   "frmMain.frx":2C74A
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Image MANShp 
      Height          =   165
      Left            =   8535
      Picture         =   "frmMain.frx":2D6EB
      Top             =   6975
      Width           =   1455
   End
   Begin VB.Image Hpshp 
      Height          =   180
      Left            =   8535
      Picture         =   "frmMain.frx":2E69E
      Top             =   7350
      Width           =   1455
   End
   Begin VB.Image COMIDAsp 
      Height          =   195
      Left            =   14280
      Picture         =   "frmMain.frx":2F632
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Image AGUAsp 
      Height          =   180
      Left            =   14520
      Picture         =   "frmMain.frx":305C2
      Top             =   7695
      Width           =   1455
   End
   Begin VB.Image CMSG 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   12600
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   1275
   End
   Begin VB.Image imgPMSG 
      Height          =   300
      Left            =   12240
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invisible"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8760
      TabIndex        =   19
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Labelgm3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cr de 5 Seg"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9720
      TabIndex        =   18
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Labelgm44 
      BackStyle       =   0  'Transparent
      Caption         =   "Panel 2"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9480
      TabIndex        =   17
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Minimizar 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8520
      TabIndex        =   16
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Cerrar 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10440
      TabIndex        =   15
      Top             =   0
      Width           =   975
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   375
      Index           =   1
      Left            =   11265
      Top             =   2430
      Width           =   195
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   255
      Index           =   0
      Left            =   11265
      Top             =   2880
      Width           =   195
   End
   Begin VB.Image imgClanes 
      Height          =   375
      Left            =   10320
      Top             =   8040
      Width           =   1275
   End
   Begin VB.Image imgOpciones 
      Height          =   330
      Left            =   10320
      Top             =   6600
      Width           =   1275
   End
   Begin VB.Label lblDropGold 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   10440
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   6360
      Width           =   135
   End
   Begin VB.Label lblMinimizar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   13200
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   180
      Width           =   255
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   13470
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   180
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10140
      MouseIcon       =   "frmMain.frx":31547
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1560
      Width           =   1485
   End
   Begin VB.Label lblFPS 
      BackStyle       =   0  'Transparent
      Caption         =   "101"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   5280
      TabIndex        =   6
      Top             =   120
      Width           =   435
   End
   Begin VB.Image xz 
      Height          =   255
      Index           =   0
      Left            =   13320
      Top             =   120
      Width           =   255
   End
   Begin VB.Image xzz 
      Height          =   195
      Index           =   1
      Left            =   13365
      Top             =   120
      Width           =   225
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      MouseIcon       =   "frmMain.frx":31699
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label lblStrg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   8925
      TabIndex        =   2
      Top             =   8160
      Width           =   210
   End
   Begin VB.Label lblDext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   9660
      TabIndex        =   1
      Top             =   8100
      Width           =   210
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa 1 [50,50]"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8955
      TabIndex        =   0
      Top             =   8640
      Width           =   2505
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   6225
      Left            =   240
      Top             =   2280
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.Image InvEqu 
      Height          =   4485
      Left            =   8340
      Picture         =   "frmMain.frx":317EB
      Top             =   1560
      Width           =   3300
   End
   Begin VB.Label lblporclvl 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   11085
      TabIndex        =   75
      Top             =   990
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblporclvl 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   11100
      TabIndex        =   76
      Top             =   975
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblporclvl 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   11100
      TabIndex        =   77
      Top             =   1005
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblporclvl 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   11115
      TabIndex        =   78
      Top             =   990
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lvllbl 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "47"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   11190
      TabIndex        =   79
      Top             =   990
      Width           =   225
   End
   Begin VB.Label lvllbl 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "47"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   11205
      TabIndex        =   80
      Top             =   975
      Width           =   225
   End
   Begin VB.Label lvllbl 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "47"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   11205
      TabIndex        =   81
      Top             =   1005
      Width           =   225
   End
   Begin VB.Label lvllbl 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "47"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   11220
      TabIndex        =   82
      Top             =   990
      Width           =   225
   End
   Begin VB.Label lblarmor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   690
      TabIndex        =   84
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblarmor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   12510
      TabIndex        =   85
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblarmor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   12540
      TabIndex        =   86
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblhelm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   13815
      TabIndex        =   87
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblhelm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   13845
      TabIndex        =   88
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   89
      ToolTipText     =   " "
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   90
      ToolTipText     =   " "
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   6720
      TabIndex        =   91
      ToolTipText     =   " "
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   6750
      TabIndex        =   92
      ToolTipText     =   " "
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblclan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Desterium AO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   3
      Left            =   8670
      TabIndex        =   93
      Top             =   990
      Width           =   2355
   End
   Begin VB.Label lblclan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Desterium AO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   4
      Left            =   8730
      TabIndex        =   94
      Top             =   990
      Width           =   2355
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Anstirion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   3
      Left            =   8250
      TabIndex        =   95
      Top             =   735
      Width           =   3255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Anstirion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   4
      Left            =   8310
      TabIndex        =   96
      Top             =   735
      Width           =   3255
   End
   Begin VB.Label lblmapaname 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Catacumbas de Ullathorpe"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9000
      TabIndex        =   14
      Top             =   8640
      Width           =   2415
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20&

'Security // Evitar F8
' // Detectar posicion del cursor.
Private Declare Function GetCursorPos Lib "user32.dll" (Pt As Point) As Long

Private Type Point
    X As Long
    Y As Long
End Type

'End Security


 'Api para generar un evento de tecla, en este caso Print Screen
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Sub keybd_event _
    Lib "user32" ( _
        ByVal bVk As Byte, _
        ByVal bScan As Byte, _
        ByVal dwFlags As Long, _
        ByVal dwExtraInfo As Long)
  

' x Auto Pots
Private Enum eVentanas

    vHechizos = 1
    vInventario = 2

End Enum

Private Panel              As Byte
Private LastPanel          As Byte
Private Const InvalidSlot  As Byte = 255
' x Auto Pots

' x button
Private mouse_Down         As Boolean
Private mouse_UP           As Boolean
' x button

Public n As Byte

Public Pulsacion_Fisica As Boolean

Private MouseInvBoton As Long

Public Attack As Boolean
Private Last_I      As Long

#If Wgl = 0 Then
    Public WithEvents dragInventory As clsGrapchicalInventory
Attribute dragInventory.VB_VarHelpID = -1
#Else
    Public WithEvents dragInventory As clsGrapchicalInventoryWgl
Attribute dragInventory.VB_VarHelpID = -1
#End If

Dim Ancho As Integer
Dim alto As Integer
Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long

Public IsPlaying As Byte

Private clsFormulario As clsFormMovementManager

Private cBotonDiamArriba As clsGraphicalButton
Private cBotonDiamAbajo As clsGraphicalButton
Private cBotonMapa As clsGraphicalButton
Private cBotonGrupo As clsGraphicalButton
Private cBotonOpciones As clsGraphicalButton
Private cBotonEstadisticas As clsGraphicalButton
Private cBotonClanes As clsGraphicalButton
Private cBotonAsignarSkill As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Public picSkillStar As Picture

Private cmsgSupr As Boolean
Private bCMSG As Boolean
Private btmpCMSG As Boolean
Private sPartyChat As String

'recibe la ruta donde crear el BMP
''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Capturar_Guardar(path As String)
      
    ' borra el portapapeles
    Clipboard.Clear
      
    ' Manda la pulsación de teclas para capturar la imagen de la pantalla
    Call keybd_event(44, 2, 0, 0)
      
    DoEvents
    ' Si el formato del clipboard es un bitmap
    If Clipboard.GetFormat(vbCFBitmap) Then
      
        'Guardamos la imagen en disco
        SavePicture Clipboard.GetData(vbCFBitmap), path
    Else
        MsgBox " Error ", vbCritical
    End If
  
End Sub
Private Sub Cerrar_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    If MsgBox("¿Desea cerrar Desterium AO?", vbYesNo + vbQuestion, "Desterium AO") _
          = vbYes Then
30      prgRun = False
40    Else
50                Exit Sub
60            End If
End Sub


Private Sub Command1_Click()
    Dim iStr As Integer
    Dim tStr As String
    Dim LoopC As Integer

    'Call ParseUserCommand("/cr 5")
    'iStr = InputBox("Escriba la cantidad de mensajes.", "Mensaje por consola de RoleMaster")
    
   ' For LoopC = 1 To iStr
   '     tStr = InputBox("Escriba el mensaje (" & LoopC & "):", "Mensaje por consola de RoleMaster")
   '     If LenB(tStr) <> 0 Then Call WriteServerMessage(tStr)
  '  Next LoopC

End Sub

Private Sub CmdLanzar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call modAnalisis.ClickEnObjetoPos(eTipo.BotonLanzar, X, Y)
End Sub

Private Sub DropGold_Click()
10        Inventario.SelectGold
20        If UserGLD > 0 Then
30            If Not Comerciando Then FrmCantidad.Show , frmMain
40        End If
          
End Sub

Private Sub Form_Load()
       Dim CursorDir As String
          Dim Cursor As Long
          
          'Drag And Drop
10        Set dragInventory = Inventario
          
20        CursorDir = App.path & "\Recursos\Cursor.ani" 'normal1.ani
30        hSwapCursor = SetClassLong(frmMain.hWnd, GLC_HCURSOR, _
              LoadCursorFromFile(CursorDir))
40        hSwapCursor = SetClassLong(frmMain.picInv.hWnd, GLC_HCURSOR, _
              LoadCursorFromFile(CursorDir))
50        hSwapCursor = SetClassLong(frmMain.hlst.hWnd, GLC_HCURSOR, _
              LoadCursorFromFile(CursorDir))
          
          'Consola Inteligente
60        Detectar RecTxt.hWnd, Me.hWnd
          
          ' Handles Form movement (drag and drop).
          'If NoRes Then
             ' Set clsFormulario = New clsFormMovementManager
             ' clsFormulario.Initialize Me, 120
          'End If
          
          'Me.Picture = LoadPicture(DirGraficos & "VentanaPrincipal.JPG")
          
70        InvEqu.Picture = LoadPicture(DirGraficos & "CentroInventario.JPG")
          
80        Call LoadButtons
          
90        Me.Left = 0
100       Me.Top = 0
          Me.Width = 12000
          Me.Height = 9000
          
        'Call SetWindowLong(RecTxt.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
        
        frmMain.lblWeapon(1).Left = frmMain.lblWeapon(0).Left + 1
        frmMain.lblWeapon(1).Top = frmMain.lblWeapon(0).Top + 1
        frmMain.lblWeapon(1).Width = frmMain.lblWeapon(0).Width
        frmMain.lblWeapon(1).Height = frmMain.lblWeapon(0).Height

        frmMain.lblWeapon(2).Left = frmMain.lblWeapon(0).Left - 1
        frmMain.lblWeapon(2).Top = frmMain.lblWeapon(0).Top - 1
        frmMain.lblWeapon(2).Width = frmMain.lblWeapon(0).Width
        frmMain.lblWeapon(2).Height = frmMain.lblWeapon(0).Height
        
        frmMain.lblarmor(1).Left = frmMain.lblarmor(0).Left + 1
        frmMain.lblarmor(1).Top = frmMain.lblarmor(0).Top + 1
        frmMain.lblarmor(1).Width = frmMain.lblarmor(0).Width
        frmMain.lblarmor(1).Height = frmMain.lblarmor(0).Height

        frmMain.lblarmor(2).Left = frmMain.lblarmor(0).Left - 1
        frmMain.lblarmor(2).Top = frmMain.lblarmor(0).Top - 1
        frmMain.lblarmor(2).Width = frmMain.lblarmor(0).Width
        frmMain.lblarmor(2).Height = frmMain.lblarmor(0).Height
        
        frmMain.lblShielder(1).Left = frmMain.lblShielder(0).Left + 1
        frmMain.lblShielder(1).Top = frmMain.lblShielder(0).Top + 1
        frmMain.lblShielder(1).Width = frmMain.lblShielder(0).Width
        frmMain.lblShielder(1).Height = frmMain.lblShielder(0).Height

        frmMain.lblShielder(2).Left = frmMain.lblShielder(0).Left - 1
        frmMain.lblShielder(2).Top = frmMain.lblShielder(0).Top - 1
        frmMain.lblShielder(2).Width = frmMain.lblShielder(0).Width
        frmMain.lblShielder(2).Height = frmMain.lblShielder(0).Height
        
        frmMain.lblhelm(1).Left = frmMain.lblhelm(0).Left + 1
        frmMain.lblhelm(1).Top = frmMain.lblhelm(0).Top + 1
        frmMain.lblhelm(1).Width = frmMain.lblhelm(0).Width
        frmMain.lblhelm(1).Height = frmMain.lblhelm(0).Height

        frmMain.lblhelm(2).Left = frmMain.lblhelm(0).Left - 1
        frmMain.lblhelm(2).Top = frmMain.lblhelm(0).Top - 1
        frmMain.lblhelm(2).Width = frmMain.lblhelm(0).Width
        frmMain.lblhelm(2).Height = frmMain.lblhelm(0).Height
    

        Dim i      As Long
        For i = 1 To 4
            frmMain.Lblham(i).Width = frmMain.Lblham(0).Width 'UserMinHAM & "%" '& UserMaxHAM
            frmMain.Lblham(i).Height = frmMain.Lblham(0).Height
        Next i
        frmMain.Lblham(1).Top = frmMain.Lblham(0).Top + 1 'UserMinHAM & "%" '& UserMaxHAM
        frmMain.Lblham(1).Left = frmMain.Lblham(0).Left

        frmMain.Lblham(2).Top = frmMain.Lblham(0).Top - 1 'UserMinHAM & "%" '& UserMaxHAM
        frmMain.Lblham(2).Left = frmMain.Lblham(0).Left

        frmMain.Lblham(3).Top = frmMain.Lblham(0).Top 'UserMinHAM & "%" '& UserMaxHAM
        frmMain.Lblham(3).Left = frmMain.Lblham(0).Left + 1

        frmMain.Lblham(4).Top = frmMain.Lblham(0).Top 'UserMinHAM & "%" '& UserMaxHAM
        frmMain.Lblham(4).Left = frmMain.Lblham(0).Left - 1
        
        
        frmMain.lblVida(1).Top = frmMain.lblVida(0).Top + 1 'UserMinHAM & "%" '& UserMaxHAM
        frmMain.lblVida(1).Left = frmMain.lblVida(0).Left + 1

        frmMain.lblVida(2).Top = frmMain.lblVida(0).Top - 1 'UserMinHAM & "%" '& UserMaxHAM
        frmMain.lblVida(2).Left = frmMain.lblVida(0).Left + 1

        frmMain.lblVida(3).Top = frmMain.lblVida(0).Top 'UserMinHAM & "%" '& UserMaxHAM
        frmMain.lblVida(3).Left = frmMain.lblVida(0).Left + 2

        frmMain.lblVida(4).Top = frmMain.lblVida(0).Top 'UserMinHAM & "%" '& UserMaxHAM
        frmMain.lblVida(4).Left = frmMain.lblVida(0).Left - 1
        
        
        frmMain.lblMana(1).Top = frmMain.lblMana(0).Top + 1 'UserMinHAM & "%" '& UserMaxHAM
        frmMain.lblMana(1).Left = frmMain.lblMana(0).Left

        frmMain.lblMana(2).Top = frmMain.lblMana(0).Top - 1 'UserMinHAM & "%" '& UserMaxHAM
        frmMain.lblMana(2).Left = frmMain.lblMana(0).Left

        frmMain.lblMana(3).Top = frmMain.lblMana(0).Top 'UserMinHAM & "%" '& UserMaxHAM
        frmMain.lblMana(3).Left = frmMain.lblMana(0).Left + 1

        frmMain.lblMana(4).Top = frmMain.lblMana(0).Top 'UserMinHAM & "%" '& UserMaxHAM
        frmMain.lblMana(4).Left = frmMain.lblMana(0).Left - 1
        
        frmMain.lblEnergia(1).Top = frmMain.lblEnergia(0).Top + 1 'UserMinHAM & "%" '& UserMaxHAM
        frmMain.lblEnergia(1).Left = frmMain.lblEnergia(0).Left

        frmMain.lblEnergia(2).Top = frmMain.lblEnergia(0).Top - 1 'UserMinHAM & "%" '& UserMaxHAM
        frmMain.lblEnergia(2).Left = frmMain.lblEnergia(0).Left

        frmMain.lblEnergia(3).Top = frmMain.lblEnergia(0).Top 'UserMinHAM & "%" '& UserMaxHAM
        frmMain.lblEnergia(3).Left = frmMain.lblEnergia(0).Left + 1

        frmMain.lblEnergia(4).Top = frmMain.lblEnergia(0).Top 'UserMinHAM & "%" '& UserMaxHAM
        frmMain.lblEnergia(4).Left = frmMain.lblEnergia(0).Left - 1
        
        
110           lblmapaname.Visible = True
120       Coord.Visible = False
          
          
          
End Sub

Private Sub LoadButtons()
          Dim GrhPath As String
          Dim i As Integer
          
10        GrhPath = DirGraficos

20        Set cBotonDiamArriba = New clsGraphicalButton
30        Set cBotonDiamAbajo = New clsGraphicalButton
40        Set cBotonGrupo = New clsGraphicalButton
50        Set cBotonOpciones = New clsGraphicalButton
60        Set cBotonEstadisticas = New clsGraphicalButton
70        Set cBotonClanes = New clsGraphicalButton
80        Set cBotonAsignarSkill = New clsGraphicalButton
90        Set cBotonMapa = New clsGraphicalButton
          
100       Set LastPressed = New clsGraphicalButton

          'Set picSkillStar = LoadPicture(GrhPath & "BotonAsignarSkills.bmp")

          'If SkillPoints > 0 Then imgAsignarSkill.Picture = picSkillStar
          
110       imgAsignarSkill.MouseIcon = picMouseIcon
120       lblDropGold.MouseIcon = picMouseIcon
130       lblCerrar.MouseIcon = picMouseIcon
140       lblMinimizar.MouseIcon = picMouseIcon
          
150       For i = 0 To 3
160           picSM(i).MouseIcon = picMouseIcon
170       Next i
End Sub

Public Sub LightSkillStar(ByVal bTurnOn As Boolean)
10        If bTurnOn Then
             ' imgAsignarSkill.Picture = picSkillStar
20        Else
             ' Set imgAsignarSkill.Picture = Nothing
30        End If
End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)
10        If hlst.Visible = True Then
20            If hlst.ListIndex = -1 Then Exit Sub
              Dim sTemp As String
          
30            Select Case Index
                  Case 1 'subir
40                    If hlst.ListIndex = 0 Then Exit Sub
50                Case 0 'bajar
60                    If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
70            End Select
          
80            Call WriteMoveSpell(Index = 1, hlst.ListIndex + 1)
              
90            Select Case Index
                  Case 1 'subir
100                   sTemp = hlst.List(hlst.ListIndex - 1)
110                   hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
120                   hlst.List(hlst.ListIndex) = sTemp
130                   hlst.ListIndex = hlst.ListIndex - 1
140               Case 0 'bajar
150                   sTemp = hlst.List(hlst.ListIndex + 1)
160                   hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
170                   hlst.List(hlst.ListIndex) = sTemp
180                   hlst.ListIndex = hlst.ListIndex + 1
190           End Select
200       End If
End Sub

Public Sub ControlSM(ByVal Index As Byte, ByVal Mostrar As Boolean)
      Dim GrhIndex As Long
      Dim SR As RECT
      Dim DR As RECT

10    GrhIndex = GRH_INI_SM + Index + SM_CANT * (CInt(Mostrar) + 1)

20    With GrhData(GrhIndex)
30        SR.Left = .sX
40        SR.Right = SR.Left + .pixelWidth
50        SR.Top = .sY
60        SR.Bottom = SR.Top + .pixelHeight
          
70        DR.Left = 0
80        DR.Right = .pixelWidth
90        DR.Top = 0
100       DR.Bottom = .pixelHeight
110   End With

120   Select Case Index

          Case eSMType.sSafemode
130           If Mostrar Then
140               Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, _
                      255, 0, True, False, True)
150               picSM(Index).ToolTipText = "Seguro activado."
160               frmMain.IconoSeg.Caption = ""
170           Else
180               Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, _
                      255, 0, 0, True, False, True)
190               picSM(Index).ToolTipText = "Seguro desactivado."
200               frmMain.IconoSeg.Caption = "X"
210           End If
              
220       Case eSMType.DragMode
230           If Mostrar Then
240               Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_DRAG_DESACTIVADO, 255, _
                      0, 0, True, False, True)
250               frmMain.IconosegD.Caption = "X"
260           Else
270               Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_DRAG_ACTIVADO, 0, 255, _
                      0, True, False, True)
280               frmMain.IconosegD.Caption = ""
290           End If

300   End Select

310   SMStatus(Index) = Mostrar
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
      '***************************************************
      'Autor: Unknown
      'Last Modification: 18/11/2009
      '18/11/2009: ZaMa - Ahora se pueden poner comandos en los mensajes personalizados (execpto guildchat y privados)
      '***************************************************
          
      'ShowConsoleMsg KeyCode
          
           
           
10    If Pulsacion_Fisica = False Then
20    Exit Sub
30    End If
40    Pulsacion_Fisica = True
          
50        If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) And (Not _
              SendRmstxt.Visible) And (Not SendGms.Visible) Then
          
          
60            If KeyCode = vbKeyF5 Then
70                WriteRequestRetos
                  
              'Call AddtoRichTextBox(frmMain.RecTxt, "Sistema Desactivado", 0, 200, 200, False, False, True)
80            End If
              
90            If KeyCode = vbKeyF1 Then
100               WriteLevel
110           End If

                If KeyCode = vbKeyB Then
                    If esGM(UserCharIndex) Then
                    frmBuscar.Show , frmMain
                    End If
                End If
              
120           If KeyCode = vbKeyF2 Then
130
140           WriteReset
150                  Call ShowConsoleMsg("Reloguea para ver los cambios.", , , , True)
170           End If
              
180           If KeyCode = vbKeyF3 Then
                    WritePartyClient 5
200           End If
              
210           If KeyCode = vbKeyEnd Then
220               WriteMeditate
230           End If
                
                
                If KeyCode = vbKeyZ And TieneClan Then
                    SeguroClanes = Not SeguroClanes
                    Call WriteSeguroClan
                End If
              
270           If KeyCode = vbKeyF9 Then
              'Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
280               Call frmOpciones.Show(vbModeless, frmMain)
290           Exit Sub
300           End If
                   

              
310           If KeyCode = vbKeyF10 Then
                  'Intervalo permite usar este sistema?
320               If Not FotoD_CanSend Then
330               Call AddtoRichTextBox(frmMain.RecTxt, _
                      "Haz alcanzado el máximo de envio de 1 FotoDenuncia por minuto. Esperá unos instantes y volve a intentar.", _
                      0, 200, 200, False, False, True)
340               Exit Sub
350               End If
                  'Aca guardamos el string que nos devuelve FotoD_Capturar.
                  Dim nString    As String
360               FotoD_Capturar nString
                  'Si el string da nullo, es por que nadie esta insultando.
370               If nString = vbNullString Then
380               Call AddtoRichTextBox(frmMain.RecTxt, _
                      "Nadie te esta insultando. Las FotoDenuncias solo sirven para denunciar agravios.", _
                      0, 200, 200, False, False, True)
390               Else 'Si no, enviamos.
400               Call AddtoRichTextBox(frmMain.RecTxt, _
                      "La FotoDenuncia fue sacada correctamente.", 0, 200, 200, False, _
                      False, True)
410               WriteDenounce "[FOTODENUNCIAS] : " & nString
420               End If
430               End If


              ' Anti F8 provisorio
              Select Case KeyCode
                  Case vbKeyNumpad1
                        If CustomKeys.BindedKey(eKeyType.mKeyUseObject) <> vbKeyNumpad1 Then
                            CantKey1 = CantKey1 + 1
                        End If
                        
                        If CantKey1 = 5 Then
                            CantKey1 = 0
                           ' WriteReportcheat UserName, "Posible Uso de F8"
                        End If
                  Case vbKeyNumpad0
                        If CustomKeys.BindedKey(eKeyType.mKeyUseObject) <> vbKeyNumpad0 Then
                            CantKey0 = CantKey0 + 1
                        End If
                        
                        If CantKey0 = 5 Then
                            CantKey0 = 0
                           ' WriteReportcheat UserName, "Posible Uso de F8"
                        End If
                  Case vbKeyNumpad2
                        If CustomKeys.BindedKey(eKeyType.mKeyUseObject) <> vbKeyNumpad2 Then
                            CantKey2 = CantKey2 + 1
                        End If
                        
                        If CantKey2 = 5 Then
                            CantKey2 = 0
                           ' WriteReportcheat UserName, "Posible Uso de F8"
                        End If
                  Case vbKeyF8
                        If CustomKeys.BindedKey(eKeyType.mKeyUseObject) <> vbKeyF8 Then
                            CantF8 = CantF8 + 1
                        End If
                        
                        If CantF8 = 5 Then
                            CantF8 = 0
                            If Not esGM(UserCharIndex) Then
                                WriteReportcheat UserName, "Posible Uso de F8"
                            End If
                        End If
              End Select
              
              
              If esGM(UserCharIndex) Then
                    If KeyCode = vbKeyQ Then
                        If SendTxt.Visible Or SendGms.Visible Then Exit Sub
                        
                        If Not FrmCantidad.Visible Then
                            ShowConsoleMsg "Escriba un mensaje global.", 0, 255, 255
                            SendRmstxt.Visible = True
                            SendRmstxt.SetFocus
                            End If
                        End If
             
                    If KeyCode = vbKeyG Then
                        If SendTxt.Visible Or SendRmstxt.Visible Then Exit Sub
                        
                        If Not FrmCantidad.Visible Then
                            ShowConsoleMsg "Escriba un mensaje a los Game Masters.", 0, 255, 255
                            SendGms.Visible = True
                            SendGms.SetFocus
                        End If
                    End If
              
                    If KeyCode = vbKeyI Then
                        Call ParseUserCommand("/INVISIBLE")
                    End If
    
                    If KeyCode = vbKeyW Then
                        Call ParseUserCommand("/TRABAJANDO")
                    End If
                    
                    If KeyCode = vbKeyP Then
                        Call ParseUserCommand("/PANELGM")
                    End If
              
              End If
              
              'Checks if the key is valid
440           If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
450               Select Case KeyCode
                      Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
460                       Audio.MusicActivated = Not Audio.MusicActivated
                          
470                   Case CustomKeys.BindedKey(eKeyType.mKeyToggleSound)
480                       Audio.SoundActivated = Not Audio.SoundActivated
                          
490                   Case CustomKeys.BindedKey(eKeyType.mKeyToggleFxs)
500                       Audio.SoundEffectsActivated = Not _
                              Audio.SoundEffectsActivated
                      
510                   Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
520                       Call AgarrarItem
                      
530                   Case CustomKeys.BindedKey(eKeyType.mKeyToggleCombatMode)
540                       Call WriteCombatModeToggle
550                       Iscombate = Not Iscombate
                      
560                       Case vbKeyMultiply:
570                       If frmMain.IconoSeg.Visible Then
580                       With FontTypes(FontTypeNames.FONTTYPE_INFO)
590                       Call ShowConsoleMsg("Escribe /SEG para quitar el seguro", _
                              .red, .green, .blue, .bold, .italic)
600                       End With
                          'Call AddtoRichTextBox(frmMain.RecTxt, "Escribe /SEG para quitar el seguro", 0, 200, 200, False, False, True)
610                       Else
620                       Call WriteSafeToggle
630                       End If
                      
640                   Case vbKeyZ:
650                       If DialogosClanes.Activo = False Then
660                           Call _
                                  ShowConsoleMsg("Consola flotante de clanes activada.", _
                                  255, 200, 200)
670                           DialogosClanes.Activo = True
680                       Else
690                           Call _
                                  ShowConsoleMsg("Consola flotante de clanes desactivada.", _
                                  255, 200, 200)
700                           DialogosClanes.Activo = False
710                       End If
                      
                      
                      
720                   Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
730                       Call EquiparItem
                      
740                   Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
750                       Nombres = Not Nombres
                      
760                   Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
770                       If UserEstado = 1 Then
780                           With FontTypes(FontTypeNames.FONTTYPE_INFO)
790                               Call ShowConsoleMsg("¡¡Estás muerto!!", .red, _
                                      .green, .blue, .bold, .italic)
800                           End With
810                       Else
820                           Call WriteWork(eSkill.Domar)
830                       End If
                          
840                   Case CustomKeys.BindedKey(eKeyType.mKeySteal)
850                       If UserEstado = 1 Then
860                           With FontTypes(FontTypeNames.FONTTYPE_INFO)
870                               Call ShowConsoleMsg("¡¡Estás muerto!!", .red, _
                                      .green, .blue, .bold, .italic)
880                           End With
890                       Else
900                           Call WriteWork(eSkill.Robar)
910                       End If
                          
920                   Case CustomKeys.BindedKey(eKeyType.mKeyRETOS)
930                       WriteRequestRetos
                          
                           'frmRetos.Show , frmMain
                          
940                   Case CustomKeys.BindedKey(eKeyType.mKeyHide)
950                       If UserEstado = 1 Then
960                           With FontTypes(FontTypeNames.FONTTYPE_INFO)
970                               Call ShowConsoleMsg("¡¡Estás muerto!!", .red, _
                                      .green, .blue, .bold, .italic)
980                           End With
990                       Else
1000                          Call WriteWork(eSkill.Ocultarse)
1010                      End If
                                          
1020                  Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
1030                      Call TirarItem
                      
1040                  Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
1050                      If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
                              
1060                      If MainTimer.Check(TimersIndex.UseItemWithU) Then
1070                          Call UsarItem(0)
1080                      End If
                      
1090                  Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
1100                      If MainTimer.Check(TimersIndex.SendRPU) Then
1110                          Call WriteRequestPositionUpdate
1120                          Beep
1130                      End If
                     ' Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                          'Call WriteSafeToggle

1140                  Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe)
1150                      Call WriteResuscitationToggle
1160              End Select
1170          Else
1180              Select Case KeyCode
                      'Custom messages!
                      Case vbKey0 To vbKey9
                          Dim CustomMessage As String
                          
1190                      CustomMessage = CustomMessages.Message((KeyCode - 39) Mod _
                              10)
1200                      If LenB(CustomMessage) <> 0 Then
                              ' No se pueden mandar mensajes personalizados de clan o privado!
1210                          If UCase(Left(CustomMessage, 5)) <> "/CMSG" And _
                                  Left(CustomMessage, 1) <> "\" Then
                                  
1220                              Call ParseUserCommand(CustomMessage)
1230                          End If
1240                      End If
1250              End Select
1260          End If
1270      End If
          
1280  Select Case KeyCode
              Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
1290                      If (Not Comerciando) And (Not Canjeando) And (Not _
                              MirandoAsignarSkills) And (Not frmMSG.Visible) And (Not _
                              MirandoForo) And (Not frmEstadisticas.Visible) And (Not _
                              FrmCantidad.Visible) Then
1300              End If
1310     If bCMSG = True Then Exit Sub 'Si está activado el cmsgimg lo cancelamos
1320                  SendTxt.Visible = True 'Mostramos el Sendtxt
1330                  SendTxt.SetFocus 'Lo priorizamos
1340                  cmsgSupr = True 'Activamos que fue con la tecla suprimir con lo que fue abierta
1350                  bCMSG = True 'Activamos que se puso el CMSGimg
                  
1360          Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
1370              Call ScreenCapture
              
1380          Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
                  'If UserMinMAN = UserMaxMAN Then Exit Sub
                  
1390              If UserEstado = 1 Then
1400                  With FontTypes(FontTypeNames.FONTTYPE_INFO)
1410                      Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, _
                              .bold, .italic)
1420                  End With
1430                  Exit Sub
1440              End If
                      
1450          Call WriteMeditate
            
1460          Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)
              Call WritePartyClient(1)
              
1480          Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)
1490              If UserEstado = 1 Then
1500                  With FontTypes(FontTypeNames.FONTTYPE_INFO)
1510                      Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, _
                              .bold, .italic)
1520                  End With
1530                  Exit Sub
1540              End If
                  
1550              If macrotrabajo.Enabled Then
1560                  Call DesactivarMacroTrabajo
1570              Else
1580                  Call ActivarMacroTrabajo
1590              End If
              
1600          Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
1610              If frmMain.macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
1620              Call WriteQuit
                  
1630          Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
1640                      If Shift <> 0 Then Exit Sub
             
1650                      If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit _
                              Sub 'Check if arrows interval has finished.
1660                      If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
1670                          If Not MainTimer.Check(TimersIndex.CastAttack) Then _
                                  Exit Sub 'Corto intervalo Golpe-Hechizo
1680                      Else
1690                          If Not MainTimer.Check(TimersIndex.Attack) Or _
                                  UserDescansar Or UserMeditar Then Exit Sub
1700                      End If
           
1710                     If TrainingMacro.Enabled Then DesactivarMacroHechizos
1720                     If macrotrabajo.Enabled Then DesactivarMacroTrabajo
1730                 Call WriteAttack
1740                 Attack = True
1750             charlist(UserCharIndex).Arma.WeaponWalk(charlist(UserCharIndex).Heading).Started _
                     = 1
1760             charlist(UserCharIndex).Escudo.ShieldWalk(charlist(UserCharIndex).Heading).Started _
                     = 1
                   
1770               If Iscombate = False Then
1780               Attack = False
1790                charlist(UserCharIndex).Arma.WeaponWalk(charlist(UserCharIndex).Heading).Started _
                        = 0
1800             charlist(UserCharIndex).Escudo.ShieldWalk(charlist(UserCharIndex).Heading).Started _
                     = 0
1810             End If
              
1820   Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
1830              If SendCMSTXT.Visible Then Exit Sub
                  
1840              If (Not Comerciando) And (Not Canjeando) And (Not _
                      MirandoAsignarSkills) And (Not frmMSG.Visible) And (Not _
                      MirandoForo) And (Not frmEstadisticas.Visible) And (Not _
                      FrmCantidad.Visible) Then
1850                  SendTxt.Visible = True
1860                  SendTxt.SetFocus
1870              End If
                  
1880      End Select
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
10        MouseBoton = Button
20        MouseShift = Shift
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As _
    Single)
10        clicX = X
20        clicY = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
10        MouseX = X - MainViewPic.Left
20        MouseY = Y - MainViewPic.Top
         
            'Trim to fit screen
30        If MouseX < 0 Then
40            MouseX = 0
50        ElseIf MouseX > MainViewShp.Width Then
60            MouseX = MainViewPic.Width
70        End If
       
          'Trim to fit screen
80        If MouseY < 0 Then
90            MouseY = 0
100       ElseIf MouseY > MainViewShp.Height Then
110           MouseY = MainViewShp.Height
120       End If
          
130        Ancho = lblmapaname.Left + lblmapaname.Width
140       alto = lblmapaname.Top + lblmapaname.Height
150       If X > lblmapaname.Left And X < Ancho And Y > lblmapaname.Top And Y < alto _
              Then
160           lblmapaname.Visible = False
170           Coord.Visible = True
180       Else
190           lblmapaname.Visible = True
200           Coord.Visible = False
210       End If
          
220       Ancho = lvllbl(0).Left + lvllbl(0).Width
230       alto = lvllbl(0).Top + lvllbl(0).Height
240       If X > lvllbl(0).Left And X < Ancho And Y > lvllbl(0).Top And Y < alto Then
250           lvllbl(0).Visible = False
260           lvllbl(1).Visible = False
270           lvllbl(2).Visible = False
280           lvllbl(3).Visible = False
290           lvllbl(4).Visible = False
300           lblporclvl(0).Visible = True
310            lblporclvl(1).Visible = True
320             lblporclvl(2).Visible = True
330              lblporclvl(3).Visible = True
340               lblporclvl(4).Visible = True
350       Else
360           lvllbl(0).Visible = True
370            lvllbl(1).Visible = True
380             lvllbl(2).Visible = True
390             lvllbl(3).Visible = True
400             lvllbl(4).Visible = True
410           lblporclvl(0).Visible = False
420           lblporclvl(1).Visible = False
430           lblporclvl(2).Visible = False
440           lblporclvl(3).Visible = False
450           lblporclvl(4).Visible = False
460       End If

          
          'Trim to fit screen
470       If MouseY < 0 Then
480           MouseY = 0
490       ElseIf MouseY > MainViewShp.Height Then
500           MouseY = MainViewShp.Height
510       End If
520       Inventario.uMoveItem = False
530       Inventario.sMoveItem = False
          
540         If SendTxt.Visible Then
550           SendTxt.SetFocus
560       End If
          
End Sub


Private Sub CMSG_Click()
10    Call Audio.PlayWave(SND_CLICK)
20        If Not CharTieneClan Then
30        Call AddtoRichTextBox(frmMain.RecTxt, "¡No perteneces a ningún clan!", 0, _
              200, 200, False, False, True)
40          If bCMSG = False Then
50          cmsgSupr = False
60        Exit Sub
70        End If
80    Else
90        bCMSG = Not bCMSG
100       If bCMSG Then
110       cmsgSupr = False
120           CMSG.Picture = LoadPicture(App.path & "\Recursos\CMSG.jpg")
130       Call AddtoRichTextBox(frmMain.RecTxt, _
              "Todo lo que digas sera escuchado por tu clan.", 0, 200, 200, False, False)
140       Else
150       Call AddtoRichTextBox(frmMain.RecTxt, _
              "Dejas de ser escuchado por tu clan. ", 0, 200, 200, False, False)
160           CMSG.Picture = LoadPicture("")
170       End If
180       End If
End Sub

Private Sub hlst_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call modAnalisis.ClickLista
End Sub

Private Sub IconoSeg_Click()
10    WriteSafeToggle
End Sub

Private Sub IconosegD_Click()
      'Sistema Deniega el Item
10    WriteDragToggle
End Sub




Private Sub ImageRanking_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    Call AddtoRichTextBox(frmMain.RecTxt, _
          "Presiona doble click para abrir el Ranking de Personajes.", 0, 200, 200, False, _
          False)
End Sub
Private Sub imageRanking_dblclick()
10                Call FrmRanking.Show(vbModeless, frmMain)
End Sub

Private Sub imageparty_click()
    
    'Dim Bytes() As Byte
   ' Capturar_Guardar App.path & "\10.jpg"
    'If FileExist(App.path & "\10.jpg", vbArchive) Then
           
         
    

       ' Bytes = LeerJPG(App.path & "\10.jpg")
        'WriteSendCaptureImage Bytes()
    
       ' Exit Sub
        
    'Else
       ' MsgBox "No existe"
       ' Exit Sub
    'End If
    
    WritePartyClient 1

End Sub
Public Function LeerJPG(ByRef file_path As String) As Byte()

    If Len(Dir$(file_path)) <> 0 Then

        Dim fFile  As Integer
        Dim temp() As Byte
    
        fFile = FreeFile()
        
        ReDim temp(FileLen(file_path)) As Byte

        Open file_path For Binary As #fFile

        Get #fFile, , temp()

        Close #fFile
 
        LeerJPG = temp()
 
    End If

End Function
Private Sub ImageQUest_DblClick()
10    Call WriteQuestListRequest
End Sub
Private Sub imageQUest_click()
10    Call Audio.PlayWave(SND_CLICK)
20    Call AddtoRichTextBox(frmMain.RecTxt, _
          "Presiona doble click para ver la información de tus quests.", 0, 200, 200, _
          False, False)
End Sub


Private Sub imgEstadisticas_Click()

10    Call Audio.PlayWave(SND_CLICK)

       Dim i As Integer
20        If SkillPoints > 0 Then
30        imgAsignarSkill.Visible = True
40        Else
50        imgAsignarSkill.Visible = False
60        imgAsignarSkill.Enabled = False
70        End If

          
80        LlegaronSkills = False
90        Call WriteRequestSkills
100       Call FlushBuffer
          
110       Do While Not LlegaronSkills
120           DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
130       Loop
140       LlegaronSkills = False
          
150       For i = 1 To NUMSKILLS
160           frmSkills3.text1(i).Caption = UserSkills(i)
170       Next i
          
180       Alocados = SkillPoints
190        LlegaronAtrib = False
200       LlegaronSkills = False
210       LlegoFama = False
220       Call WriteRequestAtributes
230       Call WriteRequestSkills
240       Call WriteRequestMiniStats
250       Call WriteRequestFame
260       Call FlushBuffer
270       Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
280           DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
290       Loop
300       frmSkills3.Iniciar_Labels
310       frmSkills3.Show , frmMain
320       frmSkills3.lbldatos.Caption = "Nivel: " & UserLvl & " Experiencia: " & _
              UserExp & "/" & UserPasarNivel
330       Alocados = SkillPoints
340       frmSkills3.puntos.Caption = SkillPoints
350       frmSkills3.Show , frmMain

360       LlegaronAtrib = False
370       LlegaronSkills = False
380       LlegoFama = False
End Sub

Private Sub imgPMSG_Click()
10    Call Audio.PlayWave(SND_CLICK)
      '----Boton partys Style TDS by IRuleDK----
20    PMSG = False 'Nos fijamos que no este activado con la tecla suprimir
30    If PMSGimg = False Then 'Si no había apretado el botón -> lo activamos y le ponemos la imagen estilo TDS
40    PMSGimg = True
50    imgPMSG.Picture = LoadPicture(App.path & "\Recursos\Pmsg.jpg") 'Grafico del botón estilo tds
60    Call AddtoRichTextBox(frmMain.RecTxt, _
          "Todo lo que digas sera escuchado por tu party. ", 255, 200, 200, False, False)
70    Else 'si ya estaba apretado lo desactivamos
80    PMSGimg = False 'desactivamos el boton
90    imgPMSG.Picture = LoadPicture("") 'lo ponemos normal sacandole la imagen verde
100   Call AddtoRichTextBox(frmMain.RecTxt, "Dejas de ser escuchado por tu party. ", _
          255, 200, 200, False, False)
110   Call ControlSM(eSMType.mWork, True)
120   End If
End Sub

Private Sub Label1_Click()
10    Call ParseUserCommand("/invisible")
End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
          
          
10        LastPressed.ToggleToNormal
20        Inventario.uMoveItem = False
30        Inventario.sMoveItem = False



End Sub

Private Sub Label16_Click()

End Sub

Private Sub Label3_Click()
 Call Audio.PlayWave(SND_CLICK)
          
    MsgBox "Desactivado Temporalmente", vbModeless
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Call modAnalisis.ClickEnObjetoPos(eTipo.BotonInventario, X, Y)
End Sub

Private Sub Label5_Click()
10    Call WriteWorking
End Sub

Private Sub Label6_Click()
    Call ParseUserCommand("/ONLINE")
End Sub

Private Sub Labelgm1_Click()
10    Call ParseUserCommand("/telep yo 1 50 50")
End Sub

Private Sub Labelgm2_Click()
10    If MsgBox("Esta todo listo para empezar la daga rusa?", vbYesNo, "Daga rusa") = _
          vbYes Then
20    Call _
          ParseUserCommand("/RMSG Luego de la cuenta envien los interesados en la Daga Rusa")
30    Call ParseUserCommand("/cr 5")
40    End If
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call modAnalisis.ClickEnObjetoPos(eTipo.BotonHechizos, X, Y)
End Sub

Private Sub Labelgm3_Click()
10    Call ParseUserCommand("/cr 5")
End Sub

Private Sub Labelgm4_Click()
10    frmPanelGm.Show , frmMain
End Sub

Private Sub Labelgm44_Click()
10    frmPanelGMS.Show , frmMain
End Sub

Private Sub Labelgm5_Click()
10    Call ParseUserCommand("/online")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
10        If prgRun = True Then
20            prgRun = False
30            Cancel = 1
40        End If
End Sub

Private Sub imgClanes_Click()
10    Call Audio.PlayWave(SND_CLICK)
20        If frmGuildLeader.Visible Then Unload frmGuildLeader
30        Call WriteRequestGuildLeaderInfo
End Sub

Private Sub imgAsignarSkill_Click()
          Dim i As Integer
10        If SkillPoints > 0 Then
20        imgAsignarSkill.Visible = True
30        Else
40        imgAsignarSkill.Visible = False
50        imgAsignarSkill.Enabled = False
60        End If

          
70        LlegaronSkills = False
80        Call WriteRequestSkills
90        Call FlushBuffer
          
100       Do While Not LlegaronSkills
110           DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
120       Loop
130       LlegaronSkills = False
          
140       For i = 1 To NUMSKILLS
150           frmSkills3.text1(i).Caption = UserSkills(i)
160       Next i
          
170       Alocados = SkillPoints
180        LlegaronAtrib = False
190       LlegaronSkills = False
200       LlegoFama = False
210       Call WriteRequestAtributes
220       Call WriteRequestSkills
230       Call WriteRequestMiniStats
240       Call WriteRequestFame
250       Call FlushBuffer
260       Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
270           DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
280       Loop
290       frmSkills3.Iniciar_Labels
300       frmSkills3.Show , frmMain
310       frmSkills3.lbldatos.Caption = "Nivel: " & UserLvl & " Experiencia: " & _
              UserExp & "/" & UserPasarNivel
320       Alocados = SkillPoints
330       frmSkills3.puntos.Caption = SkillPoints
340       frmSkills3.Show , frmMain

350       LlegaronAtrib = False
360       LlegaronSkills = False
370       LlegoFama = False

End Sub

Private Sub imgGrupo_Click()
10    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub imgInvScrollDown_Click()
10        Call Inventario.ScrollInventory(True)
End Sub

Private Sub imgInvScrollUp_Click()
10        Call Inventario.ScrollInventory(False)
End Sub

Private Sub imgOpciones_Click()
10    Call Audio.PlayWave(SND_CLICK)
20        Call frmOpciones.Show(vbModeless, frmMain)
End Sub

Private Sub InvEqu_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        LastPressed.ToggleToNormal
20            Inventario.uMoveItem = False
30        Inventario.sMoveItem = False
End Sub
Private Sub lblScroll_Click(Index As Integer)
10        Inventario.ScrollInventory (Index = 0)
End Sub


Private Sub lblFPS_Click()
10        FrmDuelos.Show
End Sub

Private Sub lblMinimizar_Click()
10        Me.WindowState = 1
End Sub

Private Sub macrotrabajo_Timer()
10        If Inventario.SelectedItem = 0 Then
20            DesactivarMacroTrabajo
30            Exit Sub
40        End If
          
          'Macros are disabled if not using Argentum!
50        If Not Application.IsAppActive() Then
60            DesactivarMacroTrabajo
70            Exit Sub
80        End If
          
90        If UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = _
              eSkill.Mineria Or UsingSkill = FundirMetal Or (UsingSkill = eSkill.Herreria _
              And Not frmHerrero.Visible) Then
100           Call WriteWorkLeftClick(tX, tY, UsingSkill)
110           UsingSkill = 0
120       End If
          
          'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
130        If Not (frmCarp.Visible = True) Then Call UsarItem(0)
End Sub

Public Sub ActivarMacroTrabajo()
10        If Iscombate Then
20        With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
30                            Call _
                                  ShowConsoleMsg("No puedes trabajar en modo combate.", _
                                  .red, .green, .blue, .bold, .italic)
40                        End With
50         Exit Sub
60       End If

70        macrotrabajo.Interval = INT_MACRO_TRABAJO
80        macrotrabajo.Enabled = True
90        Call AddtoRichTextBox(frmMain.RecTxt, "Empiezas a trabajar", 0, 200, 200, _
              False, False, True)
100       Call ControlSM(eSMType.mWork, True)
        

End Sub

Public Sub DesactivarMacroTrabajo()

10        macrotrabajo.Enabled = False
20        MacroBltIndex = 0
30        UsingSkill = 0
40        MousePointer = vbDefault
50        Call AddtoRichTextBox(frmMain.RecTxt, "Dejas de trabajar", 0, 200, 200, _
              False, False, True)
60        Call ControlSM(eSMType.mWork, False)
       
End Sub

Private Sub Minimizar_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    Me.WindowState = 1
End Sub

Private Sub mnuEquipar_Click()
10        Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
10        Call WriteLeftClick(tX, tY)
20        Call WriteCommerceStart
End Sub

Private Sub mnuNpcDesc_Click()
10        Call WriteLeftClick(tX, tY)
End Sub
Private Sub MainViewPic_Click()
10        Form_Click
20          If SendTxt.Visible Then
30            SendTxt.SetFocus
40        End If
End Sub

Private Sub MainViewPic_DblClick()
10        Form_DblClick
20          If SendTxt.Visible Then
30            SendTxt.SetFocus
40        End If
End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        MouseBoton = Button
20        MouseShift = Shift

30        Call ConvertCPtoTP(X, Y, tX, tY)
          
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
          
10        MouseX = X
20        MouseY = Y
          
          
          'LastPressed.ToggleToNormal
          
30        Call ConvertCPtoTP(X, Y, tX, tY)
          
40        If Inventario.sMoveItem And Not vbKeyShift Then
50            General_Drop_X_Y tX, tY
60            Inventario.uMoveItem = False
70        Else
80            If Inventario.sMoveItem And vbKeyShift Then
90            FrmCantidad.Show , frmMain
100           End If
110       End If

120         If SendTxt.Visible Then
130           SendTxt.SetFocus
140       End If

End Sub

Private Sub MainViewPic_MouseUp(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
          
10        clicX = X
20        clicY = Y

         
           
End Sub


Private Sub mnuTirar_Click()
10        Call TirarItem
20        Inventario.uMoveItem = False
30        Inventario.sMoveItem = False
End Sub

Private Sub mnuUsar_Click()
10        Call UsarItem(0)
End Sub

Private Sub PicMH_Click()
10        Call AddtoRichTextBox(frmMain.RecTxt, _
              "Auto lanzar hechizos. Utiliza esta habilidad para entrenar únicamente. Para activarlo/desactivarlo utiliza F7.", _
              255, 255, 255, False, False, True)
End Sub

Private Sub lblmapaname_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        lblmapaname.Visible = False
20        Coord.Visible = True
End Sub
Private Sub coord_click()
10    Call Audio.PlayWave(SND_CLICK)
20     Call AddtoRichTextBox(frmMain.RecTxt, _
           "Presiona doble click para abrir el mapa del mundo.", 0, 200, 200, False, _
           False)
End Sub
Private Sub coord_dblclick()
10    Call frmMapa.Show(vbModeless, frmMain)
End Sub

Private Sub picSM_DblClick(Index As Integer)
10    Select Case Index
          Case eSMType.sResucitation
20            Call WriteResuscitationToggle
              
30        Case eSMType.sSafemode
40            Call WriteSafeToggle
              
50        Case eSMType.mSpells
60            If UserEstado = 1 Then
70                With FontTypes(FontTypeNames.FONTTYPE_INFO)
80                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, _
                          .bold, .italic)
90                End With
100               Exit Sub
110           End If
              
       
120       Case eSMType.mWork
130           If UserEstado = 1 Then
140               With FontTypes(FontTypeNames.FONTTYPE_INFO)
150                   Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, _
                          .bold, .italic)
160               End With
170               Exit Sub
180           End If
              
190           If macrotrabajo.Enabled Then
200               Call DesactivarMacroTrabajo
210           Else
220               Call ActivarMacroTrabajo
230           End If
240   End Select
End Sub



Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
              'Sistema de botón clanes estilo TDS by AmenO
       

          
          
10       If KeyCode = CustomKeys.BindedKey(eKeyType.mKeyTalk) Then 'Si se apretó enter entonces:
20            Call Dialogos.RemoveDialog(UserCharIndex)
30                    If PMSGimg = True Then 'Si está activado el PMSGimg
40                           sPartyChat = SendTxt.Text 'Mandamos lo que sea de Party
                         
                              '// Es mas rapido comprar byts que cadenas de letras :P
                             ' If sPartyChat <> "" Then
                         
50                            If LenB(sPartyChat) <> 0 Then
60                                    Call ParseUserCommand("/PMSG " & sPartyChat)
70                            End If
                              'Reiniciamos los valores
80                           sPartyChat = vbNullString ' // Mejor vbnullstring que ""
90                           SendTxt.Text = vbNullString ' // Mejor vbnullstring que ""
100                   End If
110       End If
       
       
120          If KeyCode = CustomKeys.BindedKey(eKeyType.mKeyTalk) Then 'Si se apretó enter entonces:
130                   Call Dialogos.RemoveDialog(UserCharIndex)
140                   If bCMSG = True Then 'Si está activado el CMSGimg
150                          stxtbuffercmsg = SendTxt.Text 'Mandamos lo que sea de CLAN
                             
                              '// Es mas rapido comprar byts que cadenas de letras :P
                             ' If stxtbuffercmsg <> "" Then
                             
160                           If LenB(stxtbuffercmsg) <> 0 Then
170                                   Call ParseUserCommand("/CMSG " & stxtbuffercmsg)
180                           End If
       
                              'Reiniciamos los valores
190                          stxtbuffercmsg = vbNullString ' // Mejor vbnullstring que ""
200                           SendTxt.Text = vbNullString ' // Mejor vbnullstring que ""
       
210                          If cmsgSupr = True Then 'Revisamos si fue con Suprimir
220                                   bCMSG = False 'Si fue así desactivamos el cmsgimg
230                          End If
       
240                          KeyCode = 0
250                          SendTxt.Visible = False
       
260                          If picInv.Visible Then
270                                  picInv.SetFocus
280                          Else
290                                  hlst.SetFocus
300                          End If
       
310                          Exit Sub
       
320                  End If
       
330                  If LenB(stxtbuffer) <> 0 Then
340                          Call ParseUserCommand(stxtbuffer) ' Y si no había nada de CMSG hacemos el proceso común para hablar
350                   End If
       
360                   stxtbuffer = vbNullString ' // Mejor vbnullstring que ""
370                  SendTxt.Text = vbNullString ' // Mejor vbnullstring que ""
380                   KeyCode = 0
390                   SendTxt.Visible = False
             
400                   If picInv.Visible Then
410                           picInv.SetFocus
420                   Else
430                           hlst.SetFocus
440                   End If
450           End If
       
              '----Boton clanes Style TDS by AmenO----
End Sub


Private Sub Second_Timer()
10        If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
          
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
10        If UserEstado = 1 Then
20            With FontTypes(FontTypeNames.FONTTYPE_INFO)
30                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, _
                      .italic)
40            End With
50        Else
60            If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < _
                  MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
70                If Inventario.Amount(Inventario.SelectedItem) = 1 Then
80                    Call WriteDrop(Inventario.SelectedItem, 1)
90                    Inventario.uMoveItem = False
100                   Inventario.sMoveItem = False
110               Else
120                   If Inventario.Amount(Inventario.SelectedItem) > 1 Then
130                       If Not Comerciando Then FrmCantidad.Show , frmMain
140                   End If
150               End If
160           End If
170       End If
End Sub

Private Sub AgarrarItem()
10        If UserEstado = 1 Then
20            With FontTypes(FontTypeNames.FONTTYPE_INFO)
30                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, _
                      .italic)
40            End With
50        Else
60            Call WritePickUp
70        End If
End Sub

Private Sub UsarItem(ByVal SecondaryClick As Byte)
10        If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub
          If Not CheckInterval(SecondaryClick) Then Exit Sub
          'If (timeGetTime - Intervalos(SecondaryClick).ModifyTime) <= 200 Then Exit Sub
          
          'ShowConsoleMsg
          
20        If Comerciando Then Exit Sub
30        If Canjeando Then Exit Sub
          
          Dim ItemIndex As Integer
              
40        ItemIndex = Inventario.SelectedItem
          
50        If (ItemIndex > 0) And (ItemIndex < MAX_INVENTORY_SLOTS + 1) Then
              
60            If Inventario.ObjType(ItemIndex) <> eOBJType.otBarcos Then
70                If UserEstado = 1 Then

80                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
90                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, _
                              .bold, .italic)

100                   End With

110                   Exit Sub

120               End If

130           End If

140           Call WriteUseItem(ItemIndex, SecondaryClick)
              Call AssignedInterval(SecondaryClick)
            
150       End If

End Sub

Private Sub EquiparItem()
10        If UserEstado = 1 Then
20            With FontTypes(FontTypeNames.FONTTYPE_INFO)
30                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, _
                          .bold, .italic)
40            End With
50        Else
60            If Comerciando Then Exit Sub
70            If Canjeando Then Exit Sub
              
80            If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < _
                  MAX_INVENTORY_SLOTS + 1) Then Call _
                  WriteEquipItem(Inventario.SelectedItem)
90        End If
End Sub



Private Sub Timer1_Timer()
    LoopInterval
  '  If frmMain.Visible Then RandomMove
End Sub





Private Sub TimerF8_Timer()
    CantKey0 = 0
    CantKey1 = 0
    CantKey2 = 0
    CantF8 = 0
End Sub

Private Sub TimerPing_Timer()

    Static i As Integer
          '//
10        i = i + 1
20        If pingTime = 0 Then Exit Sub
          
30        If i >= 3 Then
40            i = 0
50            pingTime = 0
60        End If
End Sub

Private Sub tPotas_Timer()
'   UsarItem 0
  '  UsarItem 1
End Sub

''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''

Private Sub TrainingMacro_Timer()
10        If Not hlst.Visible Then
20            DesactivarMacroHechizos
30            Exit Sub
40        End If
          
          'Macros are disabled if focus is not on Argentum!
50        If Not Application.IsAppActive() Then
60            DesactivarMacroHechizos
70            Exit Sub
80        End If
          
90        If Comerciando Then Exit Sub
100       If Canjeando Then Exit Sub
          
110      If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.CastSpell, False) Then
        Call WriteCastSpell(hlst.ListIndex + 1)
        Call WriteWork(eSkill.Magia)
    End If
          
150       Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
          
          'If UsingSkill = Magia And Not MainTimer.Check(TimersIndex.CastSpell) Then Exit Sub
          
          'If UsingSkill = Proyectiles And Not MainTimer.Check(TimersIndex.Attack) Then Exit Sub
          
160       Call WriteWorkLeftClick(tX, tY, UsingSkill)
170       UsingSkill = 0
End Sub

Private Sub cmdLanzar_Click()

10     If hlst.List(hlst.ListIndex) = "(Vacio)" Then Exit Sub
        
20        If Iscombate = False Then
30       With FontTypes(FontTypeNames.FONTTYPE_INFO)
40        Call _
              ShowConsoleMsg("¡¡No puedes lanzar hechizos si no estas en modo combate!!", _
              .red, .green, .blue, .bold, .italic)
50       End With
60        Exit Sub
70        End If
        


80        If hlst.List(hlst.ListIndex) <> "(None)" And _
              MainTimer.Check(TimersIndex.Work, False) Then
90            If UserEstado = 1 Then
100               With FontTypes(FontTypeNames.FONTTYPE_INFO)
110                   Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, _
                          .bold, .italic)
120               End With
130           Else
                Call modAnalisis.ClickLanzar
140               Call WriteCastSpell(hlst.ListIndex + 1)
150               Call WriteWork(eSkill.Magia)
160               UsaMacro = True
170           End If
180       End If


End Sub


Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        UsaMacro = False
20        CnTd = 0
End Sub

Private Sub cmdINFO_Click()
10        If hlst.ListIndex <> -1 Then
20            Call WriteSpellInfo(hlst.ListIndex + 1)
30        End If
End Sub

Private Sub DespInv_Click(Index As Integer)
10        Inventario.ScrollInventory (Index = 0)
End Sub
Private Sub PicInv_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        If Not Inventario.uMoveItem Then
20            picInv.MousePointer = vbDefault
30        End If
End Sub
Private Sub Form_Click()
20        If Not Comerciando Then
30            Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
              'CharFichado = MapData(tX, tY).CharIndex
              
40            If MouseShift = 0 Then
50                If MouseBoton <> vbRightButton Then
                      '[ybarra]
60                    If UsaMacro Then
70                        CnTd = CnTd + 1
80                        If CnTd = 3 Then
90                            Call WriteUseSpellMacro
100                           CnTd = 0
110                       End If
120                       UsaMacro = False
130                   End If
                      '[/ybarra]
140                   If UsingSkill = 0 Then
150                       Call WriteLeftClick(tX, tY)
160                   Else
                      
170                       If TrainingMacro.Enabled Then Call DesactivarMacroHechizos
180                       If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
                          
190                      If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
200                           frmMain.MousePointer = vbDefault
210                           UsingSkill = 0
220                           With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                '  Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .red, .green, .blue, .bold, .italic)
230                           End With
240                           Exit Sub
250                       End If
                          
                          'Splitted because VB isn't lazy!
260                       If UsingSkill = Proyectiles Then
270                           If Not MainTimer.Check(TimersIndex.Arrows) Then
280                               frmMain.MousePointer = vbDefault
290                               UsingSkill = 0
300                               With FontTypes(FontTypeNames.FONTTYPE_TALK)
                               '       Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .red, .green, .blue, .bold, .italic)
310                               End With
320                               Exit Sub
330                           End If
340                       End If
                          
                          'Splitted because VB isn't lazy!
350                       If UsingSkill = Magia Then
360                           If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
370                               If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
380                                   frmMain.MousePointer = vbDefault
390                                   UsingSkill = 0
400                                   With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                        '  Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rápido.", .red, .green, .blue, .bold, .italic)
410                                   End With
420                                   Exit Sub
430                               End If
440                           Else
450                               If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
460                                   frmMain.MousePointer = vbDefault
470                                   UsingSkill = 0
480                                   With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                         ' Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rapido.", .red, .green, .blue, .bold, .italic)
490                                   End With
500                                   Exit Sub
510                               End If
520                           End If
530                       End If
                          
                          'Splitted because VB isn't lazy!
540                       If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill _
                              = Talar Or UsingSkill = Mineria Or UsingSkill = _
                              FundirMetal) Then
550                           If Not MainTimer.Check(TimersIndex.Work) Then
560                               frmMain.MousePointer = vbDefault
570                               UsingSkill = 0
580                               Exit Sub
590                           End If
600                       End If
                          
                          'If frmMain.MousePointer <> 2 Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                          
610                       If frmMain.MousePointer <> 2 Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                          
620                       frmMain.MousePointer = vbDefault
630                       Call WriteWorkLeftClick(tX, tY, UsingSkill)
640                       UsingSkill = 0
650                   End If
660               Else
                      ' Descastea
670                   If UsingSkill = Magia Or UsingSkill = Proyectiles Then
680                       frmMain.MousePointer = vbDefault
690                       UsingSkill = 0
700                   Else
                          
710                       If Config.ClickDerecho Then
720                           Call WriteRightClick(tX, tY)
730                       End If
740                   End If
750               End If
                  
If MouseBoton = vbRightButton Then
                    Call WriteWarpChar("YO", UserMap, tX, tY)
                End If
            End If
        End If
End Sub
Private Sub Form_DblClick()
      '**************************************************************
      'Author: Unknown
      'Last Modify Date: 12/27/2007
      '12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
      '**************************************************************
10        If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
20            Call WriteDoubleClick(tX, tY)
30        End If
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
10           KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
10           KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
10            KeyCode = 0
End Sub

Private Sub lblDropGold_Click()

10        Inventario.SelectGold
20        If UserGLD > 0 Then
30            If Not Comerciando Then FrmCantidad.Show , frmMain
40        End If
          
End Sub

Private Sub Label4_Click()

    Dim Pt As Point
10        Call Audio.PlayWave(SND_CLICK)

20        InvEqu.Picture = LoadPicture(App.path & "\Recursos\Centroinventario.JPG")

30        Panel = eVentanas.vInventario

          GetCursorPos Pt
          
40                                                  'If Panel <> LastPanel Then
50        Call WriteSetMenu(Panel, 255, Pt.X, Pt.Y)
60        LastPanel = Panel
70                                                  'End If
            modAnalisis.ClickCambioInv
          ' Activo controles de inventario

80        picInv.Visible = True
90        IconosegD.Visible = True
100       IconoSeg.Visible = True
110       Label6.Visible = True
          
          'imgInvScrollUp.Visible = True
          'imgInvScrollDown.Visible = True

          ' Desactivo controles de hechizo
120       hlst.Visible = False
130       cmdINFO.Visible = False
140       CmdLanzar.Visible = False
          
150       cmdMoverHechi(0).Visible = False
160       cmdMoverHechi(1).Visible = False
          
End Sub
Private Sub label4_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        LastPressed.ToggleToNormal
20            Inventario.uMoveItem = False
30        Inventario.sMoveItem = False
End Sub

Private Sub Label7_Click()
10        Call Audio.PlayWave(SND_CLICK)

20        InvEqu.Picture = LoadPicture(App.path & "\Recursos\Centrohechizos.JPG")
          
30        Panel = eVentanas.vHechizos

40        'If Panel <> LastPanel Then

          Dim TempInv As Byte
          Dim Pt As Point
          
          
          GetCursorPos Pt
          
50        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < _
                MAX_INVENTORY_SLOTS + 1) Then
60              TempInv = CByte(Inventario.SelectedItem)
70        Else
80              TempInv = 255 ' @@ Pasamos y tenemos ningun slot seleccionado entonces 255 ...

90        End If
            
              
100       Call WriteSetMenu(Panel, TempInv, Pt.X, Pt.Y)
110       LastPanel = Panel

120       'End If

          modAnalisis.ClickCambioHech
          
          ' Activo controles de hechizos
130       hlst.Visible = True
140       cmdINFO.Visible = True
150       CmdLanzar.Visible = True
          
160       cmdMoverHechi(0).Visible = True
170       cmdMoverHechi(1).Visible = True
          
          ' Desactivo controles de inventario
180       picInv.Visible = False
190       IconosegD.Visible = False
200       IconoSeg.Visible = False
210       Label6.Visible = False
          'imgInvScrollUp.Visible = False
          'imgInvScrollDown.Visible = False

End Sub
Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        LastPressed.ToggleToNormal
20            Inventario.uMoveItem = False
30        Inventario.sMoveItem = False
End Sub

Private Sub picInv_DblClick()
    Call Audio.PlayWave(SND_CLICK)
10  If (mouse_Down <> False) And (mouse_UP = True) Then Exit Sub
    If (timeGetTime - Intervalos(eInterval.iUseItemClick).ModifyTime) <= 200 Then Exit Sub
          
          

    
20        mouse_UP = False
          ' x button
          
30        If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
          
40        If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
          
50        If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo

60            Inventario.uMoveItem = False
              
70            If MouseInvBoton = vbRightButton Then Exit Sub
          
          

80            Call UsarItem(1)


        'ShowConsoleMsg "Uso Click"
End Sub

Private Sub PicInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
                                 
          '    / x button
10        If (mouse_Down = False) Then Exit Sub
20        mouse_Down = False
30        mouse_UP = True
          '    / x button
          
40        Call Audio.PlayWave(SND_CLICK)
50        Inventario.uMoveItem = False
60        MouseInvBoton = Button
End Sub

Private Sub RecTxt_Change()
10    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
20        If Not Application.IsAppActive() Then Exit Sub
          
30       If SendTxt.Visible Then
40            SendTxt.SetFocus
50      ElseIf Me.SendRmstxt.Visible Then
60            SendRmstxt.SetFocus
70        ElseIf Me.SendGms.Visible Then
80            SendGms.SetFocus
90            ElseIf SendCMSTXT.Visible Then
100           SendCMSTXT.SetFocus
110       ElseIf (Not Comerciando) And (Not Canjeando) And (Not MirandoAsignarSkills) _
              And (Not frmMSG.Visible) And (Not MirandoForo) And (Not _
              frmEstadisticas.Visible) And (Not FrmCantidad.Visible) Then
               
120           If picInv.Visible Then
130               picInv.SetFocus
140           ElseIf hlst.Visible Then
150               hlst.SetFocus
160           End If
170       End If
End Sub
Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
10        If picInv.Visible Then
20            picInv.SetFocus
30        Else
40            hlst.SetFocus
50        End If
End Sub
Private Function InGameArea() As Boolean
      '***************************************************
      'Author: NicoNZ
      'Last Modification: 04/07/08
      'Checks if last click was performed within or outside the game area.
      '***************************************************
10        If clicX < MainViewPic.Left Or clicX > MainViewPic.Left + MainViewPic.Width _
              Then Exit Function
20        If clicY < MainViewPic.Top Or clicY > MainViewPic.Top + MainViewPic.Height _
              Then Exit Function
          
30        InGameArea = True
End Function
Private Sub SendTxt_Change()
      '**************************************************************
      'Author: Unknown
      'Last Modify Date: 3/06/2006
      '3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
      '**************************************************************

10    If Pulsacion_Fisica = False Then
20    Exit Sub
30    End If
40    Pulsacion_Fisica = True

50        If Len(SendTxt.Text) > 160 Then
60            stxtbuffer = ""
70        Else
              'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
              Dim i As Long
              Dim TempStr As String
              Dim CharAscii As Integer
              
80            For i = 1 To Len(SendTxt.Text)
90                CharAscii = Asc(mid$(SendTxt.Text, i, 1))
100               If CharAscii >= vbKeySpace And CharAscii <= 250 Then
110                   TempStr = TempStr & Chr$(CharAscii)
120               End If
130           Next i
              
140           If TempStr <> SendTxt.Text Then
                  'We only set it if it's different, otherwise the event will be raised
                  'constantly and the client will crush
150               SendTxt.Text = TempStr
160           End If
              
170           stxtbuffer = SendTxt.Text
180           frmMain.SendTxt.SetFocus
190       End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
10        If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii _
              <= 250) Then KeyAscii = 0
End Sub


''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
#If UsarWrench = 1 Then

Private Sub Socket1_Connect()
          'Clean input and output buffers
          
10        Second.Enabled = True

        Select Case EstadoLogin
            Case E_MODO.BorrarPJ
                FrmRECBORR.Show vbModal
            
            Case E_MODO.RecuperarPJ
                FrmRECBORR.Show vbModal
            
            Case E_MODO.CrearNuevoPj
                Call Login
        
            Case E_MODO.Normal
                Call Login
        
            Case E_MODO.Dados
                frmCrearPersonaje.Show vbModal
        End Select
End Sub

Private Sub Socket1_Disconnect()
          Dim i As Long
          ClaveActual = 0
          
10        Second.Enabled = False
20        Connected = False
          
30        Socket1.Cleanup
          
40        frmConnect.MousePointer = vbNormal
          
50        Do While i < Forms.Count - 1
60            i = i + 1
              
70            If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name And _
                  Forms(i).Name <> frmCrearPersonaje.Name Then
80                Unload Forms(i)
90            End If
100       Loop
          
110       On Local Error GoTo 0
          
120       If Not frmCrearPersonaje.Visible Then
130           frmConnect.Visible = True
140       End If
          
150       frmMain.Visible = False
          
160       pausa = False
170       UserMeditar = False

180       UserClase = 0
190       UserSexo = 0
200       UserRaza = 0
210       UserHogar = 0
220       UserEmail = ""


          ResetKeyPackets
          
          
230       For i = 1 To NUMSKILLS
240           UserSkills(i) = 0
250       Next i

260       For i = 1 To NUMATRIBUTOS
270           UserAtributos(i) = 0
280       Next i
          
290       For i = 1 To MAX_INVENTORY_SLOTS
              
300       Next i
          
310       macrotrabajo.Enabled = False

320       SkillPoints = 0
330       Alocados = 0
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, _
    Response As Integer)
          '*********************************************
          'Handle socket errors
          '*********************************************
10        If ErrorCode = 24036 Then
20            Call MsgBox("Por favor espere, intentando completar conexion.", _
                  vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, _
                  "Error")
30            Exit Sub
40        End If
          
50        Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + _
              vbDefaultButton1, "Error")
60        frmConnect.MousePointer = 1
70        Response = 0
80        Second.Enabled = False

90        frmMain.Socket1.Disconnect
          
100       If Not frmCrearPersonaje.Visible Then
110           frmConnect.Show
120       Else
130           frmCrearPersonaje.MousePointer = 0
140       End If
End Sub

Private Sub Socket1_Read(dataLength As Integer, IsUrgent As Integer)
          Dim RD As String
          Dim data() As Byte
          
10        Call Socket1.Read(RD, dataLength)


20        data = StrConv(RD, vbFromUnicode)
         
            
30        If Len(RD) = 0 Then Exit Sub

          'Put data in the buffer
40        Call incomingData.WriteBlock(data)
          
          'Send buffer to Handle data
50        Call HandleIncomingData
End Sub


#End If

Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

10    If tX >= MinXBorder And tY >= MinYBorder And tY <= MaxYBorder And tX <= _
          MaxXBorder Then
20        If MapData(tX, tY).CharIndex > 0 Then
30            If charlist(MapData(tX, tY).CharIndex).Invisible = False Then
              
                  Dim i As Long
                  Dim m As New frmMenuseFashion
                  
40                Load m
50                m.SetCallback Me
60                m.SetMenuId 1
70                m.ListaInit 2, False
                  
80                If charlist(MapData(tX, tY).CharIndex).Nombre <> "" Then
90                    m.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).Nombre, _
                          True
100               Else
110                   m.ListaSetItem 0, "<NPC>", True
120               End If
130               m.ListaSetItem 1, "Comerciar"
                  
140               m.ListaFin
150               m.Show , Me

160           End If
170       End If
180   End If

#End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
10    Select Case MenuId

      Case 0 'Inventario
20        Select Case Sel
          Case 0
30        Case 1
40        Case 2 'Tirar
50            Call TirarItem
60        Case 3 'Usar
70            If MainTimer.Check(TimersIndex.UseItemWithDblClick) Then
80                Call UsarItem(1)
90            End If
100       Case 3 'equipar
110           Call EquiparItem
120       End Select
          
130   Case 1 'Menu del ViewPort del engine
140       Select Case Sel
          Case 0 'Nombre
150           Call WriteLeftClick(tX, tY)
              
160       Case 1 'Comerciar
170           Call WriteLeftClick(tX, tY)
180           Call WriteCommerceStart
190       End Select
200   End Select
End Sub


Private Sub tSec_Timer()
          Static Contador As Byte
          Static Count As Byte
          
10        Contador = Contador + 1
          
20        If Contador >= 5 Then
30            Contador = 0
40            CheckPrincipales False
50        End If
          
60        If Count >= 15 Then
70            Count = 0
80            CheckPrincipales True
90        End If
End Sub

'
' -------------------
'    W I N S O C K
' -------------------
'

#If UsarWrench <> 1 Then

Private Sub Winsock1_Close()
          Dim i As Long
          
10        Debug.Print "WInsock Close"
          
20        Second.Enabled = False
30        Connected = False
          
40        If Winsock1.State <> sckClosed Then Winsock1.Close
          
50        frmConnect.MousePointer = vbNormal
          
60        Do While i < Forms.Count - 1
70            i = i + 1
              
80            If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name And _
                  Forms(i).Name <> frmCrearPersonaje.Name Then
90                Unload Forms(i)
100           End If
110       Loop
120       On Local Error GoTo 0
          
130       If Not frmCrearPersonaje.Visible Then
140           frmConnect.Visible = True
150       End If
          
160       frmMain.Visible = False

170       pausa = False
180       UserMeditar = False

190       UserClase = 0
200       UserSexo = 0
210       UserRaza = 0
220       UserHogar = 0
230       UserEmail = ""
          
240       For i = 1 To NUMSKILLS
250           UserSkills(i) = 0
260       Next i

270       For i = 1 To NUMATRIBUTOS
280           UserAtributos(i) = 0
290       Next i

300       SkillPoints = 0
310       Alocados = 0

320       Dialogos.CantidadDialogos = 0
End Sub

Private Sub Winsock1_Connect()
10        Debug.Print "Winsock Connect"
          
          'Clean input and output buffers
20        Call incomingData.ReadASCIIStringFixed(incomingData.Length)
30        Call outgoingData.ReadASCIIStringFixed(outgoingData.Length)

          
40        Second.Enabled = True
          
50        Select Case EstadoLogin
          Case E_MODO.BorrarPJ
60           FrmRECBORR.Show vbModal
70            Case E_MODO.RecuperarPJ
80           FrmRECBORR.Show vbModal
90            Case E_MODO.CrearNuevoPj
100               Call Login
                  
110   Case E_MODO.BorrarPersonaje
120               Call Login
                  

130           Case E_MODO.Normal
140               Call Login

150           Case E_MODO.Dados
160               Call Audio.PlayMIDI("7.mid")
170               frmCrearPersonaje.Show vbModal
                  
180       End Select
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
          Dim RD As String
          Dim data() As Byte
          
          'Socket1.Read RD, DataLength
10        Winsock1.GetData RD

20        data = StrConv(RD, vbFromUnicode)
          
          'Set data in the buffer
30        Call incomingData.WriteBlock(data)
          
          'Send buffer to Handle data
40        Call HandleIncomingData
End Sub

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, _
    ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal _
    HelpContext As Long, CancelDisplay As Boolean)
          '*********************************************
          'Handle socket errors
          '*********************************************
          
10        Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + _
              vbDefaultButton1, "Error")
20        frmConnect.MousePointer = 1
30        Second.Enabled = False

40        If Winsock1.State <> sckClosed Then Winsock1.Close

50        If Not frmCrearPersonaje.Visible Then
60            frmConnect.Show
70        Else
80            frmCrearPersonaje.MousePointer = 0
90        End If
End Sub
#End If

Public Sub DesactivarMacroHechizos()
10        TrainingMacro.Enabled = False
20        Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos desactivado", 0, _
              150, 150, False, True, True)
30        Call ControlSM(eSMType.mSpells, False)
End Sub

Private Sub PicInv_MouseDown(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
      Dim Position As Integer
      Dim i As Long
      Dim file_path As String
      Dim data() As Byte
      Dim bmpInfo As BITMAPINFO
      Dim handle As Integer
      Dim bmpData As StdPicture

          '    / x button
10        mouse_Down = True
20        mouse_UP = False
          '    / x button

30    If (Button = vbRightButton) And (Not Comerciando) And (Not Canjeando) Then

40    If Inventario.GrhIndex(Inventario.SelectedItem) < 1 Then
        'Call MsgBox("Primero debes seleccionar un item de tu inventario.", vbCritical + vbOKOnly)
50      Exit Sub
60    End If
        
70    If Inventario.GrhIndex(Inventario.SelectedItem) > 0 Then

80            Last_I = Inventario.SelectedItem
90            If Last_I > 0 And Last_I <= MAX_INVENTORY_SLOTS Then
                          
100               Position = Search_GhID(Inventario.GrhIndex(Inventario.SelectedItem))
                  
110               If Position = 0 Then
120                   i = _
                          GrhData(Inventario.GrhIndex(Inventario.SelectedItem)).FileNum
130                   Call Get_Bitmapp(DirGraficos, _
                          CStr(GrhData(Inventario.GrhIndex(Inventario.SelectedItem)).FileNum) _
                          & ".BMP", bmpInfo, data)
140                   Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1) ' GSZAO ' GSZAO
150                   frmMain.ImageList1.ListImages.Add , CStr("g" & _
                          Inventario.GrhIndex(Inventario.SelectedItem)), Picture:=bmpData
160                   Position = frmMain.ImageList1.ListImages.Count
170                   Set bmpData = Nothing
180               End If
                  
                  
190               Inventario.uMoveItem = True
                  
200               Set picInv.MouseIcon = _
                      frmMain.ImageList1.ListImages(Position).ExtractIcon
210               frmMain.picInv.MousePointer = vbCustom

220               Exit Sub
230           End If
240       End If
250   End If
End Sub

Private Function Search_GhID(gh As Integer) As Integer

      Dim i As Integer

10    For i = 1 To frmMain.ImageList1.ListImages.Count
20        If frmMain.ImageList1.ListImages(i).key = "g" & CStr(gh) Then
30            Search_GhID = i
40            Exit For
50        End If
60    Next i

End Function

Public Sub dragInventory_dragDone(ByVal originalSlot As Integer, ByVal newSlot _
    As Integer)
10    Call Protocol.WriteDragInventory(originalSlot, newSlot, eMoveType.Inventory)
20    Inventario.uMoveItem = False
30    Inventario.sMoveItem = False
End Sub
Private Sub Label2_Click()
10    If UserLvl < 47 Then
20    Call ShowConsoleMsg("Nivel: " & UserLvl & " Experiencia: " & Format$(UserExp, _
          "#,###") & "/" & Format$(UserPasarNivel, "#,###") & " " & "(" & _
          Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%)", 0, 240, 240)
30    Else
40    Call AddtoRichTextBox(frmMain.RecTxt, "Nivel: " & UserLvl & _
          " ^\\\\\\Máximo///////^", 0, 200, 200, False, False, True)
50    End If
End Sub

Private Sub label2_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        lvllbl(0).Visible = False
20        lvllbl(1).Visible = False
30        lvllbl(2).Visible = False
40        lvllbl(3).Visible = False
50        lvllbl(4).Visible = False
60        lblporclvl(0).Visible = True
70        lblporclvl(1).Visible = True
80        lblporclvl(2).Visible = True
90        lblporclvl(3).Visible = True
100       lblporclvl(4).Visible = True
End Sub

Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
          'Send text
10        If KeyCode = vbKeyReturn Then
              'Say
20            If stxtbuffercmsg <> "" Then
30                Call ParseUserCommand("/CMSG " & stxtbuffercmsg)
40            End If

50            stxtbuffercmsg = ""
60            SendCMSTXT.Text = ""
70            KeyCode = 0
80            Me.SendCMSTXT.Visible = False
              
90            If picInv.Visible Then
100               picInv.SetFocus
110           Else
120               hlst.SetFocus
130           End If
140       End If
End Sub

Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
10        If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii _
              <= 250) Then KeyAscii = 0
End Sub

Private Sub SendCMSTXT_Change()
10        If Len(SendCMSTXT.Text) > 160 Then
20            stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
30        Else
              'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
              Dim i As Long
              Dim TempStr As String
              Dim CharAscii As Integer
              
40            For i = 1 To Len(SendCMSTXT.Text)
50                CharAscii = Asc(mid$(SendCMSTXT.Text, i, 1))
60                If CharAscii >= vbKeySpace And CharAscii <= 250 Then
70                    TempStr = TempStr & Chr$(CharAscii)
80                End If
90            Next i
              
100           If TempStr <> SendCMSTXT.Text Then
                  'We only set it if it's different, otherwise the event will be raised
                  'constantly and the client will crush
110               SendCMSTXT.Text = TempStr
120           End If
              
130           stxtbuffercmsg = SendCMSTXT.Text
140       End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
10    If Not GetAsyncKeyState(KeyCode) < 0 Then
20    Pulsacion_Fisica = False
30    Exit Sub
40    End If
50    Pulsacion_Fisica = True
End Sub
Private Sub SendRMSTXT_Change()
10        stxtbufferrmsg = SendRmstxt.Text
End Sub
Private Sub sendgms_change()
10    stxtbufferrmsg = SendGms.Text
End Sub
Private Sub SendRMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
          'Send text
10        If KeyCode = vbKeyReturn Then
              'Say
20            If stxtbufferrmsg <> "" Then
30                Call ParseUserCommand("/RMSG " & stxtbufferrmsg)
40            End If
             ' frmMain.Label2 = ""
50            stxtbufferrmsg = ""
60            SendRmstxt.Text = ""
70            KeyCode = 0
80            Me.SendRmstxt.Visible = False
90        End If
End Sub

Private Sub SendRMSTXT_KeyPress(KeyAscii As Integer)
10        If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii _
              <= 250) Then KeyAscii = 0
End Sub
Private Sub SendGms_KeyUp(KeyCode As Integer, Shift As Integer)
          'Send text
10        If KeyCode = vbKeyReturn Then
              'Say
20            If stxtbufferrmsg <> "" Then
30                Call ParseUserCommand("/GMSG " & stxtbufferrmsg)
40            End If
             ' frmMain.Label2 = ""
50            stxtbufferrmsg = ""
60            SendGms.Text = ""
70            KeyCode = 0
80            Me.SendGms.Visible = False
90        End If
End Sub

Private Sub SendGms_KeyPress(KeyAscii As Integer)
10        If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii _
              <= 250) Then KeyAscii = 0
End Sub
