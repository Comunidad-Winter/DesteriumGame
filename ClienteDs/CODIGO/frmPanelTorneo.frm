VERSION 5.00
Begin VB.Form frmPanelTorneo 
   Caption         =   "TORNEOS"
   ClientHeight    =   10200
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form4"
   ScaleHeight     =   946.088
   ScaleMode       =   0  'User
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      BackColor       =   &H80000009&
      Caption         =   "MAS..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2160
      Index           =   1
      Left            =   5940
      TabIndex        =   60
      Top             =   7425
      Width           =   6060
      Begin VB.CommandButton Command4 
         Caption         =   "CLICKEAME PARA INVASIONES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -390
         TabIndex        =   61
         Top             =   480
         Width           =   5475
      End
   End
   Begin VB.PictureBox PicInvasiones 
      Height          =   7230
      Left            =   15
      ScaleHeight     =   7170
      ScaleWidth      =   11070
      TabIndex        =   62
      Top             =   7215
      Visible         =   0   'False
      Width           =   11130
      Begin VB.CommandButton Command5 
         Caption         =   "CERRAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4245
         TabIndex        =   63
         Top             =   795
         Width           =   1380
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         Caption         =   "Invasiones automáticas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   5280
         Left            =   0
         TabIndex        =   64
         Top             =   0
         Width           =   5865
         Begin VB.TextBox txtDesc 
            Alignment       =   1  'Right Justify
            Height          =   795
            Left            =   1170
            TabIndex        =   70
            Text            =   "Ohh.. Una increíble cantidad de mutantes ha renacido. ¡Acabad con ellos!"
            Top             =   780
            Width           =   2760
         End
         Begin VB.TextBox txtInvasion 
            Height          =   285
            Left            =   1365
            TabIndex        =   73
            Text            =   "1"
            Top             =   1755
            Width           =   795
         End
         Begin VB.CommandButton Command3 
            Caption         =   "SOLICITAR LISTA DE INVASIONES [AYUDA GM]"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   390
            TabIndex        =   72
            Top             =   3315
            Width           =   4110
         End
         Begin VB.TextBox txtMap 
            Height          =   285
            Left            =   1170
            TabIndex        =   67
            Text            =   "1"
            Top             =   2340
            Width           =   795
         End
         Begin VB.CommandButton Command2 
            Caption         =   "ABRIR"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4305
            TabIndex        =   66
            Top             =   255
            Width           =   1290
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   1170
            TabIndex        =   65
            Text            =   ">INVASIÓN MUTANTE"
            Top             =   390
            Width           =   2355
         End
         Begin VB.Label Label22 
            Caption         =   "Acá eligan el mapa que ustedes quieran.        1 Es ullathorpe."
            ForeColor       =   &H0000FF00&
            Height          =   495
            Left            =   2025
            TabIndex        =   78
            Top             =   2250
            Width           =   3315
         End
         Begin VB.Label Label16 
            Caption         =   "*OJO*, solo hay dateada 1 SOLA INVASIÓN. Pongan 1 por ahora."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   465
            Left            =   2220
            TabIndex        =   77
            Top             =   1710
            Width           =   3645
         End
         Begin VB.Line Line 
            X1              =   195
            X2              =   5265
            Y1              =   2925
            Y2              =   2925
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Invasión:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   405
            TabIndex        =   71
            Top             =   1740
            Width           =   1380
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Mapa:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   390
            TabIndex        =   69
            Top             =   2340
            Width           =   600
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   390
            TabIndex        =   68
            Top             =   390
            Width           =   795
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Información de los eventos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   6450
      Left            =   6240
      TabIndex        =   30
      Top             =   0
      Width           =   4890
      Begin VB.Frame Frame3 
         Caption         =   "Información del evento seleccionado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4500
         Left            =   195
         TabIndex        =   35
         Top             =   1755
         Width           =   4305
         Begin VB.ListBox lstUsers 
            Height          =   3180
            Left            =   195
            TabIndex        =   40
            Top             =   1170
            Width           =   3720
         End
         Begin VB.Label lblUsers 
            Caption         =   "Usuarios inscriptos y DISPONIBLES:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   195
            TabIndex        =   41
            Top             =   975
            Width           =   3330
         End
         Begin VB.Label lblDspCurso 
            Caption         =   "DSP- Poso acumulado:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   195
            TabIndex        =   39
            Top             =   780
            Width           =   3330
         End
         Begin VB.Label lblOroCurso 
            Caption         =   "ORO- Poso acumulado:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   195
            TabIndex        =   38
            Top             =   585
            Width           =   3330
         End
         Begin VB.Label lblNivelCurso 
            Caption         =   "Nivel mínimo/máximo:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   195
            TabIndex        =   37
            Top             =   390
            Width           =   3330
         End
         Begin VB.Label lblQuotasCurso 
            Caption         =   "Cupos:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   195
            TabIndex        =   36
            Top             =   195
            Width           =   3330
         End
      End
      Begin VB.ComboBox cmbModalityCurso 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2145
         TabIndex        =   33
         Text            =   "Vacio"
         Top             =   1365
         Width           =   2355
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Solicitar eventos en curso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   390
         TabIndex        =   32
         Top             =   195
         Width           =   3915
      End
      Begin VB.Label lblClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   4545
         TabIndex        =   42
         Top             =   1170
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label11 
         Caption         =   "Una vez que aparezca la lista de los eventos que hay disponibles (en curso) al seleccionar uno se actualizarán sus datos."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   600
         Left            =   195
         TabIndex        =   34
         Top             =   585
         Width           =   4305
      End
      Begin VB.Label Label10 
         Caption         =   "Eventos disponibles:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   31
         Top             =   1365
         Width           =   1770
      End
   End
   Begin VB.CommandButton cmbCrear 
      Caption         =   "OK!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3510
      TabIndex        =   9
      Top             =   6240
      Width           =   2160
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nuevo torneo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   7230
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6060
      Begin VB.CheckBox chkFaccion 
         Caption         =   "Armada Real"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   3510
         TabIndex        =   76
         Top             =   3900
         Value           =   1  'Checked
         Width           =   2160
      End
      Begin VB.CheckBox chkFaccion 
         Caption         =   "Legión Oscura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   3510
         TabIndex        =   75
         Top             =   3705
         Value           =   1  'Checked
         Width           =   2160
      End
      Begin VB.CheckBox chkFaccion 
         Caption         =   "Ciudadano"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   3510
         TabIndex        =   74
         Top             =   3510
         Value           =   1  'Checked
         Width           =   2160
      End
      Begin VB.Frame Frame 
         Height          =   15
         Index           =   0
         Left            =   0
         TabIndex        =   59
         Top             =   7215
         Width           =   5670
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "Valen items [ORO]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   195
         TabIndex        =   58
         Top             =   6240
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.TextBox txtObjPremio 
         Height          =   285
         Left            =   1950
         TabIndex        =   57
         Text            =   "0"
         Top             =   5850
         Width           =   795
      End
      Begin VB.ComboBox cmbOroPremio 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   56
         Text            =   "Vacio"
         Top             =   5460
         Width           =   1185
      End
      Begin VB.ComboBox cmbDspPremio 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   55
         Text            =   "Vacio"
         Top             =   5070
         Width           =   1185
      End
      Begin VB.TextBox txtRojas 
         Height          =   285
         Left            =   1950
         TabIndex        =   51
         Text            =   "0"
         Top             =   4680
         Width           =   795
      End
      Begin VB.CheckBox chkAcum 
         Caption         =   "Dar poso acumulado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   195
         TabIndex        =   49
         Top             =   2730
         Width           =   2745
      End
      Begin VB.CheckBox chkFaccion 
         Caption         =   "Criminal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   3510
         TabIndex        =   48
         Top             =   3315
         Value           =   1  'Checked
         Width           =   2160
      End
      Begin VB.Frame FrameDuelos 
         Caption         =   "Opciones de enfrentamientos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1380
         Left            =   3315
         TabIndex        =   43
         Top             =   4485
         Width           =   2550
         Begin VB.CheckBox chkGanadorSigue 
            Caption         =   "Ganador sigue"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   195
            TabIndex        =   46
            Top             =   780
            Visible         =   0   'False
            Width           =   1770
         End
         Begin VB.ComboBox cmbTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1300
            TabIndex        =   45
            Text            =   "Vacio"
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label12 
            Caption         =   "Cantidad por TEAM:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   130
            TabIndex        =   44
            Top             =   390
            Width           =   1380
         End
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Pirata"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   11
         Left            =   3510
         TabIndex        =   29
         Top             =   2340
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   3510
         TabIndex        =   28
         Top             =   2145
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Cazador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   3510
         TabIndex        =   27
         Top             =   1950
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Paladin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   3510
         TabIndex        =   26
         Top             =   1755
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Druida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   3510
         TabIndex        =   25
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Bardo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   3510
         TabIndex        =   24
         Top             =   1365
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Ladrón"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   3510
         TabIndex        =   23
         Top             =   2535
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Asesino"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   3510
         TabIndex        =   22
         Top             =   1170
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Guerrero"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   3510
         TabIndex        =   21
         Top             =   975
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Clerigo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   3510
         TabIndex        =   20
         Top             =   780
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Mago"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   3510
         TabIndex        =   19
         Top             =   585
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.ComboBox cmbInit 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   17
         Text            =   "Vacio"
         Top             =   4095
         Width           =   1185
      End
      Begin VB.ComboBox cmbCancel 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   16
         Text            =   "Vacio"
         Top             =   3510
         Width           =   1185
      End
      Begin VB.ComboBox cmbDsp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   15
         Text            =   "Vacio"
         Top             =   2280
         Width           =   1185
      End
      Begin VB.ComboBox cmbOro 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   14
         Text            =   "Vacio"
         Top             =   1880
         Width           =   1185
      End
      Begin VB.ComboBox cmbLvlMax 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   13
         Text            =   "Vacio"
         Top             =   1500
         Width           =   1185
      End
      Begin VB.ComboBox cmbLvlMin 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   12
         Text            =   "Vacio"
         Top             =   1120
         Width           =   1185
      End
      Begin VB.ComboBox cmbQuotas 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   11
         Text            =   "Vacio"
         Top             =   730
         Width           =   1185
      End
      Begin VB.ComboBox cmbModality 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   10
         Text            =   "Vacio"
         Top             =   350
         Width           =   1185
      End
      Begin VB.Label Label20 
         Caption         =   "Objeto Premio:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   54
         Top             =   5850
         Width           =   1380
      End
      Begin VB.Label Label19 
         Caption         =   "Oro Premio:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   53
         Top             =   5460
         Width           =   1185
      End
      Begin VB.Label Label18 
         Caption         =   "Dsp Premio:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   52
         Top             =   5070
         Width           =   1185
      End
      Begin VB.Label Label17 
         Caption         =   "Limite de rojas:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   50
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Facciones permitidas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   3510
         TabIndex        =   47
         Top             =   2925
         Width           =   1965
      End
      Begin VB.Label Label9 
         Caption         =   "Clases permitidas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   3510
         TabIndex        =   18
         Top             =   195
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   3315
         X2              =   3315
         Y1              =   390
         Y2              =   4290
      End
      Begin VB.Label Label8 
         Caption         =   "Tiempo para comenzar:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   195
         TabIndex        =   8
         Top             =   3900
         Width           =   1380
      End
      Begin VB.Label Label7 
         Caption         =   "Tiempo para cancelar:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   195
         TabIndex        =   7
         Top             =   3315
         Width           =   1380
      End
      Begin VB.Label Label6 
         Caption         =   "Dsp inscripción:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   6
         Top             =   2340
         Width           =   1380
      End
      Begin VB.Label Label5 
         Caption         =   "Oro inscripción:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   5
         Top             =   1950
         Width           =   1380
      End
      Begin VB.Label Label4 
         Caption         =   "Nivel máximo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   4
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Nivel mínimo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   3
         Top             =   1170
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "Cupos:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   2
         Top             =   780
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Modalidad:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   1
         Top             =   390
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmPanelTorneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbCrear_Click()
          
          
          Dim AllowedClasses(1 To NUMCLASES) As Byte
          Dim AllowedFaction(1 To 4) As Byte
          
          Dim LoopC As Byte
          
10        If CheckForm(AllowedClasses) Then
                ' Clases válidas
20            For LoopC = 1 To NUMCLASES
30                AllowedClasses(LoopC) = Val(chkClass(LoopC).value)
40            Next LoopC

                ' Facciones válidas
              For LoopC = 1 To 4
                    AllowedFaction(LoopC) = Val(chkClass(LoopC).value)
              Next LoopC
              
50            Debug.Print "Modality: " & Val(cmbModality.ListIndex + 1)
              
60            WriteNewEvent Val(cmbModality.ListIndex + 1), _
                  Val(cmbQuotas.List(cmbQuotas.ListIndex)), _
                  Val(cmbLvlMin.List(cmbLvlMin.ListIndex)), _
                  Val(cmbLvlMax.List(cmbLvlMax.ListIndex)), _
                  Val(cmbOro.List(cmbOro.ListIndex)), Val(cmbDsp.List(cmbDsp.ListIndex)), _
                  Val(cmbInit.List(cmbInit.ListIndex)) * 60, _
                  Val(cmbCancel.List(cmbCancel.ListIndex)) * 60, _
                  Val(cmbTeam.List(cmbTeam.ListIndex)), chkAcum.value, Val(txtRojas.Text), _
                  cmbDspPremio.ListIndex, cmbOroPremio.ListIndex, Val(ReadField(1, txtObjPremio.Text, Asc("-"))), _
                  Val(ReadField(2, txtObjPremio.Text, Asc("-"))), chkItem.value, chkGanadorSigue.value, AllowedFaction(), AllowedClasses()
70        End If
          
End Sub

Private Function CheckForm(ByRef AllowedClasses() As Byte) As Boolean
10        CheckForm = False
          
20        If cmbModality.Text = "Vacio" Then
30            MsgBox "Seleccione la modalidad del evento"
40            Exit Function
50        End If
          
60        If cmbQuotas.Text = "Vacio" Then
70            MsgBox "Seleccione la cantidad de cupos que tendra"
80            Exit Function
90        End If
          
100       If cmbLvlMin.Text = "Vacio" Then
110           MsgBox "Seleccione el nivel minimo que requerirá el evento"
120           Exit Function
130       End If
          
140       If cmbLvlMax.Text = "Vacio" Then
150           MsgBox "Seleccione el nivel máximo que requerirá el evento"
160           Exit Function
170       End If
          
180       If cmbOro.Text = "Vacio" Then
190           MsgBox "Seleccione la inscripción por ORO"
200           Exit Function
210       End If
          
220       If cmbDsp.Text = "Vacio" Then
230           MsgBox "Seleccione la inscripción por DSP"
240           Exit Function
250       End If
          
260       If cmbInit.Text = "Vacio" Then
270           MsgBox "Seleccione el tiempo que tardará en iniciar las incripciones"
280           Exit Function
290       End If
          
300       If cmbCancel.Text = "Vacio" Then
310           MsgBox _
                  "Seleccione el tiempo que tendrá para cancelarse el evento si no se completan cupos"
320           Exit Function
330       End If
          
          Dim LoopC As Integer, Puede As Boolean
          
340       For LoopC = 1 To NUMCLASES
350           If AllowedClasses(LoopC) = 1 Then
360               Puede = True
370               Exit For
380           End If
390       Next LoopC
              
400       CheckForm = True
End Function


Private Sub cmbInvasiones_Change()

End Sub

Private Sub cmbModalityCurso_Click()
10        If cmbModalityCurso.ListIndex = -1 Then Exit Sub
          
20        lblClose.Visible = IIf(cmbModalityCurso.List(cmbModalityCurso.ListIndex) <> _
              "Vacio", True, False)
          
30        If cmbModalityCurso.List(cmbModalityCurso.ListIndex) = "Vacio" Then
40            MsgBox "No se puede ver la información del evento que seleccionaste."
50            Exit Sub
60        End If
          
70        Protocol.WriteRequiredDataEvent cmbModalityCurso.ListIndex + 1
End Sub

Private Sub Command1_Click()
10        Protocol.WriteRequiredEvents
End Sub

Private Sub Command2_Click()
    
    
    If txtName.Text = vbNullString Then Exit Sub
    If txtDesc.Text = vbNullString Then Exit Sub
    If Val(txtMap.Text) = 0 Then Exit Sub
    If Val(txtInvasion.Text) = 0 Then Exit Sub
    
    WriteCreateInvasion txtName.Text, txtDesc.Text, Val(txtInvasion.Text), Val(txtMap.Text)
End Sub

Private Sub Command4_Click()
    PicInvasiones.Visible = True
    
End Sub

Private Sub Command5_Click()
10  WriteTerminateInvasion

End Sub


Private Sub Form_Load()
          
          Dim LoopC As Integer
          
10        cmbModality.AddItem "Castle Mode"
20        cmbModality.AddItem "DagaRusa"
30        cmbModality.AddItem "DeathMatch"
40        cmbModality.AddItem "Aracnus"
50        cmbModality.AddItem "HombreLobo"
60        cmbModality.AddItem "Minotauro"
70        cmbModality.AddItem "Busqueda"
80        cmbModality.AddItem "Unstoppable"
90        cmbModality.AddItem "Invasion"
100       cmbModality.AddItem "Enfrentamientos"
          
110       For LoopC = 2 To 64
120           cmbQuotas.AddItem LoopC
130       Next LoopC
          
140       For LoopC = 1 To 47
150           cmbLvlMin.AddItem LoopC
160           cmbLvlMax.AddItem LoopC
170       Next LoopC
          
180       cmbLvlMin.ListIndex = 0
190       cmbLvlMax.ListIndex = 46
          
          
200       For LoopC = 1 To 10
210           cmbTeam.AddItem LoopC
220       Next LoopC
          
230       cmbTeam.ListIndex = 0
          
240       cmbOro.AddItem "0"
250       cmbOro.AddItem "25000"
260       cmbOro.AddItem "50000"
270       cmbOro.AddItem "100000"
280       cmbOro.AddItem "200000"
290       cmbOro.AddItem "300000"
300       cmbOro.AddItem "400000"
310       cmbOro.AddItem "500000"
320       cmbOro.AddItem "600000"
330       cmbOro.AddItem "700000"
340       cmbOro.AddItem "800000"
350       cmbOro.AddItem "900000"
360       cmbOro.AddItem "1000000"
370       cmbOro.ListIndex = 0
          
380       cmbDsp.AddItem "0"
390       cmbDsp.AddItem "1"
400       cmbDsp.AddItem "2"
410       cmbDsp.AddItem "5"
420       cmbDsp.AddItem "10"
430       cmbDsp.AddItem "15"
440       cmbDsp.AddItem "20"
450       cmbDsp.AddItem "25"
460       cmbDsp.AddItem "30"
470       cmbDsp.AddItem "35"
480       cmbDsp.AddItem "40"
490       cmbDsp.AddItem "45"
500       cmbDsp.AddItem "50"
          
510       cmbDsp.ListIndex = 0
          
520       For LoopC = 1 To 10
530           cmbCancel.AddItem LoopC
540           cmbInit.AddItem LoopC
550       Next LoopC
          
560       cmbCancel.ListIndex = 7
570       cmbInit.ListIndex = 0
          
          
          
End Sub


Private Sub lblClose_Click()
10        Protocol.WriteCloseEvent cmbModalityCurso.ListIndex + 1
End Sub

