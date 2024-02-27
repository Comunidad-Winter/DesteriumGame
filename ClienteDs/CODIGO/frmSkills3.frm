VERSION 5.00
Begin VB.Form frmSkills3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSkills3.frx":0000
   ScaleHeight     =   418
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image ImgMenos22 
      Height          =   195
      Left            =   3240
      Top             =   480
      Width           =   225
   End
   Begin VB.Image ImgMas22 
      Height          =   195
      Left            =   5880
      Top             =   480
      Width           =   225
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   22
      Left            =   4950
      TabIndex        =   80
      Top             =   480
      Width           =   405
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resistencia Mágica:"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   3
      Left            =   3585
      TabIndex        =   79
      Top             =   480
      Width           =   1410
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Equitación: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   2
      Left            =   3585
      TabIndex        =   78
      Top             =   720
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   285
      TabIndex        =   77
      Top             =   2700
      Width           =   270
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   900
      TabIndex        =   76
      Top             =   510
      Width           =   180
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   960
      TabIndex        =   75
      Top             =   720
      Width           =   180
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   1230
      TabIndex        =   74
      Top             =   945
      Width           =   180
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   960
      TabIndex        =   73
      Top             =   1155
      Width           =   180
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   1290
      TabIndex        =   72
      Top             =   1365
      Width           =   180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   945
      TabIndex        =   71
      Top             =   2235
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   960
      TabIndex        =   70
      Top             =   2475
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   885
      TabIndex        =   69
      Top             =   2940
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   795
      TabIndex        =   68
      Top             =   3180
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   780
      TabIndex        =   67
      Top             =   3420
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   285
      TabIndex        =   66
      Top             =   3660
      Width           =   270
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   1725
      TabIndex        =   65
      Top             =   4380
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1845
      TabIndex        =   64
      Top             =   4620
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1620
      TabIndex        =   63
      Top             =   4860
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   1380
      TabIndex        =   62
      Top             =   5100
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   735
      TabIndex        =   61
      Top             =   5340
      Width           =   1785
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   2205
      TabIndex        =   60
      Top             =   5580
      Width           =   555
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fuerza:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   59
      Top             =   510
      Width           =   555
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agilidad:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   300
      TabIndex        =   58
      Top             =   735
      Width           =   615
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inteligencia:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   300
      TabIndex        =   57
      Top             =   945
      Width           =   885
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Carisma:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   300
      TabIndex        =   56
      Top             =   1155
      Width           =   630
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Constitución:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   300
      TabIndex        =   55
      Top             =   1365
      Width           =   945
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Asesino:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   285
      TabIndex        =   54
      Top             =   2235
      Width           =   615
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bandido:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   285
      TabIndex        =   53
      Top             =   2475
      Width           =   630
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ladrón:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   285
      TabIndex        =   52
      Top             =   2940
      Width           =   555
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Noble:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   285
      TabIndex        =   51
      Top             =   3180
      Width           =   465
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Plebe:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   285
      TabIndex        =   50
      Top             =   3420
      Width           =   450
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Criminales Matados:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   240
      TabIndex        =   49
      Top             =   4380
      Width           =   1440
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudadanos Matados:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   240
      TabIndex        =   48
      Top             =   4620
      Width           =   1560
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuarios Matados:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   240
      TabIndex        =   47
      Top             =   4860
      Width           =   1335
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NPCs matados:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   240
      TabIndex        =   46
      Top             =   5100
      Width           =   1095
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clase:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   240
      TabIndex        =   45
      Top             =   5340
      Width           =   450
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tiempo restante en carcel:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   240
      TabIndex        =   44
      Top             =   5580
      Width           =   1920
   End
   Begin VB.Label lbldatos 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel: 17 Experiencia"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   43
      Top             =   5925
      Width           =   3975
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Skills Libres:"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   1
      Left            =   3840
      TabIndex        =   42
      Top             =   5925
      Width           =   855
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Magia:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   3585
      TabIndex        =   41
      Top             =   945
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Robar: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   0
      Left            =   3585
      TabIndex        =   40
      Top             =   1170
      Width           =   540
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tacticas de combate: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   22
      Left            =   3585
      TabIndex        =   39
      Top             =   1410
      Width           =   1575
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supervivencia: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   23
      Left            =   3585
      TabIndex        =   38
      Top             =   2580
      Width           =   1110
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ocultarse: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   24
      Left            =   3585
      TabIndex        =   37
      Top             =   2355
      Width           =   795
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apuñalar: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   25
      Left            =   3585
      TabIndex        =   36
      Top             =   2115
      Width           =   750
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Meditar: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   26
      Left            =   3585
      TabIndex        =   35
      Top             =   1875
      Width           =   645
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mineria: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   27
      Left            =   3585
      TabIndex        =   34
      Top             =   3765
      Width           =   615
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pesca: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   28
      Left            =   3585
      TabIndex        =   33
      Top             =   3525
      Width           =   525
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Defensa con escudos: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   29
      Left            =   3585
      TabIndex        =   32
      Top             =   3285
      Width           =   1635
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comercio: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   30
      Left            =   3585
      TabIndex        =   31
      Top             =   3060
      Width           =   765
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Talar árboles: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   31
      Left            =   3585
      TabIndex        =   30
      Top             =   2820
      Width           =   1035
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Navegacion: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   32
      Left            =   3600
      TabIndex        =   29
      Top             =   5400
      Width           =   945
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Combate sin armas: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   33
      Left            =   3600
      TabIndex        =   28
      Top             =   5160
      Width           =   1470
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Armas de proyectiles: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   34
      Left            =   3585
      TabIndex        =   27
      Top             =   4935
      Width           =   1605
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Domar animales: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   35
      Left            =   3585
      TabIndex        =   26
      Top             =   4695
      Width           =   1230
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Liderazgo: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   36
      Left            =   3585
      TabIndex        =   25
      Top             =   4470
      Width           =   795
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Herreria: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   37
      Left            =   3585
      TabIndex        =   24
      Top             =   4230
      Width           =   690
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Carpinteria: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   38
      Left            =   3585
      TabIndex        =   23
      Top             =   3990
      Width           =   900
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Combate con armas: "
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   39
      Left            =   3585
      TabIndex        =   22
      Top             =   1650
      Width           =   1530
   End
   Begin VB.Image imgMas21 
      Height          =   195
      Left            =   5880
      Top             =   720
      Width           =   225
   End
   Begin VB.Image imgMenos21 
      Height          =   195
      Left            =   3240
      Top             =   720
      Width           =   225
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   21
      Left            =   4335
      TabIndex        =   21
      Top             =   735
      Width           =   405
   End
   Begin VB.Image imgCancelar 
      Height          =   360
      Left            =   7080
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4110
      TabIndex        =   20
      Top             =   960
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4110
      TabIndex        =   19
      Top             =   1185
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   5145
      TabIndex        =   18
      Top             =   1425
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   5100
      TabIndex        =   13
      Top             =   1665
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   4200
      TabIndex        =   17
      Top             =   1890
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   4320
      TabIndex        =   16
      Top             =   2130
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   4365
      TabIndex        =   15
      Top             =   2370
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   4665
      TabIndex        =   14
      Top             =   2595
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   4620
      TabIndex        =   12
      Top             =   2835
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   4335
      TabIndex        =   11
      Top             =   3075
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   5190
      TabIndex        =   10
      Top             =   3300
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   4065
      TabIndex        =   9
      Top             =   3540
      Width           =   285
   End
   Begin VB.Image imgMas1 
      Height          =   195
      Left            =   5880
      Top             =   960
      Width           =   180
   End
   Begin VB.Image imgMas2 
      Height          =   195
      Left            =   5880
      Top             =   1200
      Width           =   180
   End
   Begin VB.Image imgMenos2 
      Height          =   195
      Left            =   3240
      Top             =   1200
      Width           =   195
   End
   Begin VB.Image imgMas3 
      Height          =   180
      Left            =   5880
      Top             =   1440
      Width           =   180
   End
   Begin VB.Image imgMenos3 
      Height          =   180
      Left            =   3240
      Top             =   1440
      Width           =   180
   End
   Begin VB.Image imgMas4 
      Height          =   165
      Left            =   5880
      Top             =   1680
      Width           =   180
   End
   Begin VB.Image imgMenos4 
      Height          =   180
      Left            =   3240
      Top             =   1680
      Width           =   195
   End
   Begin VB.Image imgMas5 
      Height          =   180
      Left            =   5880
      Top             =   1920
      Width           =   180
   End
   Begin VB.Image imgMenos5 
      Height          =   150
      Left            =   3240
      Top             =   1920
      Width           =   180
   End
   Begin VB.Image imgMas6 
      Height          =   180
      Left            =   5880
      Top             =   2160
      Width           =   180
   End
   Begin VB.Image imgMenos6 
      Height          =   180
      Left            =   3240
      Top             =   2145
      Width           =   180
   End
   Begin VB.Image imgMas7 
      Height          =   180
      Left            =   5880
      Top             =   2385
      Width           =   180
   End
   Begin VB.Image imgMenos7 
      Height          =   180
      Left            =   3240
      Top             =   2370
      Width           =   180
   End
   Begin VB.Image imgMas8 
      Height          =   195
      Left            =   5880
      Top             =   2580
      Width           =   180
   End
   Begin VB.Image imgMenos8 
      Height          =   180
      Left            =   3240
      Top             =   2595
      Width           =   195
   End
   Begin VB.Image imgMas9 
      Height          =   195
      Left            =   5880
      Top             =   2820
      Width           =   180
   End
   Begin VB.Image imgMenos9 
      Height          =   180
      Left            =   3240
      Top             =   2820
      Width           =   195
   End
   Begin VB.Image imgMas10 
      Height          =   195
      Left            =   5880
      Top             =   3060
      Width           =   180
   End
   Begin VB.Image imgMenos10 
      Height          =   180
      Left            =   3240
      Top             =   3060
      Width           =   195
   End
   Begin VB.Image imgMas11 
      Height          =   195
      Left            =   5880
      Top             =   3300
      Width           =   195
   End
   Begin VB.Image imgMenos11 
      Height          =   165
      Left            =   3240
      Top             =   3300
      Width           =   195
   End
   Begin VB.Image imgMas12 
      Height          =   195
      Left            =   5880
      Top             =   3525
      Width           =   195
   End
   Begin VB.Image imgMenos12 
      Height          =   195
      Left            =   3240
      Top             =   3525
      Width           =   195
   End
   Begin VB.Image imgMas13 
      Height          =   195
      Left            =   5880
      Top             =   3780
      Width           =   195
   End
   Begin VB.Image imgMenos13 
      Height          =   195
      Left            =   3240
      Top             =   3765
      Width           =   195
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   4185
      TabIndex        =   8
      Top             =   3780
      Width           =   285
   End
   Begin VB.Image imgMas14 
      Height          =   195
      Left            =   5880
      Top             =   4020
      Width           =   195
   End
   Begin VB.Image imgMenos14 
      Height          =   165
      Left            =   3240
      Top             =   4005
      Width           =   195
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   4470
      TabIndex        =   7
      Top             =   4020
      Width           =   285
   End
   Begin VB.Image imgMas15 
      Height          =   165
      Left            =   5880
      Top             =   4260
      Width           =   195
   End
   Begin VB.Image imgMenos15 
      Height          =   195
      Left            =   3240
      Top             =   4230
      Width           =   195
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   4215
      TabIndex        =   6
      Top             =   4245
      Width           =   285
   End
   Begin VB.Image imgMas16 
      Height          =   195
      Left            =   5880
      Top             =   4470
      Width           =   195
   End
   Begin VB.Image imgMenos16 
      Height          =   195
      Left            =   3240
      Top             =   4470
      Width           =   195
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   4365
      TabIndex        =   5
      Top             =   4485
      Width           =   285
   End
   Begin VB.Image imgMas17 
      Height          =   195
      Left            =   5880
      Top             =   4695
      Width           =   195
   End
   Begin VB.Image imgMenos17 
      Height          =   195
      Left            =   3240
      Top             =   4695
      Width           =   195
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   4800
      TabIndex        =   4
      Top             =   4710
      Width           =   285
   End
   Begin VB.Image imgMas18 
      Height          =   195
      Left            =   5880
      Top             =   4935
      Width           =   195
   End
   Begin VB.Image imgMenos18 
      Height          =   165
      Left            =   3240
      Top             =   4920
      Width           =   195
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   5160
      TabIndex        =   3
      Top             =   4950
      Width           =   285
   End
   Begin VB.Image imgMenos1 
      Height          =   195
      Left            =   3240
      Top             =   960
      Width           =   195
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   5055
      TabIndex        =   2
      Top             =   5175
      Width           =   285
   End
   Begin VB.Image imgMas19 
      Height          =   165
      Left            =   5880
      Top             =   5160
      Width           =   195
   End
   Begin VB.Image imgMenos19 
      Height          =   165
      Left            =   3240
      Top             =   5160
      Width           =   195
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   4500
      TabIndex        =   1
      Top             =   5415
      Width           =   285
   End
   Begin VB.Image imgMas20 
      Height          =   195
      Left            =   5880
      Top             =   5400
      Width           =   195
   End
   Begin VB.Image imgMenos20 
      Height          =   195
      Left            =   3240
      Top             =   5400
      Width           =   195
   End
   Begin VB.Image imgAceptar 
      Height          =   345
      Left            =   6000
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label puntos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4740
      TabIndex        =   0
      Top             =   5940
      Width           =   45
   End
End
Attribute VB_Name = "frmSkills3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonMas(1 To NUMSKILLS) As clsGraphicalButton
Private cBotonMenos(1 To NUMSKILLS) As clsGraphicalButton
Private cSkillNames(1 To NUMSKILLS) As clsGraphicalButton
Private cBtonAceptar As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton

Public LastPressed As clsGraphicalButton


Private cBotonCerrar As clsGraphicalButton



Private bPuedeMagia As Boolean
Private bPuedeMeditar As Boolean
Private bPuedeEscudo As Boolean
Private bPuedeCombateDistancia As Boolean

Private vsHelp(1 To NUMSKILLS) As String
Private Const ANCHO_BARRA As Byte = 73 'pixeles
Private Const BAR_LEFT_POS As Integer = 361 'pixeles

Public Sub Iniciar_Labels()
      'Iniciamos los labels con los valores de los atributos y los skills
      Dim i As Integer
      Dim Ancho As Integer

10    For i = 1 To NUMATRIBUTOS
20        Atri(i).Caption = UserAtributos(i)
30    Next




40    Label4(1).Caption = UserReputacion.AsesinoRep
50    Label4(2).Caption = UserReputacion.BandidoRep
60    Label4(3).Caption = "Burgues: " & UserReputacion.BurguesRep
70    Label4(4).Caption = UserReputacion.LadronesRep
80    Label4(5).Caption = UserReputacion.NobleRep
90    Label4(6).Caption = UserReputacion.PlebeRep

100   If UserReputacion.Promedio < 0 Then
110       Label4(7).ForeColor = &H8080FF
120       Label4(7).Caption = "Status: CRIMINAL"
130   Else
140       Label4(7).ForeColor = &HC0C000
150       Label4(7).Caption = "Status: Ciudadano"
160   End If

170   With UserEstadisticas
180       Label6(0).Caption = .CriminalesMatados
190       Label6(1).Caption = .CiudadanosMatados
200       Label6(2).Caption = .UsuariosMatados
210       Label6(3).Caption = .NpcsMatados
220       Label6(4).Caption = .Clase
230       Label6(5).Caption = .PenaCarcel
240   End With

End Sub

Private Sub Form_Load()
10        MirandoAsignarSkills = True
          
          ' Handles Form movement (drag and drop).
20        Set clsFormulario = New clsFormMovementManager
30        clsFormulario.Initialize Me
          
          'Flags para saber que skills se modificaron
40        ReDim flags(1 To NUMSKILLS)
          
50        Call ValidarSkills
          
          ' Handles Form movement (drag and drop).
60        Set clsFormulario = New clsFormMovementManager
70        clsFormulario.Initialize Me
          
          'Me.Picture = LoadPicture(App.path & "\Recursos\VentanaEstadisticas.jpg")
          
80        Call LoadButtons
          
End Sub
Private Sub LoadButtons()
          Dim GrhPath As String
          Dim i As Long
          
10        GrhPath = DirGraficos


20        For i = 1 To NUMSKILLS
30            Set cBotonMas(i) = New clsGraphicalButton
40            Set cBotonMenos(i) = New clsGraphicalButton
50            Set cSkillNames(i) = New clsGraphicalButton
60        Next i
          
70        Set cBtonAceptar = New clsGraphicalButton
80        Set cBotonCancelar = New clsGraphicalButton
          
90        Set LastPressed = New clsGraphicalButton
          
100       Call cBotonMas(1).Initialize(imgMas1, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me, GrhPath & "BotonMasSkills.jpg", Not bPuedeMagia)

110       Call cBotonMas(2).Initialize(imgMas2, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)

120       Call cBotonMas(3).Initialize(imgMas3, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)

130       Call cBotonMas(4).Initialize(imgMas4, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)
          
140       Call cBotonMas(5).Initialize(imgMas5, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me, GrhPath & "BotonMasSkills.jpg", Not bPuedeMeditar)

150       Call cBotonMas(6).Initialize(imgMas6, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)

160       Call cBotonMas(7).Initialize(imgMas7, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)

170       Call cBotonMas(8).Initialize(imgMas8, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)
          
180       Call cBotonMas(9).Initialize(imgMas9, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)

190       Call cBotonMas(10).Initialize(imgMas10, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)

200       Call cBotonMas(11).Initialize(imgMas11, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me, GrhPath & "BotonMasSkills.jpg", Not bPuedeEscudo)

210       Call cBotonMas(12).Initialize(imgMas12, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)
          
220       Call cBotonMas(13).Initialize(imgMas13, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)

230       Call cBotonMas(14).Initialize(imgMas14, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)

240       Call cBotonMas(15).Initialize(imgMas15, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)

250       Call cBotonMas(16).Initialize(imgMas16, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)
          
260       Call cBotonMas(17).Initialize(imgMas17, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)

270       Call cBotonMas(18).Initialize(imgMas18, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me, GrhPath & "BotonMasSkills.jpg", Not bPuedeCombateDistancia)

280       Call cBotonMas(19).Initialize(imgMas19, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)

290       Call cBotonMas(20).Initialize(imgMas20, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)
                                          
300       Call cBotonMas(21).Initialize(imgMas21, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)
                                          
310       Call cBotonMas(22).Initialize(ImgMas22, GrhPath & "BotonMasSkills.jpg", _
              GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", _
              Me)
          
          
320       Call cBotonMenos(1).Initialize(imgMenos1, GrhPath & "BotonMenosSkills.jpg", _
              GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me, GrhPath & "BotonMenosSkills.jpg", Not _
              bPuedeMagia)

330       Call cBotonMenos(2).Initialize(imgMenos2, GrhPath & "BotonMenosSkills.jpg", _
              GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)

340       Call cBotonMenos(3).Initialize(imgMenos3, GrhPath & "BotonMenosSkills.jpg", _
              GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)

350       Call cBotonMenos(4).Initialize(imgMenos4, GrhPath & "BotonMenosSkills.jpg", _
              GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)
          
360       Call cBotonMenos(5).Initialize(imgMenos5, GrhPath & "BotonMenosSkills.jpg", _
              GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me, GrhPath & "BotonMenosSkills.jpg", Not _
              bPuedeMeditar)

370       Call cBotonMenos(6).Initialize(imgMenos6, GrhPath & "BotonMenosSkills.jpg", _
              GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)

380       Call cBotonMenos(7).Initialize(imgMenos7, GrhPath & "BotonMenosSkills.jpg", _
              GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)

390       Call cBotonMenos(8).Initialize(imgMenos8, GrhPath & "BotonMenosSkills.jpg", _
              GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)
          
400       Call cBotonMenos(9).Initialize(imgMenos9, GrhPath & "BotonMenosSkills.jpg", _
              GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)

410       Call cBotonMenos(10).Initialize(imgMenos10, GrhPath & _
              "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)

420       Call cBotonMenos(11).Initialize(imgMenos11, GrhPath & _
              "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me, GrhPath & "BotonMenosSkills.jpg", Not _
              bPuedeEscudo)

430       Call cBotonMenos(12).Initialize(imgMenos12, GrhPath & _
              "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)
          
440       Call cBotonMenos(13).Initialize(imgMenos13, GrhPath & _
              "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)

450       Call cBotonMenos(14).Initialize(imgMenos14, GrhPath & _
              "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)

460       Call cBotonMenos(15).Initialize(imgMenos15, GrhPath & _
              "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)

470       Call cBotonMenos(16).Initialize(imgMenos16, GrhPath & _
              "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)
          
480       Call cBotonMenos(17).Initialize(imgMenos17, GrhPath & _
              "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)

490       Call cBotonMenos(18).Initialize(imgMenos18, GrhPath & _
              "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me, GrhPath & "BotonMenosSkills.jpg", Not _
              bPuedeCombateDistancia)

500       Call cBotonMenos(19).Initialize(imgMenos19, GrhPath & _
              "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)

510       Call cBotonMenos(20).Initialize(imgMenos20, GrhPath & _
              "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)
                                          
520       Call cBotonMenos(21).Initialize(imgMenos21, GrhPath & _
              "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)
                                          
530       Call cBotonMenos(22).Initialize(ImgMenos22, GrhPath & _
              "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & _
              "BotonMenosClickSkills.jpg", Me)


        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
10        LastPressed.ToggleToNormal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
10    Unload Me
End Sub

Private Sub imgCerrar_Click()
10        Unload Me
End Sub

Private Sub imgCerrar_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
    'If imgCerrar.Tag = 1 Then
        'imgCerrar.Picture = LoadPicture(App.path & "\Recursos\BotonCerrarApretadoEstadisticas.jpg")
        'imgCerrar.Tag = 0
   ' End If

End Sub

Private Sub SumarSkillPoint(ByVal SkillIndex As Integer)
10        If Alocados > 0 Then

20            If Val(text1(SkillIndex).Caption) < MAXSKILLPOINTS Then
30                text1(SkillIndex).Caption = Val(text1(SkillIndex).Caption) + 1
40                flags(SkillIndex) = flags(SkillIndex) + 1
50                Alocados = Alocados - 1
60            End If
                  
70        End If
          
80        puntos.Caption = Alocados
End Sub

Private Sub RestarSkillPoint(ByVal SkillIndex As Integer)
10        If Alocados < SkillPoints Then
              
20            If Val(text1(SkillIndex).Caption) > 0 And flags(SkillIndex) > 0 Then
30                text1(SkillIndex).Caption = Val(text1(SkillIndex).Caption) - 1
40                flags(SkillIndex) = flags(SkillIndex) - 1
50                Alocados = Alocados + 1
60            End If
70        End If
          
80        puntos.Caption = Alocados
End Sub



Private Sub Form_Unload(Cancel As Integer)
10        MirandoAsignarSkills = False
End Sub

Private Sub imgAceptar_Click()
          Dim skillChanges(NUMSKILLS) As Byte
          Dim i As Long

10        For i = 1 To NUMSKILLS
20            skillChanges(i) = CByte(text1(i).Caption) - UserSkills(i)
              'Actualizamos nuestros datos locales
30            UserSkills(i) = Val(text1(i).Caption)
40        Next i
          
50        Call WriteModifySkills(skillChanges())
          
60        SkillPoints = Alocados
          
70        Unload Me
End Sub

Private Sub imgApunialar_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        Call ShowHelp(eSkill.Apuñalar)
End Sub

Private Sub imgCancelar_Click()
10        Unload Me
End Sub

Private Sub imgCarpinteria_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        Call ShowHelp(eSkill.Carpinteria)
End Sub

Private Sub imgCombateArmas_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        Call ShowHelp(eSkill.Armas)
End Sub

Private Sub imgCombateDistancia_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
10        Call ShowHelp(eSkill.Proyectiles)
End Sub

Private Sub imgCombateSinArmas_MouseMove(Button As Integer, Shift As Integer, X _
    As Single, Y As Single)
10        Call ShowHelp(eSkill.Wrestling)
End Sub

Private Sub imgComercio_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        Call ShowHelp(eSkill.Comerciar)
End Sub

Private Sub imgDomar_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        Call ShowHelp(eSkill.Domar)
End Sub

Private Sub imgEquitacion_Click()
10    Call ShowHelp(eSkill.Equitacion)
End Sub

Private Sub imgEscudos_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        Call ShowHelp(eSkill.Defensa)
End Sub

Private Sub imgEvasion_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        Call ShowHelp(eSkill.Tacticas)
End Sub

Private Sub imgHerreria_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        Call ShowHelp(eSkill.Herreria)
End Sub

Private Sub imgLiderazgo_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        Call ShowHelp(eSkill.Liderazgo)
End Sub

Private Sub imgMagia_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        Call ShowHelp(eSkill.Magia)
End Sub

Private Sub imgMas1_Click()
10        Call SumarSkillPoint(1)
End Sub

Private Sub imgMas10_Click()
10        Call SumarSkillPoint(10)
End Sub

Private Sub imgMas11_Click()
10        Call SumarSkillPoint(11)
End Sub

Private Sub imgMas12_Click()
10        Call SumarSkillPoint(12)
End Sub

Private Sub imgMas13_Click()
10        Call SumarSkillPoint(13)
End Sub

Private Sub imgMas14_Click()
10        Call SumarSkillPoint(14)
End Sub

Private Sub imgMas15_Click()
10        Call SumarSkillPoint(15)
End Sub

Private Sub imgMas16_Click()
10        Call SumarSkillPoint(16)
End Sub

Private Sub imgMas17_Click()
10        Call SumarSkillPoint(17)
End Sub

Private Sub imgMas18_Click()
10        Call SumarSkillPoint(18)
End Sub

Private Sub imgMas19_Click()
10        Call SumarSkillPoint(19)
End Sub

Private Sub imgMas2_Click()
10        Call SumarSkillPoint(2)
End Sub

Private Sub imgMas20_Click()
10        Call SumarSkillPoint(20)
End Sub

Private Sub imgMas21_Click()
10    Call SumarSkillPoint(21)
End Sub

Private Sub ImgMas22_Click()
10    Call SumarSkillPoint(22)
End Sub

Private Sub imgMas3_Click()
10        Call SumarSkillPoint(3)
End Sub

Private Sub imgMas4_Click()
10        Call SumarSkillPoint(4)
End Sub

Private Sub imgMas5_Click()
10        Call SumarSkillPoint(5)
End Sub

Private Sub imgMas6_Click()
10        Call SumarSkillPoint(6)
End Sub

Private Sub imgMas7_Click()
10        Call SumarSkillPoint(7)
End Sub

Private Sub imgMas8_Click()
10        Call SumarSkillPoint(8)
End Sub

Private Sub imgMas9_Click()
10        Call SumarSkillPoint(9)
End Sub

Private Sub imgMeditar_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        Call ShowHelp(eSkill.Meditar)
End Sub

Private Sub imgMenos1_Click()
10        Call RestarSkillPoint(1)
End Sub

Private Sub imgMenos10_Click()
10        Call RestarSkillPoint(10)
End Sub

Private Sub imgMenos11_Click()
10        Call RestarSkillPoint(11)
End Sub

Private Sub imgMenos12_Click()
10        Call RestarSkillPoint(12)
End Sub

Private Sub imgMenos13_Click()
10        Call RestarSkillPoint(13)
End Sub

Private Sub imgMenos14_Click()
10        Call RestarSkillPoint(14)
End Sub

Private Sub imgMenos15_Click()
10        Call RestarSkillPoint(15)
End Sub

Private Sub imgMenos16_Click()
10        Call RestarSkillPoint(16)
End Sub

Private Sub imgMenos17_Click()
10        Call RestarSkillPoint(17)
End Sub

Private Sub imgMenos18_Click()
10        Call RestarSkillPoint(18)
End Sub

Private Sub imgMenos19_Click()
10        Call RestarSkillPoint(19)
End Sub

Private Sub imgMenos2_Click()
10        Call RestarSkillPoint(2)
End Sub

Private Sub imgMenos20_Click()
10        Call RestarSkillPoint(20)
End Sub

Private Sub imgMenos21_Click()
10    Call RestarSkillPoint(21)
End Sub

Private Sub ImgMenos22_Click()
10    Call RestarSkillPoint(22)
End Sub

Private Sub imgMenos3_Click()
10        Call RestarSkillPoint(3)
End Sub

Private Sub imgMenos4_Click()
10        Call RestarSkillPoint(4)
End Sub

Private Sub imgMenos5_Click()
10        Call RestarSkillPoint(5)
End Sub

Private Sub imgMenos6_Click()
10        Call RestarSkillPoint(6)
End Sub

Private Sub imgMenos7_Click()
10        Call RestarSkillPoint(7)
End Sub

Private Sub imgMenos8_Click()
10        Call RestarSkillPoint(8)
End Sub

Private Sub imgMenos9_Click()
10        Call RestarSkillPoint(9)
End Sub



Private Sub imgequitacion_mousemove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        Call ShowHelp(eSkill.Equitacion)
End Sub

Private Sub imgMineria_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        Call ShowHelp(eSkill.Mineria)
End Sub

Private Sub imgNavegacion_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        Call ShowHelp(eSkill.Navegacion)
End Sub

Private Sub imgOcultarse_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        Call ShowHelp(eSkill.Ocultarse)
End Sub

Private Sub imgPesca_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        Call ShowHelp(eSkill.Pesca)
End Sub

Private Sub imgRobar_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        Call ShowHelp(eSkill.Robar)
End Sub

Private Sub imgSupervivencia_MouseMove(Button As Integer, Shift As Integer, X _
    As Single, Y As Single)
10        Call ShowHelp(eSkill.Supervivencia)
End Sub

Private Sub imgTalar_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        Call ShowHelp(eSkill.Talar)
End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
10        LastPressed.ToggleToNormal
End Sub

Private Sub ShowHelp(ByVal eeSkill As eSkill)
    'lblHelp.Caption = vsHelp(eeSkill)
End Sub

Private Sub ValidarSkills()

10        bPuedeMagia = True
20        bPuedeMeditar = True
30        bPuedeEscudo = True
40        bPuedeCombateDistancia = True

50        Select Case UserClase
              Case eClass.Warrior, eClass.Hunter, eClass.Worker, eClass.Thief
60                bPuedeMagia = False
70                bPuedeMeditar = False
              
80            Case eClass.Pirat
90                bPuedeMagia = False
100               bPuedeMeditar = False
110               bPuedeEscudo = False
              
120           Case eClass.Mage, eClass.Druid
130               bPuedeEscudo = False
140               bPuedeCombateDistancia = False
                  
150       End Select
          
          ' Magia
160       imgMas1.Enabled = bPuedeMagia
170       imgMenos1.Enabled = bPuedeMagia

          ' Meditar
180       imgMas5.Enabled = bPuedeMeditar
190       imgMenos5.Enabled = bPuedeMeditar

          ' Escudos
200       imgMas11.Enabled = bPuedeEscudo
210       imgMenos11.Enabled = bPuedeEscudo

          ' Proyectiles
220       imgMas18.Enabled = bPuedeCombateDistancia
230       imgMenos18.Enabled = bPuedeCombateDistancia
End Sub

