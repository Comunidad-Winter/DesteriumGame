VERSION 5.00
Begin VB.Form FrmInterv 
   Caption         =   "Intervalos"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar Intervalos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   36
      Top             =   4320
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Frame Frame11 
      Caption         =   "NPCs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2880
      TabIndex        =   49
      Top             =   2160
      Width           =   1695
      Begin VB.Frame Frame4 
         Caption         =   "A.I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   150
         TabIndex        =   50
         Top             =   240
         Width           =   1365
         Begin VB.TextBox txtAI 
            Height          =   285
            Left            =   150
            TabIndex        =   52
            Text            =   "0"
            Top             =   1080
            Width           =   1050
         End
         Begin VB.TextBox txtNPCPuedeAtacar 
            Height          =   285
            Left            =   135
            TabIndex        =   51
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "AI"
            Height          =   195
            Left            =   165
            TabIndex        =   54
            Top             =   840
            Width           =   150
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Puede atacar"
            Height          =   195
            Left            =   150
            TabIndex        =   53
            Top             =   255
            Width           =   960
         End
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Clima && Ambiente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4680
      TabIndex        =   39
      Top             =   2160
      Width           =   2865
      Begin VB.Frame Frame7 
         Caption         =   "Frio y Fx Ambientales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   2625
         Begin VB.TextBox txtCmdExec 
            Height          =   285
            Left            =   1320
            TabIndex        =   44
            Text            =   "0"
            Top             =   1110
            Width           =   915
         End
         Begin VB.TextBox txtIntervaloPerdidaStaminaLluvia 
            Height          =   300
            Left            =   1320
            TabIndex        =   43
            Text            =   "0"
            Top             =   480
            Width           =   930
         End
         Begin VB.TextBox txtIntervaloWAVFX 
            Height          =   300
            Left            =   150
            TabIndex        =   42
            Text            =   "0"
            Top             =   480
            Width           =   930
         End
         Begin VB.TextBox txtIntervaloFrio 
            Height          =   285
            Left            =   180
            TabIndex        =   41
            Text            =   "0"
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "TimerExec"
            Height          =   195
            Left            =   1320
            TabIndex        =   48
            Top             =   840
            Width           =   750
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Stamina Lluvia"
            Height          =   195
            Left            =   1350
            TabIndex        =   47
            Top             =   270
            Width           =   1035
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "FxS"
            Height          =   195
            Left            =   180
            TabIndex        =   46
            Top             =   270
            Width           =   270
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Frio"
            Height          =   195
            Left            =   195
            TabIndex        =   45
            Top             =   810
            Width           =   255
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Usuarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   7455
      Begin VB.Frame Frame9 
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   90
         TabIndex        =   24
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtIntervaloParaConexion 
            Height          =   300
            Left            =   45
            TabIndex        =   26
            Text            =   "0"
            Top             =   495
            Width           =   930
         End
         Begin VB.TextBox txtTrabajo 
            Height          =   300
            Left            =   60
            TabIndex        =   25
            Text            =   "0"
            Top             =   1020
            Width           =   930
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "IntervaloCon"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   270
            Width           =   900
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Trabajo"
            Height          =   195
            Left            =   165
            TabIndex        =   27
            Top             =   780
            Width           =   540
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Combate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   1545
         TabIndex        =   19
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtPuedeAtacar 
            Height          =   300
            Left            =   135
            TabIndex        =   22
            Text            =   "0"
            Top             =   1200
            Width           =   930
         End
         Begin VB.TextBox txtIntervaloLanzaHechizo 
            Height          =   300
            Left            =   150
            TabIndex        =   20
            Text            =   "0"
            Top             =   525
            Width           =   930
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Puede Atacar"
            Height          =   195
            Left            =   135
            TabIndex        =   23
            Top             =   930
            Width           =   975
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Lanza Spell"
            Height          =   195
            Left            =   150
            TabIndex        =   21
            Top             =   285
            Width           =   825
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Hambre y sed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   5925
         TabIndex        =   14
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtIntervaloHambre 
            Height          =   285
            Left            =   150
            TabIndex        =   16
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.TextBox txtIntervaloSed 
            Height          =   285
            Left            =   150
            TabIndex        =   15
            Text            =   "0"
            Top             =   1185
            Width           =   1050
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Hambre"
            Height          =   195
            Left            =   180
            TabIndex        =   18
            Top             =   255
            Width           =   555
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Sed"
            Height          =   195
            Left            =   165
            TabIndex        =   17
            Top             =   930
            Width           =   285
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sanar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   4470
         TabIndex        =   9
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtSanaIntervaloDescansar 
            Height          =   285
            Left            =   150
            TabIndex        =   11
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.TextBox txtSanaIntervaloSinDescansar 
            Height          =   285
            Left            =   150
            TabIndex        =   10
            Text            =   "0"
            Top             =   1185
            Width           =   1050
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descansando"
            Height          =   195
            Left            =   180
            TabIndex        =   13
            Top             =   255
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sin descansar"
            Height          =   195
            Left            =   165
            TabIndex        =   12
            Top             =   930
            Width           =   1005
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Stamina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   3015
         TabIndex        =   4
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtStaminaIntervaloSinDescansar 
            Height          =   285
            Left            =   150
            TabIndex        =   6
            Text            =   "0"
            Top             =   1185
            Width           =   1050
         End
         Begin VB.TextBox txtStaminaIntervaloDescansar 
            Height          =   285
            Left            =   165
            TabIndex        =   5
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Sin descansar"
            Height          =   195
            Left            =   165
            TabIndex        =   8
            Top             =   930
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Descansando"
            Height          =   195
            Left            =   180
            TabIndex        =   7
            Top             =   255
            Width           =   990
         End
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Magia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   2655
      Begin VB.Frame Frame10 
         Caption         =   "Duracion Spells"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   135
         TabIndex        =   29
         Top             =   270
         Width           =   2400
         Begin VB.TextBox txtInvocacion 
            Height          =   300
            Left            =   1170
            TabIndex        =   37
            Text            =   "0"
            Top             =   1170
            Width           =   900
         End
         Begin VB.TextBox txtIntervaloInvisible 
            Height          =   300
            Left            =   1170
            TabIndex        =   34
            Text            =   "0"
            Top             =   495
            Width           =   900
         End
         Begin VB.TextBox txtIntervaloParalizado 
            Height          =   300
            Left            =   195
            TabIndex        =   31
            Text            =   "0"
            Top             =   1170
            Width           =   795
         End
         Begin VB.TextBox txtIntervaloVeneno 
            Height          =   300
            Left            =   195
            TabIndex        =   30
            Text            =   "0"
            Top             =   510
            Width           =   795
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Invocacion"
            Height          =   195
            Left            =   1170
            TabIndex        =   38
            Top             =   960
            Width           =   795
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Invisible"
            Height          =   195
            Left            =   1170
            TabIndex        =   35
            Top             =   285
            Width           =   570
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Paralizado"
            Height          =   195
            Left            =   225
            TabIndex        =   33
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Veneno"
            Height          =   180
            Left            =   225
            TabIndex        =   32
            Top             =   300
            Width           =   555
         End
      End
   End
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   2055
   End
End
Attribute VB_Name = "FrmInterv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Sub AplicarIntervalos()

      '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿ Intervalos del main loop ¿?¿?¿?¿?¿?¿?¿?¿?¿
10    SanaIntervaloSinDescansar = val(txtSanaIntervaloSinDescansar.Text)
20    StaminaIntervaloSinDescansar = val(txtStaminaIntervaloSinDescansar.Text)
30    SanaIntervaloDescansar = val(txtSanaIntervaloDescansar.Text)
40    StaminaIntervaloDescansar = val(txtStaminaIntervaloDescansar.Text)
50    IntervaloSed = val(txtIntervaloSed.Text)
60    IntervaloHambre = val(txtIntervaloHambre.Text)
70    IntervaloVeneno = val(txtIntervaloVeneno.Text)
80    IntervaloParalizado = val(txtIntervaloParalizado.Text)
90    IntervaloInvisible = val(txtIntervaloInvisible.Text)
100   IntervaloFrio = val(txtIntervaloFrio.Text)
110   IntervaloWavFx = val(txtIntervaloWAVFX.Text)
120   IntervaloInvocacion = val(txtInvocacion.Text)
130   IntervaloParaConexion = val(txtIntervaloParaConexion.Text)

      '///////////////// TIMERS \\\\\\\\\\\\\\\\\\\

140   IntervaloUserPuedeCastear = val(txtIntervaloLanzaHechizo.Text)
150   frmMain.npcataca.Interval = val(txtNPCPuedeAtacar.Text)
160   frmMain.TIMER_AI.Interval = val(txtAI.Text)
170   IntervaloUserPuedeTrabajar = val(txtTrabajo.Text)
180   IntervaloUserPuedeAtacar = val(txtPuedeAtacar.Text)



End Sub

Private Sub Command1_Click()
10    On Error Resume Next
20    Call AplicarIntervalos

End Sub

Private Sub Command2_Click()

10    On Error GoTo Err

      'Intervalos
20    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar", Str(SanaIntervaloSinDescansar))
30    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar", Str(StaminaIntervaloSinDescansar))
40    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar", Str(SanaIntervaloDescansar))
50    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar", Str(StaminaIntervaloDescansar))
60    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed", Str(IntervaloSed))
70    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre", Str(IntervaloHambre))
80    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno", Str(IntervaloVeneno))
90    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado", Str(IntervaloParalizado))
100   Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible", Str(IntervaloInvisible))
110   Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio", Str(IntervaloFrio))
120   Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX", Str(IntervaloWavFx))
130   Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion", Str(IntervaloInvocacion))
140   Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion", Str(IntervaloParaConexion))

      '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&

150   Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo", Str(IntervaloUserPuedeCastear))
160   Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI", frmMain.TIMER_AI.Interval)
170   Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar", frmMain.npcataca.Interval)
180   Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo", Str(IntervaloUserPuedeTrabajar))
190   Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar", Str(IntervaloUserPuedeAtacar))


200   MsgBox "Los intervalos se han guardado sin problemas."

210   Exit Sub
Err:
220       MsgBox "Error al intentar grabar los intervalos"
End Sub

Private Sub ok_Click()
10    Me.Visible = False
End Sub

