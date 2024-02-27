VERSION 5.00
Begin VB.Form frmEligeAlineacion 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desterium AO"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Legionario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Neutral"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Armada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label imgReal 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEligeAlineacion.frx":0000
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Label imgNeutral 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEligeAlineacion.frx":00AB
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   5655
   End
   Begin VB.Label imgCaos 
      BackStyle       =   0  'Transparent
      Caption         =   "La legión oscura clan caotico solo para personas con maldad y ganas de sangre Real y todo lo que encuentre en su camino."
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   3360
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Escoje la alineacion que llevara tu clan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmEligeAlineacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmEligeAlineacion.frm
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonCriminal As clsGraphicalButton
Private cBotonCaos As clsGraphicalButton
Private cBotonLegal As clsGraphicalButton
Private cBotonNeutral As clsGraphicalButton
Private cBotonReal As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Enum eAlineacion
    ieREAL = 0
    ieCAOS = 1
    ieNeutral = 2
    ieLegal = 4
    ieCriminal = 5
End Enum

Private Sub Command1_Click()
10    Unload Me
End Sub

Private Sub Form_Load()
          ' Handles Form movement (drag and drop).
10        Set clsFormulario = New clsFormMovementManager
20        clsFormulario.Initialize Me
          
          'Me.Picture = LoadPicture(App.path & "\Recursos\VentanaFundarClan.jpg")
          
30        Call LoadButtons
End Sub

Private Sub LoadButtons()
          Dim GrhPath As String
          
10        GrhPath = DirGraficos

20        Set cBotonCriminal = New clsGraphicalButton
30        Set cBotonCaos = New clsGraphicalButton
40        Set cBotonLegal = New clsGraphicalButton
50        Set cBotonNeutral = New clsGraphicalButton
60        Set cBotonReal = New clsGraphicalButton
70        Set cBotonSalir = New clsGraphicalButton
          
80        Set LastPressed = New clsGraphicalButton

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
10        LastPressed.ToggleToNormal
End Sub

Private Sub imgCaos_Click()
10        Call WriteGuildFundation(eAlineacion.ieCAOS)
20        Unload Me
End Sub

Private Sub imgCriminal_Click()
10        Call WriteGuildFundation(eAlineacion.ieCriminal)
20        Unload Me
End Sub

Private Sub imgLegal_Click()
10        Call WriteGuildFundation(eAlineacion.ieLegal)
20        Unload Me
End Sub

Private Sub imgNeutral_Click()
10        Call WriteGuildFundation(eAlineacion.ieNeutral)
20        Unload Me
End Sub

Private Sub imgReal_Click()
10        Call WriteGuildFundation(eAlineacion.ieREAL)
20        Unload Me
End Sub

