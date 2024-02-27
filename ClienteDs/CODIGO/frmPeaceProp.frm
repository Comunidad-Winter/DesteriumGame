VERSION 5.00
Begin VB.Form frmPeaceProp 
   BorderStyle     =   0  'None
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   5055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   Picture         =   "frmPeaceProp.frx":0000
   ScaleHeight     =   3285
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1980
      ItemData        =   "frmPeaceProp.frx":FE48
      Left            =   240
      List            =   "frmPeaceProp.frx":FE4A
      TabIndex        =   0
      Top             =   510
      Width           =   4575
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   2520
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   3720
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   1200
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Top             =   2640
      Width           =   975
   End
End
Attribute VB_Name = "frmPeaceProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
Private clsFormulario As clsFormMovementManager
Private tipoprop As TIPO_PROPUESTA

Public Enum TIPO_PROPUESTA
    ALIANZA = 1
    PAZ = 2
End Enum

Public Property Let ProposalType(ByVal nValue As TIPO_PROPUESTA)
10        tipoprop = nValue
End Property

Private Sub Command1_Click()
10        Unload Me
End Sub


Private Sub Command3_Click()
          'Me.Visible = False
10        If tipoprop = PAZ Then
20            Call WriteGuildAcceptPeace(lista.List(lista.ListIndex))
30        Else
40            Call WriteGuildAcceptAlliance(lista.List(lista.ListIndex))
50        End If
60        Me.Hide
70        Unload Me
End Sub

Private Sub Command4_Click()
10        If tipoprop = PAZ Then
20            Call WriteGuildRejectPeace(lista.List(lista.ListIndex))
30        Else
40            Call WriteGuildRejectAlliance(lista.List(lista.ListIndex))
50        End If
60        Me.Hide
70        Unload Me
End Sub

Private Sub Form_Load()
          ' Handles Form movement (drag and drop).
10        Set clsFormulario = New clsFormMovementManager
20        clsFormulario.Initialize Me
          
30        Call LoadBackGround

End Sub

Private Sub Image1_Click()
10    Call Audio.PlayWave(SND_CLICK)
20    Unload Me
End Sub

Private Sub Image2_Click()
10    Call Audio.PlayWave(SND_CLICK)
      'Me.Visible = False
20    If tipoprop = PAZ Then
30        Call WriteGuildPeaceDetails(lista.List(lista.ListIndex))
40    Else
50        Call WriteGuildAllianceDetails(lista.List(lista.ListIndex))
60    End If
End Sub

Private Sub Image3_Click()
10    Call Audio.PlayWave(SND_CLICK)
          'Me.Visible = False
20        If tipoprop = PAZ Then
30            Call WriteGuildAcceptPeace(lista.List(lista.ListIndex))
40        Else
50            Call WriteGuildAcceptAlliance(lista.List(lista.ListIndex))
60        End If
70        Me.Hide
80        Unload Me
End Sub

Private Sub Image4_Click()
10    Call Audio.PlayWave(SND_CLICK)
20        If tipoprop = PAZ Then
30            Call WriteGuildRejectPeace(lista.List(lista.ListIndex))
40        Else
50            Call WriteGuildRejectAlliance(lista.List(lista.ListIndex))
60        End If
70        Me.Hide
80        Unload Me
End Sub
Private Sub LoadBackGround()
10        If tipoprop = TIPO_PROPUESTA.ALIANZA Then
20            Me.Picture = LoadPicture(DirGraficos & "VentanaOfertaAlianza.jpg")
30        Else
40            Me.Picture = LoadPicture(DirGraficos & "VentanaOfertaPaz.jpg")
50        End If
End Sub
