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
    tipoprop = nValue
End Property

Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub Command3_Click()
    'Me.Visible = False
    If tipoprop = PAZ Then
        Call WriteGuildAcceptPeace(lista.List(lista.ListIndex))
    Else
        Call WriteGuildAcceptAlliance(lista.List(lista.ListIndex))
    End If
    Me.Hide
    Unload Me
End Sub

Private Sub Command4_Click()
    If tipoprop = PAZ Then
        Call WriteGuildRejectPeace(lista.List(lista.ListIndex))
    Else
        Call WriteGuildRejectAlliance(lista.List(lista.ListIndex))
    End If
    Me.Hide
    Unload Me
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Call LoadBackGround

End Sub

Private Sub Image1_Click()
Call Audio.PlayWave(SND_CLICK)
Unload Me
End Sub

Private Sub Image2_Click()
Call Audio.PlayWave(SND_CLICK)
'Me.Visible = False
If tipoprop = PAZ Then
    Call WriteGuildPeaceDetails(lista.List(lista.ListIndex))
Else
    Call WriteGuildAllianceDetails(lista.List(lista.ListIndex))
End If
End Sub

Private Sub Image3_Click()
Call Audio.PlayWave(SND_CLICK)
    'Me.Visible = False
    If tipoprop = PAZ Then
        Call WriteGuildAcceptPeace(lista.List(lista.ListIndex))
    Else
        Call WriteGuildAcceptAlliance(lista.List(lista.ListIndex))
    End If
    Me.Hide
    Unload Me
End Sub

Private Sub Image4_Click()
Call Audio.PlayWave(SND_CLICK)
    If tipoprop = PAZ Then
        Call WriteGuildRejectPeace(lista.List(lista.ListIndex))
    Else
        Call WriteGuildRejectAlliance(lista.List(lista.ListIndex))
    End If
    Me.Hide
    Unload Me
End Sub
Private Sub LoadBackGround()
    If tipoprop = TIPO_PROPUESTA.ALIANZA Then
        Me.Picture = LoadPicture(DirGraficos & "VentanaOfertaAlianza.jpg")
    Else
        Me.Picture = LoadPicture(DirGraficos & "VentanaOfertaPaz.jpg")
    End If
End Sub
